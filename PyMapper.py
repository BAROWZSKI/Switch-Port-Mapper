#!/usr/bin/env python3
import re
import os
import sys
from getpass import getpass
import time
import serial
import pandas as pd
from netmiko import ConnectHandler
from netmiko.ssh_autodetect import SSHDetect
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException
from colorama import Fore, Style, init
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
# urllib3 for disabling untrusted https traffic and unnecessary warning in cmd.

try:
    from napalm import get_network_driver
    NAPALM_AVAILABLE = True
except Exception:
    NAPALM_AVAILABLE = False

COMMANDS = {
    'cisco_ios': {
        'interfaces': 'show ip interface brief',
        'descriptions': 'show interfaces description',
        'switchport': 'show interfaces switchport',
        'lag': 'show etherchannel summary',
        'vlans': 'show vlan brief'
    },
    'juniper_junos': {
        'interfaces': 'show interfaces terse',
        'descriptions': 'show interfaces descriptions',
        'switchport': 'show ethernet-switching interfaces',
        'lag': 'show lacp interfaces',
        'vlans': 'show vlans brief'
    },
    'aruba_os': {
        'interfaces': 'show interfaces brief',
        'descriptions': 'show interfaces brief',
        'switchport': 'show vlan port',
        'lag': 'show lacp interfaces',
        'vlans': 'show vlan'
    },
    'hp_procurve': {
        'interfaces': 'show interfaces brief',
        'descriptions': 'show interfaces brief',
        'switchport': 'show vlan',
        'lag': 'show lacp info',
        'vlans': 'show vlan'
    },
    'aruba_aoscx': {  
        'interfaces': 'show interface brief',
        'descriptions': 'show interface brief',
        'switchport': 'show vlan',
        'lag': 'show lacp interfaces',
        'vlans': 'show vlan'
    }
}

def clean_output(raw, cmd):
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
    raw = ansi_escape.sub('', raw)
    lines = raw.splitlines()
    cleaned = []
    for line in lines:
        if cmd.strip() in line:
            continue
        # Prompt içeren satırları tamamen sil (tek veya çift prompt)
        if re.search(r'\S+[>#]\s*$', line):
            continue
        cleaned.append(line)
    return '\n'.join(cleaned)

def normalize_interface_names(ifname: str, vendor: str = None) -> str:

    if not ifname or not isinstance(ifname, str):
        return ifname

    s = ifname.strip()

    # common replacements
    s = s.replace('Ethernet ', 'Ethernet').replace('Gi', 'GigabitEthernet').replace('Fa', 'FastEthernet')
    s = s.replace('Te', 'TenGigabitEthernet').replace('Po', 'Port-channel').replace('Port-channel', 'Port-channel')
    # normalize common junos style like ge-0/0/0 -> keep as is but unify separators
    s = s.replace('.', '/')

    # Try match: prefix (letters and punctuation) + numeric part regex kodun formatını standartlaştırır.
    m = re.match(r'^(?P<prefix>[A-Za-z/_-]*[A-Za-z]+)?\s*(?P<num>[\d/.:]+.*)$', s)
    if m:
        prefix = (m.group('prefix') or '').strip()
        num = (m.group('num') or '').strip()
        # map some common short names
        mapping = {
            'Gi': 'GigabitEthernet', 'GigabitEthernet': 'GigabitEthernet',
            'Fa': 'FastEthernet', 'FastEthernet': 'FastEthernet',
            'Eth': 'Ethernet', 'Ethernet': 'Ethernet',
            'Te': 'TenGigabitEthernet', 'Port-channel': 'Port-channel',
            'Po': 'Port-channel', 'ae': 'ae', 'ge': 'ge', 'xe': 'xe',
            'lo': 'Loopback','loopback': 'Loopback','l': 'Loopback','loop': 'Loopback',
        }
        # try to find a mapping key that matches the prefix (case-insensitive)
        for k, v in mapping.items():
            if prefix.lower().startswith(k.lower()):
                return f"{v}{num}"
        # default: if prefix empty and num like 1/1/1 -> assume Ethernet on many vendors
        if not prefix and re.match(r'^\d+(\/\d+)*', num):
            # vendor-specific default can be improved
            if vendor and 'juniper' in vendor:
                return f"ge-{num}"
            else:
                return f"Ethernet{num}"
        # fallback: return cleaned original
        return f"{prefix}{num}" if prefix else s

    return s

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def display_banner():
    init(autoreset=True)
    banner_art = """
        ██████╗ ██╗   ██╗███╗   ███╗ █████╗ ██████╗ ██████╗ ███████╗██████╗ 
        ██╔══██╗╚██╗ ██╔╝████╗ ████║██╔══██╗██╔══██╗██╔══██╗██╔════╝██╔══██╗
        ██████╔╝ ╚████╔╝ ██╔████╔██║███████║██████╔╝██████╔╝█████╗  ██████╔╝
        ██╔═══╝   ╚██╔╝  ██║╚██╔╝██║██╔══██║██╔═══╝ ██╔═══╝ ██╔══╝  ██╔══██╗
        ██║        ██║   ██║ ╚═╝ ██║██║  ██║██║     ██║     ███████╗██║  ██║
        ╚═╝        ╚═╝   ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝     ╚══════╝╚═╝  ╚═╝                                                               
    """
    banner_color = Fore.CYAN + Style.BRIGHT
    print(banner_color + banner_art)
    tagline = ' '*13 + "< Switch port configuration mapper with ssh >"
    tagline_color = Fore.YELLOW + Style.BRIGHT

    print(tagline_color + tagline)
    print("=" * 75)

class SwitchManager:
    def __init__(self, ip, username, password, excel_dosyasi, platform=None, prefer_napalm=True, conn_type="ssh", secret="", secret_user = ""):
        self.ip = ip
        self.username = username
        self.password = password
        self.platform = platform
        self.conn_type = conn_type
        self.secret = secret
        self.secret_user = secret_user
        
        # If connection is serial than napalm is useless
        if self.conn_type == "serial":
            self.prefer_napalm = False
        else:
            self.prefer_napalm = prefer_napalm and NAPALM_AVAILABLE
        self.excel = excel_dosyasi

        self.toplanan_veriler = []
        self.vlan_verileri = []
        self.hostname = 'unknown'

        ok = self.run_collection()
        if ok and self.toplanan_veriler:
            self.export_to_excel()
        elif ok and not self.toplanan_veriler:
            print("-> Connection established but no interface data found.")
        else:
            print("-> Connection couldn't establilshed .")



    # Tries SSHDetect to detect platform
    def detect_platform(self):
        # If platform given
        if self.platform:
            return self.platform

        if self.conn_type == "serial":
            print("Can't run autodetect when connection is serial")
            return None

        # Using SSHDetect auto-detect in Netmiko we are trying to detect the vendor  
        device = {"device_type": "autodetect", "host": self.ip, "username": self.username, "password": self.password}
        try:
            guesser = SSHDetect(**device)
            best_match = guesser.autodetect()   # burada best_match cisco_ios juniper_junos gibi değerler döndürür platform belirlenir.
            print(f"SSHDetect result: {best_match}")
            self.platform = best_match
            return best_match
        except Exception as e:
            print("Platform couldn't identified:", e)
            return None

    def run_collection(self):
        # 1) Manually checking the platform with detect_platform function
        platform = self.platform or self.detect_platform()
        print("Platform in use :", platform)

        # Platform attribute must be fixed from now

        # NAPALM FLOW
        if self.prefer_napalm and platform:
            napalm_name = None
            # Simple mapping 
            if 'juniper' in platform:
                napalm_name = 'junos'
            elif 'cisco_ios' in platform or 'cisco' in platform:
                napalm_name = 'ios'
            elif 'aruba' in platform or 'aruba_aoscx' in platform:
                napalm_name = 'aoscx'  # community driver may be needed
            elif 'hp_procurve' in platform or 'procurve' in platform:
                napalm_name = 'procurve'
            if napalm_name:
                try:
                    driver = get_network_driver(napalm_name)    # get network driver returns object.
                    optional_args = {}      # Driver is actually just a class. We create device object from that class.
                    if self.secret:
                        optional_args['secret'] = self.secret
                    device = driver(hostname=self.ip, username=self.username, password=self.password, optional_args=optional_args)
                    device.open()
                    print("NAPALM in use")
                    facts = device.get_facts()
                    self.hostname = facts.get('hostname', self.ip)
                    # get interfaces + ips + vlans (if exist)
                    interfaces = device.get_interfaces()   # "GigabitEthernet0/1": {"is_up": True, "is_enabled": True, "description": "uplink"},
                    interfaces_ip = device.get_interfaces_ip()  # for every interface  "GigabitEthernet0/1": {"ipv4": {"192.168.1.1": {"mask": 24}}}, döner mesela
                    vlans = {}
                    # in case if no vlans in vlans attribute
                    try:
                        vlans = device.get_vlans()
                    except Exception:
                        vlans = {}
                    # First loops gathers only vlan values
                    for ifname, meta in interfaces.items():
                        normalized = normalize_interface_names(ifname, vendor=platform)
                        ip_addr = interfaces_ip.get(ifname, {})
                        ip_str = ''
                        # if there is multiple ip (type dict) seperates.
                        if ip_addr:
                            # if ipv4 dict 
                            ipv4 = ip_addr.get('ipv4') or {}
                            if ipv4:
                                ip_str = ','.join(list(ipv4.keys()))

                        status = 'up' if meta.get('is_up') else 'down'
                        description = meta.get('description', ' - ')
                        # vlan lookup (napalm get_vlans returns mapping vlan_id -> {name, interfaces})
                        vlan_info = ' - '
                        for vid, vobj in vlans.items():
                            ints = vobj.get('interfaces') or []
                            if ifname in ints or normalized in ints:
                                vlan_info = f"{vid}({vobj.get('name')})"
                                break
                            
                        self.toplanan_veriler.append({
                            "Hostname": self.hostname,
                            "Port": normalized,
                            "Status": status,
                            "Protocol": meta.get('is_enabled', ' - '),
                            "Ip_address": ip_str or ' - ',
                            "Vlan": vlan_info,
                            "Description": description,
                            "Etherchannel": ' - '
                        })

                    for vid, vobj in vlans.items():
                        ints = vobj.get('interfaces') or []
                        self.vlan_verileri.append({
                            'Hostname': self.hostname,
                            'Vlan_id': vid,
                            'Vlan Name': vobj.get('name'),
                            'Atanan_portlar': ', '.join(ints)
                        })

                    device.close()
                    print("Napalm successful.")
                    return True
                except Exception as e:
                    print("Napalm not successful , trying Netmiko:", e)

        # 3) Netmiko flow (fallback / or directly if prefer_napalm is False)
        try:
            print("Netmiko flow is running")
            
            # _serial takısını konsol bağlantıları için mecburen ekliyoruz
            used_platform = platform
            if self.conn_type == "serial" and platform:
                if not platform.endswith("_serial"):
                    used_platform = platform + "_serial"
            
            if self.conn_type == "serial":
                base_platform = platform.replace("_serial", "") if "_serial" in platform else platform
                return self._serial_collect(base_platform)
            else:
                dev = {
                    "device_type": platform,              # Örn: hp_procurve (SSH için)
                    "host": self.ip,                      # SSH için IP adresi
                    "username": self.username,
                    "password": self.password,
                    "secret": self.secret,
                    "global_delay_factor": 2
                    }
        
            with ConnectHandler(**dev) as net_connect:
                if self.conn_type == "serial":
                    net_connect.write_channel("\r\n")
                    time.sleep(2)
                    output = net_connect.read_channel()
                        
                    # send credential if login requires
                    if "username" in output.lower() or "login" in output.lower():
                        net_connect.write_channel(self.username + "\n")
                        time.sleep(1)
                        output = net_connect.read_channel()
                    
                    if "password" in output.lower():
                        net_connect.write_channel(self.password + "\n")
                        time.sleep(2)
                        output = net_connect.read_channel()
                    
                    from netmiko import redispatch
                    redispatch(net_connect, device_type=base_platform)  # örn: "hp_procurve"
                    
                if not net_connect.check_enable_mode():
                    print("Enable mode activating...")
                    if self.secret_user:
                        output = net_connect.send_command_timing("enable")
                        if "sername" in output.lower() or "login" in output.lower():
                            output += net_connect.send_command_timing(self.secret_user)
                        if "ssword" in output.lower() or "word:" in output.lower():
                            net_connect.send_command_timing(self.secret)
                    else:
                        net_connect.enable()
                
                self.hostname = net_connect.base_prompt
                
                # _serial uzantısı COMMANDS sözlüğünde olmadığı için komut ararken onu siliyoruz
                base_platform = used_platform.replace("_serial", "") if used_platform else dev.get('device_type', '').replace("_serial", "")
                cmds_for = COMMANDS.get(base_platform)

                if not cmds_for:
                    print(f"There is no command set for {base_platform}")
                    return False

                out_if = net_connect.send_command(cmds_for['interfaces'], use_textfsm=True)

                # out_if is often list of dicts when TF exists, else raw str
                parsed_if_list = []     # All interface names comes here
                if isinstance(out_if, list):
                    parsed_if_list = out_if
                else:
                    # fallback rudimentary parse
                    for line in str(out_if).splitlines():
                        if line.strip() and not line.lower().startswith('interface'):
                            # brute force: split by whitespace
                            parts = line.split()
                            if len(parts) >= 1:
                                parsed_if_list.append({"interface": parts[0], "ip_address": parts[1] if len(parts) > 1 else ' - '})
                # descriptions
                out_desc = net_connect.send_command(cmds_for['descriptions'], use_textfsm=True)
                parsed_desc = out_desc if isinstance(out_desc, list) else []
                # switchport/vlan
                out_sw = net_connect.send_command(cmds_for['switchport'], use_textfsm=True) if 'switchport' in cmds_for else []
                parsed_sw = out_sw if isinstance(out_sw, list) else []
                # lag
                out_lag = None
                if 'lag' in cmds_for:
                    try:
                        out_lag = net_connect.send_command(cmds_for['lag'], use_textfsm=True)
                    except Exception:
                        out_lag = None
                # vlan brief
                out_vlans = net_connect.send_command(cmds_for['vlans'], use_textfsm=True) if 'vlans' in cmds_for else []
                parsed_vlans = out_vlans if isinstance(out_vlans, list) else []

                # Build interface_details dict (vendor-agnostic)
                interface_details = {}
                for iface in parsed_if_list:
                    # keys differ between templates; try common ones
                    port = iface.get('interface') or iface.get('port') or iface.get('intf') or iface.get('name')
                    if not port: continue
                    normalized = normalize_interface_names(port, vendor=platform)
                    interface_details[normalized] = {
                        "ip_address": iface.get('ip_address') or iface.get('ip') or ' - ',
                        "status": iface.get('status') or iface.get('oper') or ' - ',
                        "protocol": iface.get('proto') or iface.get('protocol') or ' - ',
                        "description": " - ",
                        "vlan": " - ",
                        "etherchannel": " - "
                    }
                # fill description
                for d in parsed_desc:
                    port = d.get('port') or d.get('interface') or d.get('name')
                    if not port: continue
                    normalized = normalize_interface_names(port, vendor=used_platform)
                    if normalized in interface_details:
                        # try a few keys for description
                        interface_details[normalized]['description'] = d.get('description') or d.get('desc') or interface_details[normalized]['description']

                # fill switchport/vlan info
                for sw in parsed_sw:
                    port = sw.get('interface') or sw.get('port') or sw.get('name')
                    if not port: continue
                    normalized = normalize_interface_names(port, vendor=used_platform)
                    if normalized in interface_details:
                        mode = sw.get('mode') or ''
                        if 'access' in mode.lower():
                            interface_details[normalized]['vlan'] = f"Access({sw.get('access_vlan','')})"
                        elif 'trunk' in mode.lower():
                            interface_details[normalized]['vlan'] = f"Trunk({sw.get('trunk_vlans','')})"
                        else:
                            # hp/aruba templates may have different keys - try a few
                            if sw.get('vlan'):
                                interface_details[normalized]['vlan'] = sw.get('vlan')

                # fill etherchannel info roughly from out_lag if present
                if out_lag and isinstance(out_lag, list):
                    for g in out_lag:
                        # try typical keys
                        bundle = g.get('bundle_name') or g.get('group') or g.get('lag')
                        members = g.get('member_interface') or g.get('members') or []
                        for m in members:
                            nm = normalize_interface_names(m, vendor=used_platform)
                            if nm in interface_details:
                                interface_details[nm]['etherchannel'] = bundle

                for v in parsed_vlans:
                    # many templates provide 'interfaces' list or 'ports'
                    ports = v.get('interfaces') or v.get('ports') or v.get('assigned_ports') or []
                    for p in ports:
                        nm = normalize_interface_names(p, vendor=used_platform)
                        # reverse-lookup interface to add vlan string (simple)
                        if nm in interface_details:
                            interface_details[nm]['vlan'] = f"{v.get('vlan_id') or v.get('vlan') or v.get('vlan_id','') }({v.get('vlan_name') or v.get('name','')})"

                # convert to top-level list
                for port, det in interface_details.items():
                    self.toplanan_veriler.append({
                        "Hostname": self.hostname,
                        "Port": port,
                        "Status": det["status"],
                        "Protocol": det["protocol"],
                        "Ip_address": det["ip_address"],
                        "Vlan": det["vlan"],
                        "Description": det["description"],
                        "Etherchannel": det["etherchannel"]
                    })
                # VLAN sheet
                for v in parsed_vlans:
                    ports = v.get('interfaces') or v.get('ports') or []
                    norm_ports = [normalize_interface_names(p, vendor=used_platform) for p in ports]
                    self.vlan_verileri.append({
                        'Hostname': self.hostname,
                        'Vlan_id': v.get('vlan_id') or v.get('vlan') or v.get('id'),
                        'Vlan Name': v.get('vlan_name') or v.get('name'),
                        'Atanan_portlar': ', '.join(norm_ports)
                    })
                return True

        except NetmikoTimeoutException:
            print("Zaman aşımı:", self.ip)
            return False
        except NetmikoAuthenticationException:
            print("Auth error:", self.ip)
            return False
        except Exception as e:
            print("Unexpected error:", e)
            return False

    def _serial_collect(self, base_platform):
        from ntc_templates.parse import parse_output

        try:
            ser = serial.Serial(port=self.ip, baudrate=9600, timeout=3)

            # send function tries to wake switch up
            def send(cmd, wait=8):
                ser.write((cmd + "\r\n").encode())
                time.sleep(2)
                output = ""
                deadline = time.time() + wait
                while time.time() < deadline:
                    if ser.in_waiting:
                        output += ser.read(ser.in_waiting).decode(errors='ignore')
                        if any(output.strip().endswith(p) for p in ('#', '>')):
                            break
                    time.sleep(1)
                return output

            # ── Login flow from now ── #
            output = send("", wait=3)
            print(f"1) WAKE: {repr(output)}")

            if "press any key" in output.lower():
                output = send("", wait=3)
                print(f"2) PRESS ANY KEY: {repr(output)}")

            if "username" in output.lower() or "login" in output.lower():
                output = send(self.username, wait=3)
                print(f"3) USERNAME SENT: {repr(output)}")

            if "password" in output.lower():
                output = send(self.password, wait=3)
                print(f"4) PASSWORD SENT: {repr(output)}")

            output = send("enable", wait=3)
            print(f"5) ENABLE: {repr(output)}")

            # Switch enable için username+password istiyor
            if "username" in output.lower() or "sername" in output.lower():
                enable_user = self.secret_user if self.secret_user else self.username
                output = send(enable_user, wait=3)
                print(f"5b) ENABLE USER: {repr(output)}")

                # Username sonrası password bekliyorsa gönder
                if "password" in output.lower():
                    enable_pw = self.secret if self.secret else self.password
                    output = send(enable_pw, wait=3)
                    print(f"5c) ENABLE PW AFTER USER: {repr(output)}")

            elif "password" in output.lower():
                # Direkt password sorduysa (username sormadan)
                enable_pw = self.secret if self.secret else self.password
                output = send(enable_pw, wait=3)
                print(f"6) ENABLE PW: {repr(output)}")

            output = send("no page", wait=3)
            print(f"7) NO PAGE: {repr(output)}")

            # Hostname: prompt'un son satırından al (ANSI temizlenmiş hali)
            clean_last = re.sub(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])', '', output)
            last_line = clean_last.strip().splitlines()[-1] if clean_last.strip() else ""
            self.hostname = last_line.replace("#", "").replace(">", "").strip() or "console_device"
            print(f"HOSTNAME: {self.hostname}")

            # ── COMMANDS dict'ten komutları gönder ve parse et ───────
            cmds_for = COMMANDS.get(base_platform)
            if not cmds_for:
                print(f"No command set for {base_platform}")
                ser.close()
                return False

            raw_outputs = {}
            for key, cmd in cmds_for.items():
                raw = send(cmd, wait=5)
                raw = clean_output(raw, cmd)
                print(f"\n{'='*50}")
                print(f"CMD: {cmd}")
                print(f"RAW OUTPUT:\n{repr(raw[:500])}")
                # ──────────────────
                
                
                try:
                    parsed = parse_output(platform=base_platform, command=cmd, data=raw)
                    raw_outputs[key] = parsed if isinstance(parsed, list) else []
                except Exception as e:
                    print(f"TextFSM parse failed for '{cmd}': {e}")
                    raw_outputs[key] = []

            ser.close()

            # ── SSH flow ile AYNI parse mantığı ─────────────────────
            parsed_if_list = raw_outputs.get('interfaces', [])
            parsed_desc    = raw_outputs.get('descriptions', [])
            parsed_sw      = raw_outputs.get('switchport', [])
            parsed_vlans   = raw_outputs.get('vlans', [])
            out_lag        = raw_outputs.get('lag', [])

            interface_details = {}

            for iface in parsed_if_list:
                port = iface.get('interface') or iface.get('port') or iface.get('intf') or iface.get('name')
                if not port:
                    continue
                normalized = normalize_interface_names(port, vendor=base_platform)
                interface_details[normalized] = {
                    "ip_address":   iface.get('ip_address') or iface.get('ip') or ' - ',
                    "status":       iface.get('status') or iface.get('oper') or ' - ',
                    "protocol":     iface.get('proto') or iface.get('protocol') or ' - ',
                    "description":  " - ",
                    "vlan":         " - ",
                    "etherchannel": " - "
                }

            for d in parsed_desc:
                port = d.get('port') or d.get('interface') or d.get('name')
                if not port:
                    continue
                normalized = normalize_interface_names(port, vendor=base_platform)
                if normalized in interface_details:
                    interface_details[normalized]['description'] = (
                        d.get('description') or d.get('desc') or interface_details[normalized]['description']
                    )

            for sw in parsed_sw:
                port = sw.get('interface') or sw.get('port') or sw.get('name')
                if not port:
                    continue
                normalized = normalize_interface_names(port, vendor=base_platform)
                if normalized in interface_details:
                    mode = sw.get('mode') or ''
                    if 'access' in mode.lower():
                        interface_details[normalized]['vlan'] = f"Access({sw.get('access_vlan', '')})"
                    elif 'trunk' in mode.lower():
                        interface_details[normalized]['vlan'] = f"Trunk({sw.get('trunk_vlans', '')})"
                    else:
                        if sw.get('vlan'):
                            interface_details[normalized]['vlan'] = sw.get('vlan')

            if isinstance(out_lag, list):
                for g in out_lag:
                    bundle  = g.get('bundle_name') or g.get('group') or g.get('lag')
                    members = g.get('member_interface') or g.get('members') or []
                    for m in members:
                        nm = normalize_interface_names(m, vendor=base_platform)
                        if nm in interface_details:
                            interface_details[nm]['etherchannel'] = bundle

            for v in parsed_vlans:
                ports = v.get('interfaces') or v.get('ports') or v.get('assigned_ports') or []
                for p in ports:
                    nm = normalize_interface_names(p, vendor=base_platform)
                    if nm in interface_details:
                        interface_details[nm]['vlan'] = (
                            f"{v.get('vlan_id') or v.get('vlan', '')}({v.get('vlan_name') or v.get('name', '')})"
                        )

            # toplanan_veriler'e ekle
            for port, det in interface_details.items():
                self.toplanan_veriler.append({
                    "Hostname":     self.hostname,
                    "Port":         port,
                    "Status":       det["status"],
                    "Protocol":     det["protocol"],
                    "Ip_address":   det["ip_address"],
                    "Vlan":         det["vlan"],
                    "Description":  det["description"],
                    "Etherchannel": det["etherchannel"]
                })

            # VLAN sheet
            for v in parsed_vlans:
                ports = v.get('interfaces') or v.get('ports') or []
                norm_ports = [normalize_interface_names(p, vendor=base_platform) for p in ports]
                self.vlan_verileri.append({
                    'Hostname':      self.hostname,
                    'Vlan_id':       v.get('vlan_id') or v.get('vlan') or v.get('id'),
                    'Vlan Name':     v.get('vlan_name') or v.get('name'),
                    'Atanan_portlar': ', '.join(norm_ports)
                })

            return True

        except Exception as e:
            print(f"Serial connection error: {e}")
        return False

    def export_to_excel(self):

        print("Writing to excel:", self.excel)
        yeni_arayuz_df = pd.DataFrame(self.toplanan_veriler)
        yeni_vlan_df = pd.DataFrame(self.vlan_verileri)
        cols_ar = ['Hostname', 'Port', 'Description', 'Vlan', 'Status', 'Protocol', 'Ip_address', 'Etherchannel']
        cols_vlan = ['Hostname', 'Vlan_id', 'Vlan Name', 'Atanan_portlar']
        yeni_arayuz_df = yeni_arayuz_df.reindex(columns=cols_ar)
        yeni_vlan_df = yeni_vlan_df.reindex(columns=cols_vlan)

        # If file not exist , then create one
        if not os.path.exists(self.excel):
            try:
                with pd.ExcelWriter(self.excel, engine='openpyxl', mode='w') as writer:
                    yeni_arayuz_df.to_excel(writer, sheet_name='Interface_Information', index=False)
                    yeni_vlan_df.to_excel(writer, sheet_name='VLAN_List', index=False)
                print("Excel file created and data written.")
                return
            except Exception as e:
                print("Error creating excel file:", e)
                return

        # If file exist , read the related pages, write down it, clear the duplicate, then replace the pages.
        try:
            mevcut_ar = pd.DataFrame(columns=cols_ar)
            mevcut_vlan = pd.DataFrame(columns=cols_vlan)

            try:
                mevcut_ar = pd.read_excel(self.excel, sheet_name='Interface_Information')
            except Exception:
                # sheet yok veya okunamadı -> boş bırak
                mevcut_ar = pd.DataFrame(columns=cols_ar)

            try:
                mevcut_vlan = pd.read_excel(self.excel, sheet_name='VLAN_List')
            except Exception:
                mevcut_vlan = pd.DataFrame(columns=cols_vlan)

            # add the datas with conocat
            merged_ar = pd.concat([mevcut_ar, yeni_arayuz_df], ignore_index=True)
            merged_vlan = pd.concat([mevcut_vlan, yeni_vlan_df], ignore_index=True)

            # clear duplicates
            if not merged_ar.empty:
                merged_ar.drop_duplicates(subset=['Hostname', 'Port'], keep='last', inplace=True)
            if not merged_vlan.empty:
                merged_vlan.drop_duplicates(subset=['Hostname', 'Vlan_id'], keep='last', inplace=True)

            # Pandas >= 1.3: if_sheet_exists param available; mode='a' with if_sheet_exists='replace' replaces target sheets
            with pd.ExcelWriter(self.excel, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                merged_ar.to_excel(writer, sheet_name='Interface_Information', index=False)
                merged_vlan.to_excel(writer, sheet_name='VLAN_List', index=False)

            print("Excel updated (sheets replaced).")
        except Exception as e:
            print("Excel writing error:", e)

if __name__ == '__main__':
    excel = 'switch_info.xlsx' # Dosya adı
    clear_screen()
    display_banner()
    print("")

    try:
        while True:
            
            print("Select your connection type (1/2)")
            connection_input = input("1)SSH\n2)Console\n").strip()
            
            conn_type = "ssh" if connection_input == "1" else "serial" 
            
            platform = None
            print("\n--- Platform Selection ---")
            platform_input = input("Enter the Device type (ex: \n*cicso, \n*juniper, \n*aruba(if aruba edge sw),\n*arubaBb(if Aruba BB sw), \n*hp, \n*autodetect \nFor (q) to quit :").strip() or None
                        
            if not platform_input:
                continue

            if "cisco" in platform_input:
                platform = "cisco_ios"
            elif "juniper" in platform_input:
                platform = "juniper_junos"
            elif "arubaBb" in platform_input: 
                platform = "aruba_aoscx"
            elif "aruba" in platform_input:
                platform = "hp_procurve"
            elif "hp" in platform_input:
                platform = "hp_procurve"
            elif platform_input.lower() in ['q', 'quit', 'exit']:
                print("Exiting program...")
                break
            else:
                print("No switch type found or invalid input.")
                continue

            if conn_type == "serial":
                ip = input("Enter you port (COM3, /dev/ttyUSB0) : ").strip()
            else:
                ip = input("Switch IP/hostname: ").strip()                
            username = input("Username: ")
            password = getpass("Password: ")
            
            secret_user = input("Enable Username (Ana username ile aynıysa veya yoksa Enter'a bas): ").strip()
            secret = getpass("Enable Password (Yoksa Enter'a bas): ")
            
            prefer_napalm = False 
            SwitchManager(ip, username, password, excel, platform=platform, prefer_napalm=prefer_napalm, conn_type=conn_type, secret=secret, secret_user=secret_user)
            again = input("Add another switch (Y/n): ").strip().lower()
            if again != 'y':
                print("Shutting down the programme.")
                break

    except KeyboardInterrupt:
        print(Fore.RED + "\n\n[!] Process cancelled by user (Ctrl+C). Ending program..." + Style.RESET_ALL)
        sys.exit() 
    except Exception as e:
        print(f"\nUnexpected error occured: {e}")
