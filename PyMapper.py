#!/usr/bin/env python3
import re
import os
from getpass import getpass
import pandas as pd
from netmiko import ConnectHandler
from netmiko.ssh_autodetect import SSHDetect
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException
from colorama import Fore, Style, init

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
    }
}

# ------------------ interface normalize (daha esnek) --------------------------
def normalize_interface_names(ifname: str, vendor: str = None) -> str:
    """
    Daha genel bir normalizasyon:
      - Başlangıçtaki harf bloğunu ve sonrasındaki 'num' bloğunu ayırır.
      - Bilinen kısa isimleri uzun forma map eder.
      - Eğer sadece 1/1/1 gibi sayısal format gelirse vendor'a göre prefix ekleyebilir.
    """
    if not ifname or not isinstance(ifname, str):
        return ifname

    s = ifname.strip()

    # common replacements
    s = s.replace('Ethernet ', 'Ethernet').replace('Gi', 'GigabitEthernet').replace('Fa', 'FastEthernet')
    s = s.replace('Te', 'TenGigabitEthernet').replace('Po', 'Port-channel').replace('Port-channel', 'Port-channel')
    # normalize common junos style like ge-0/0/0 -> keep as is but unify separators
    s = s.replace('.', '/')

    # Try match: prefix (letters and punctuation) + numeric part
    m = re.match(r'^(?P<prefix>[A-Za-z\-\_\/]*[A-Za-z]+)?\s*(?P<num>[\d\/\.\:]+.*)$', s)
    if m:
        prefix = (m.group('prefix') or '').strip()
        num = (m.group('num') or '').strip()
        # map some common short names
        mapping = {
            'Gi': 'GigabitEthernet', 'GigabitEthernet': 'GigabitEthernet',
            'Fa': 'FastEthernet', 'FastEthernet': 'FastEthernet',
            'Eth': 'Ethernet', 'Ethernet': 'Ethernet',
            'Te': 'TenGigabitEthernet', 'Port-channel': 'Port-channel',
            'Po': 'Port-channel', 'ae': 'ae', 'ge': 'ge', 'xe': 'xe', 'ae': 'ae'
        }
        # try to find a mapping key that matches the prefix (case-insensitive)
        for k, v in mapping.items():
            if prefix.lower().startswith(k.lower()):
                return f"{v}{num}"
        # default: if prefix empty and num like 1/1/1 -> assume Ethernet on many vendors
        if not prefix and re.match(r'^[\d]+(\/[\d]+)*', num):
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
    print("\n" + "=" * 75 + "\n")

# ------------------ SwitchManager (özetlenmiş) -------------------------------
class SwitchManager:
    def __init__(self, ip, username, password, excel_dosyasi, platform=None, prefer_napalm=True):
        self.ip = ip
        self.username = username
        self.password = password
        self.platform = platform  # netmiko device_type / napalm driver name
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

        # SSHDetect kullanarak auto-detect denemesi (Netmiko). Bazen hatalı olabilir.
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
        # 1) platform tespiti
        platform = self.detect_platform()
        print("Platform in use :", platform)

        # Platform attribute must be fixed from now

        # 2) Eğer NAPALM tercih edildiyse ve driver var ise napalm akışı:
        if self.prefer_napalm and platform:
            napalm_name = None
            # Simple mapping (örn: netmiko 'juniper_junos' -> napalm 'junos')
            if 'juniper' in platform:
                napalm_name = 'junos'
            elif 'cisco_ios' in platform or 'cisco' in platform:
                napalm_name = 'ios'
            elif 'aruba' in platform or 'aoscx' in platform:
                napalm_name = 'aoscx'  # community driver may be needed
            elif 'hp_procurve' in platform or 'procurve' in platform:
                napalm_name = 'procurve'
            if napalm_name:
                try:
                    driver = get_network_driver(napalm_name)    # get network driver aslında sınıf döndürür.
                    optional_args = {}      # driver aslında bir sınıftır. Bu sınıftan device diye bir obje oluşturarak bağlantı kurulur.
                    device = driver(hostname=self.ip, username=self.username, password=self.password, optional_args=optional_args)
                    # device objesi ise bir bağlantı başlatmaya ve parametrelerin verilmesi işine yarıyor.
                    device.open()
                    print("NAPALM in use")
                    facts = device.get_facts()
                    self.hostname = facts.get('hostname', self.ip)
                    # get interfaces + ips + vlans (varsa)
                    interfaces = device.get_interfaces()   # bu da çıktıyı döner "GigabitEthernet0/1": {"is_up": True, "is_enabled": True, "description": "uplink"},
                    interfaces_ip = device.get_interfaces_ip()  # her interface için "GigabitEthernet0/1": {"ipv4": {"192.168.1.1": {"mask": 24}}}, döner mesela
                    vlans = {}
                    # in case if no vlans in vlans attribute
                    try:
                        vlans = device.get_vlans()
                    except Exception:
                        vlans = {}
                    # her interface için interface detayları ve meta datarrı almaya çalışıyoruz.
                    for ifname, meta in interfaces.items():
                        normalized = normalize_interface_names(ifname, vendor=platform)
                        ip_addr = interfaces_ip.get(ifname, {})
                        ip_str = ''
                        # eğer birkaç ip varsa ( yani dict varsa ) bunları ayır.
                        if ip_addr:
                            # ipv4 dict varsa
                            ipv4 = ip_addr.get('ipv4') or {}
                            if ipv4:
                                ip_str = ','.join(list(ipv4.keys()))
                        # eğer meta.get ile is_up alabiliyorsan bir interfaceden bu up olsun direkt. Diğer türlü de down olsun.
                        status = 'up' if meta.get('is_up') else 'down'
                        description = meta.get('description', ' - ')
                        # vlan lookup (napalm get_vlans returns mapping vlan_id -> {name, interfaces})
                        vlan_info = ' - '
                        for vid, vobj in vlans.items():
                            ints = vobj.get('interfaces') or []
                            if ifname in ints or normalized in ints:
                                vlan_info = f"{vid}({vobj.get('name')})"
                                self.vlan_verileri.append({
                                    'Hostname': self.hostname,
                                    'Vlan_id': vid,
                                    'Vlan Name': vobj.get('name'),
                                    'Atanan_portlar': ', '.join(ints)
                                })
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

                    device.close()
                    print("Napalm successful.")
                    return True
                except Exception as e:
                    print("Napalm not successful , trying Netmiko:", e)
                    # fallback Netmiko'ya devam edeceğiz

        # 3) Netmiko flow (fallback / or directly if prefer_napalm is False)
        try:
            print("Netmiko flow is running")
            dev = {"device_type": platform, "host": self.ip,
                   "username": self.username, "password": self.password,
                   "global_delay_factor": 2}
            with ConnectHandler(**dev) as net_connect:
                self.hostname = net_connect.base_prompt
                used_platform = platform or dev['device_type']
                cmds_for = COMMANDS.get(used_platform)

                # 1 - interfaces
                out_if = net_connect.send_command(cmds_for['interfaces'], use_textfsm=True)
                # out_if is often list of dicts when TF exists, else raw str
                parsed_if_list = []     # buraya tek tek cihazdaki tüm interface isimleri gelir
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
                    normalized = normalize_interface_names(port, vendor=used_platform)
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

    def export_to_excel(self):
        print("Writing to excel:", self.excel)
        yeni_arayuz_df = pd.DataFrame(self.toplanan_veriler)
        yeni_vlan_df = pd.DataFrame(self.vlan_verileri)
        cols_ar = ['Hostname', 'Port', 'Description', 'Vlan', 'Status', 'Protocol', 'Ip_address', 'Etherchannel']
        cols_vlan = ['Hostname', 'Vlan_id', 'Vlan Name', 'Atanan_portlar']
        yeni_arayuz_df = yeni_arayuz_df.reindex(columns=cols_ar)
        yeni_vlan_df = yeni_vlan_df.reindex(columns=cols_vlan)
        try:
            mevcut = {}
            try:
                mevcut = pd.read_excel(self.excel, sheet_name=None)
            except FileNotFoundError:
                mevcut = {}
            mevcut_ar = mevcut.get('Interface_Information', pd.DataFrame())
            mevcut_vlan = mevcut.get('VLAN_List', pd.DataFrame())
            merged_ar = pd.concat([mevcut_ar, yeni_arayuz_df], ignore_index=True)
            merged_vlan = pd.concat([mevcut_vlan, yeni_vlan_df], ignore_index=True)
            merged_ar.drop_duplicates(subset=['Hostname', 'Port'], keep='last', inplace=True)
            merged_vlan.drop_duplicates(subset=['Hostname', 'Vlan_id'], keep='last', inplace=True)
            with pd.ExcelWriter(self.excel, engine='openpyxl') as writer:
                merged_ar.to_excel(writer, sheet_name='Interface_Information', index=False)
                merged_vlan.to_excel(writer, sheet_name='VLAN_List', index=False)
            print("Excel up to date.")
        except Exception as e:
            print("Excel writing error:", e)

if __name__ == '__main__':

    excel = 'switch_info.xlsx'

    clear_screen()
    display_banner()
    print("")
    while True:
        platform = None
        platform_input = input("Enter the Device type (ex: ciso, juniper, aruba, hp, autodetect \nFor (q) to quit :").strip() or None
        if "cisco" in platform_input:
            platform = "cisco_ios"
        elif "juniper" in platform_input:
            platform = "juniper_junos"
        elif "aruba" in platform_input:
            platform = "aruba_os"
        elif "hp" in platform_input:
            platform = "hp_procurve"
        elif platform_input == 'q' or platform_input=='Q':
            break
        else:
            print("No switch type is found")
        ip = input("Switch IP/hostname: ")
        username = input("Username: ")
        password = getpass("Password: ")
        prefer_napalm = True # Make it false if you don't want to use Napalm
        SwitchManager(ip, username, password, excel, platform=platform, prefer_napalm=prefer_napalm)
