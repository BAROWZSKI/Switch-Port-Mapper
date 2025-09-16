#!/usr/bin/env python3
# normalize_portun metod değil fonksiyon olduğu hal

import pandas as pd
from netmiko import ConnectHandler
from getpass import getpass
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException

import os

class SwitchManager:
    def __init__(self,ip , username, password, excel_dosyasi):
        self.device_info = {
            'device_type': 'cisco_ios',
            'ip': ip,
            'username': username,
            'password': password,
            'global_delay_factor': 2,
        }
        self.excel_dosyasi = excel_dosyasi  # export to excel buna yazar.
        self.toplanan_veriler = []
        self.vlan_verileri = []
        self.hostname = 'unknown'

        bilgi = self.connect_and_fetch()
        if bilgi and self.toplanan_veriler:
            # tek başarılı olunan koşulda excele yaz
            self.export_to_excel()
        elif bilgi and not self.toplanan_veriler:
            print(" -> Baglanti basarili ama okunacak arayüz bilgisi bulunamadi.")
        else:
            print(f" -> {self.device_info['ip']} için işlem basarisiz oldu, Excel'e yazilamadi.")

    def connect_and_fetch(self):
        try:
            print(f"\n-> {self.device_info['ip']} adresine baglaniyor...")
            with ConnectHandler(**self.device_info) as net_connect:
                # net_connect objesinin hostname döndüren metodu
                self.hostname = net_connect.base_prompt
                print(f"-> baglanti basarili: {self.hostname}'e baglanildi")

                interface_details = {}

                # --- 1. IP Interface bilgisi ---
                interfaces = net_connect.send_command("show ip interface brief", use_textfsm=True)
                print(interfaces)
                print("1 basladi")
                if not interfaces or not isinstance(interfaces, list):
                    print("HATA: 'show ip interface brief' ciktisi alinamadi.")
                    return False

                for iface in interfaces:
                    port = iface["interface"]
                    normalized_port = normalize_interface_names(port)
                    interface_details[normalized_port] = {
                        "ip_address": iface.get("ip_address", " - "),
                        "status": iface.get("status", " - "),
                        "protocol": iface.get("proto", " - "),
                        "description": " - ",
                        "vlan": " - ",
                        "etherchannel": " - "
                    }
                print("1 de sorun yok")

                # --- 2. Description bilgisi ---
                descriptions = net_connect.send_command("show interfaces description", use_textfsm=True)
                print(descriptions)
                print("2 basladi")

                if isinstance(descriptions, list):
                    for desc in descriptions:
                        short_port = desc.get("port")  # 'Gi1/0/1' gibi kısa formatı alır
                        if not short_port: continue # eğer port ismi boşsa direkt atlasın normalize etmesin
                        normalized_port = normalize_interface_names(short_port)
                        if normalized_port in interface_details:
                            interface_details[normalized_port]["description"] = desc.get("description", "")


                print("2 de sorun yok")

                # --- 3. VLAN bilgisi ---
                switchports = net_connect.send_command("show interfaces switchport", use_textfsm=True)
                print("3 basladi")
                print(switchports)
                if switchports and isinstance(switchports, list):
                    for sw in switchports:
                        short_port = sw["interface"]
                        if not short_port: continue
                        normalized_port = normalize_interface_names(short_port)
                        if normalized_port in interface_details:
                            mode = sw.get("mode", "")
                            if "access" in mode:
                                vlan_info = f"Access({sw.get('access_vlan', '')})"
                            elif "trunk" in mode:
                                vlan_info = f"Trunk({sw.get('trunk_vlans', '')})"
                            else:
                                vlan_info = " - "
                            interface_details[normalized_port]["vlan"] = vlan_info
                print("3 de sorun yok")

                # --- 4. EtherChannel bilgisi ---
                etherchannels = net_connect.send_command("show etherchannel summary", use_textfsm=True)
                print("4 basladi")
                print(etherchannels)

                if isinstance(etherchannels, list):
                    for ch_group in etherchannels:
                        bundle_name = ch_group.get("bundle_name")   # boyle sadece portchannel ismini alirsin.
                        if not bundle_name: continue    #  boşsa devam et

                        # Bu gruba üye olan tüm arayüzlerin listesini alıyoruz
                        member_interfaces = ch_group.get("member_interface", [])
                        for short_member_port in member_interfaces:
                            normalized_member = normalize_interface_names(short_member_port)
                            if normalized_member in interface_details:
                                interface_details[normalized_member]["etherchannel"] = bundle_name

                print("4 bitti hepsi listeye yaziliyor.")


                # --- 5. Vlanları al ---
                vlan_database = net_connect.send_command("show vlan brief", use_textfsm=True)
                print("5 basladi")
                if isinstance(vlan_database, list):
                    for vlan_entry in vlan_database:
                        # 1. Önce portların kısa isimli listesini alalım
                        short_ports_list = vlan_entry.get("interfaces", [])
                        normalized_ports_list = []
                        for normal_port in short_ports_list:
                            normalized_ports_list.append(normalize_interface_names(normal_port))

                        final_ports_string = ", ".join(normalized_ports_list)

                        self.vlan_verileri.append({
                            'Hostname': self.hostname,
                            'Vlan_id': vlan_entry.get('vlan_id'),
                            'Vlan_adi': vlan_entry.get('vlan_name'),
                            'Atanan_portlar': final_ports_string
                        })
                print("5 bitti (VLAN Veritabanı)")

                # Sonuçları listeye dönüştür
                for port, details in interface_details.items():
                    self.toplanan_veriler.append({
                        "Hostname": self.hostname,
                        "Port": port,
                        "Status": details["status"],
                        "Protocol": details["protocol"],
                        "Ip_address": details["ip_address"],
                        "Vlan": details["vlan"],
                        "Description": details["description"],
                        "Etherchannel": details["etherchannel"]
                    })

                print(f"-> {len(self.toplanan_veriler)} interface kayidi kopyalandi.")
                return True

        except NetmikoTimeoutException:
            print(f"HATA: {self.device_info['ip']} baglanti zaman asimi.")
            return False
        except NetmikoAuthenticationException:
            print(f"HATA: {self.device_info['ip']} kullanici adi sifre hatali.")
            return False
        except Exception as e:
            print(f"HATA: Beklenmedik hata: {e}")
            return False

    def export_to_excel(self):
        print(f"-> Veriler '{self.excel_dosyasi}' dosyasına yazılıyor...")

        # 1. Adım: iki ayrı excel sheeti için iki farklı data frame oluşturulur.
        yeni_arayuz_df = pd.DataFrame(self.toplanan_veriler)
        yeni_vlan_df = pd.DataFrame(self.vlan_verileri)

        # 2. Adım: connect_and_fetchde show vlan brief için eklenen
        sutun_sirasi_arayuz = [
            'Hostname', 'Port', 'Description', 'Vlan', 'Status',
            'Protocol', 'Ip_address', 'Etherchannel'
        ]
        sutun_sirasi_vlan = [
            'Hostname', 'Vlan_id', 'Vlan_adi', 'Atanan_portlar'
        ]

        # Sütunları yeniden sırala
        yeni_arayuz_df = yeni_arayuz_df.reindex(columns=sutun_sirasi_arayuz)
        yeni_vlan_df = yeni_vlan_df.reindex(columns=sutun_sirasi_vlan)

        try:
            # 3. Adım: Mevcut Excel dosyasını (eğer varsa) oku
            try:
                mevcut_sheets = pd.read_excel(self.excel_dosyasi, sheet_name=None)
            except FileNotFoundError:
                mevcut_sheets = {}

            # 4. Adım: Her sayfa için eski ve yeni veriyi birleştir
            mevcut_arayuz_df = mevcut_sheets.get('Arayuz_Bilgileri', pd.DataFrame())
            mevcut_vlan_df = mevcut_sheets.get('VLAN_Listesi', pd.DataFrame())

            birlesmis_arayuz_df = pd.concat([mevcut_arayuz_df, yeni_arayuz_df], ignore_index=True)
            birlesmis_vlan_df = pd.concat([mevcut_vlan_df, yeni_vlan_df], ignore_index=True)

            # Yinelenen satırları kaldır (en son eklenen kalır)
            birlesmis_arayuz_df.drop_duplicates(subset=['Hostname', 'Port'], keep='last', inplace=True)
            birlesmis_vlan_df.drop_duplicates(subset=['Hostname', 'Vlan_id'], keep='last', inplace=True)

            # 5. Adım: ExcelWriter kullanarak iki sayfayı da aynı anda dosyaya yaz
            with pd.ExcelWriter(self.excel_dosyasi, engine='openpyxl') as writer:
                birlesmis_arayuz_df.to_excel(writer, sheet_name='Arayuz_Bilgileri', index=False)
                birlesmis_vlan_df.to_excel(writer, sheet_name='VLAN_Listesi', index=False)

            print(f"-> '{self.excel_dosyasi}' dosyası iki sayfa ile başarıyla güncellendi.")

        except Exception as e:
            print(f"HATA: Excel dosyasına yazarken bir hata oluştu: {e}")

def normalize_interface_names(non_norm_int):
    interfaces = [
        [["Ethernet", "Eth"], "Ethernet"],
        [["FastEthernet", " FastEthernet", "Fa", "interface FastEthernet"], "FastEthernet"],
        [["GigabitEthernet", "Gi", " GigabitEthernet", "interface GigabitEthernet"], "GigabitEthernet"],
        [["TenGigabitEthernet", "Te"], "TenGigabitEthernet"],
        [["Port-channel", "Po"], "Port-channel"],
        [["Serial","Ser"], "Serial"],
        [["Loopback","loopback","lo","Lo"],"Loopback"],
        [["Vlan","Vl"],"Vlan"]
    ]
    try:
        num_index = non_norm_int.index(next(x for x in non_norm_int if x.isdigit()))    # kacıncı indexte rakam basliyor
        interface_type = non_norm_int[:num_index]   # rakamdan önce
        port = non_norm_int[num_index:]             # rakamdan sonra
    except StopIteration: # Eğer içinde hiç sayı yoksa...
        interface_type = non_norm_int # ismin tamamını harf olarak kabul et
        port = ""                 # sayı kısmını boş bırak

    for int_types in interfaces:
        kisa_isimler_listesi = int_types[0] # Örn: ["GigabitEthernet", "Gi"]
        tam_isim = int_types[1]           # Örn: "GigabitEthernet"

        for name in kisa_isimler_listesi:
            if interface_type == name:
                return tam_isim + port

    return f"{non_norm_int} (normalize edilemedi)"

if __name__ == '__main__':
    excel_dosyasi = 'switch_bilgileri.xlsx'

    while True:
        print("\n" + "+++++++++++++++++++++++++++++++++++")
        print(" + Yeni Switch Bilgilerini Girin + ")
        print("+++++++++++++++++++++++++++++++++++")

        ip = input("Switch IP'sini girin: ")
        username = input("Switch kullanici adi girin: ")
        password = getpass("SSH şifresi: ")

        SwitchManager(ip=ip, username=username, password=password, excel_dosyasi=excel_dosyasi)

        devam_mi =  "h"  #input("\nBaşka bir switch eklemek ister misiniz E/h? : ")
        if devam_mi.lower() not in ['e', 'evet', '']:
            print("Program sonlandiriliyor...")
            break