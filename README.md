# Switch Port Mapper

A Python-based multi-vendor network automation tool that collects
interface, VLAN and LAG information from switches via SSH or console
connection and exports the results to Excel.

Supported platforms include Cisco IOS, Juniper Junos, Aruba and HP Procurve.

###### <p align="center"> *This is official repository maintained by me*</center> </p>
###### <p align="center"> *[yigitdrbk](https://www.instagram.com/yigitdrbk/) *</center> </p>

![PyMapper](/images/py-mapper.png "pymapper")

## Installation

```bash
git clone https://github.com/BAROWZSKI/Switch-Port-Mapper.git
cd Switch-Port-Mapper
pip install -r requirements.txt
python Main.py
```

## Specs
This tool collects interface, VLAN and LAG information from multi-vendor
network switches using SSH or console connection and exports the result to Excel.

![Sheet1](/images/sheet1.png "sheet1")
![Sheet1](/images/sheet2.png "sheet2")

## Supported Vendors
- Cisco IOS
- Juniper Junos
- Aruba
- HP Procurve

## Features
- SSH and console connection support
- Automatic platform detection
- NAPALM + Netmiko fallback
- Interface normalization
- VLAN discovery
- Excel export

### Development by

Developer / Author: [yigitdrbk](https://www.instagram.com/yigitdrbk/)
