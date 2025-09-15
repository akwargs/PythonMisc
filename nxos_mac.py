#!/usr/bin/env python3
"""Creates xlsx of show mac address-table"""

from dotenv import dotenv_values
config = dotenv_values(".env")

import logging
import os
from getpass import getpass

import xlsxwriter
from netmiko import Netmiko
from ntc_templates.parse import parse_output

os.environ["NTC_TEMPLATES_DIR"] = os.environ["MY_NTC_TEMPLATES_DIR"]

logging.basicConfig(filename="output.log", level=logging.INFO)
logger = logging.getLogger("netmiko")

# MY_PASSWORD = getpass()
# MY_USERNAME = ""
MY_PASSWORD = config["PASSWORD"]
MY_USERNAME = config["USERID"]

SHOW_MAC = "show mac address-table"

with open("switches.txt", encoding="utf-8") as file:
    my_switch_list = file.read()
file.close()

for my_switch in my_switch_list.splitlines():
    my_device = {
        "host": my_switch,
        "username": MY_USERNAME,
        "password": MY_PASSWORD,
        "device_type": "cisco_nxos",
    }
    net_conn = Netmiko(**my_device)

    MY_FILENAME = my_switch + "_mac_address-table.xlsx"
    WORKBOOK = xlsxwriter.Workbook(MY_FILENAME)

    # Excel WORKSHEET name must be <=31
    WORKSHEET = WORKBOOK.add_worksheet("show_mac"[0:30])
    COLUMN = 0
    ROW = 0

    # Header
    # e.g: VLAN/BD   MAC Address      Type      age     Secure NTFY Ports/SWID.SSID.LID
    WORKSHEET.write(ROW, 0, "VLAN")
    WORKSHEET.write(ROW, 1, "MAC")
    WORKSHEET.write(ROW, 2, "TYPE")
    WORKSHEET.write(ROW, 3, "AGE")
    WORKSHEET.write(ROW, 4, "SECURE")
    WORKSHEET.write(ROW, 5, "NTFY")
    WORKSHEET.write(ROW, 6, "PORTS")
    ROW += 1

    output = net_conn.send_command(SHOW_MAC)
    mac_parsed = parse_output(platform="cisco_nxos", command=SHOW_MAC, data=output)
    # eg. {'vlan_id': '62', 'mac_address': '1234.5691.1724', 'type': 'dynamic', 'age': '~~~', 'secure': 'F', 'ntfy': 'F', 'ports': 'Po350'}
    for mac_info in mac_parsed:
        if isinstance(mac_info, dict):
            WORKSHEET.write(ROW, 0, mac_info["vlan_id"])
            WORKSHEET.write(ROW, 1, mac_info["mac_address"])
            WORKSHEET.write(ROW, 2, mac_info["type"])
            WORKSHEET.write(ROW, 3, mac_info["age"])
            WORKSHEET.write(ROW, 4, mac_info["secure"])
            WORKSHEET.write(ROW, 5, mac_info["ntfy"])
            WORKSHEET.write(ROW, 6, mac_info["ports"])
            ROW += 1

    net_conn.disconnect()
    WORKBOOK.close()
    print(MY_FILENAME, " was created.\n")

print("\n\n")
