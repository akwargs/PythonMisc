#!/usr/bin/env python3
"""show mac address table"""

from dotenv import dotenv_values
config = dotenv_values(".env")

import logging
from getpass import getpass

import xlsxwriter
from netmiko import Netmiko

logging.basicConfig(filename="output.log", level=logging.INFO)
logger = logging.getLogger("netmiko")

# MY_PASSWORD = getpass()
# MY_USERNAME = ""
MY_PASSWORD = config["PASSWORD"]
MY_USERNAME = config["USERID"]
SHOW_MAC = "show mac-address"

with open("switches.txt", encoding="utf-8") as file:
    my_switch_list = file.read()
file.close()

for my_switch in my_switch_list.splitlines():
    my_device = {
        "host": my_switch,
        "username": MY_USERNAME,
        "password": MY_PASSWORD,
        "device_type": "hp_procurve",
    }
    net_conn = Netmiko(**my_device)

    MY_FILENAME = my_switch + "_mac_address-table.xlsx"
    WORKBOOK = xlsxwriter.Workbook(MY_FILENAME)

    # Excel WORKSHEET name must be <=31
    WORKSHEET = WORKBOOK.add_worksheet("show_mac"[0:30])
    COLUMN = 0
    ROW = 0

    # Header
    WORKSHEET.write(ROW, 0, "PORT")
    WORKSHEET.write(ROW, 1, "MAC")
    WORKSHEET.write(ROW, 2, "VLAN")
    ROW += 1

    status_parsed = net_conn.send_command(SHOW_MAC, use_textfsm=True)
    # eg. {'mac_address': '12346b-5aefc8', 'port': 'Trk1', 'vlan_id': '980'}
    for status_info in status_parsed:
        if isinstance(status_info, dict):
            WORKSHEET.write(ROW, 0, status_info["port"])
            WORKSHEET.write(ROW, 1, status_info["mac_address"])
            WORKSHEET.write(ROW, 2, status_info["vlan_id"])
            ROW += 1

    net_conn.disconnect()
    WORKBOOK.close()
    print(MY_FILENAME, " was created.\n")

print("\n\n")
