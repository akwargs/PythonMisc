#!/usr/bin/env python3
"""Creates xlsx combining show int status and show descr"""

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
SHOW_INT_STATUS = "show interface status"
SHOW_INT_DESCR = "show interface description"

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

    MY_FILENAME = my_switch + ".xlsx"
    WORKBOOK = xlsxwriter.Workbook(MY_FILENAME)

    # Excel WORKSHEET name must be <=31
    WORKSHEET = WORKBOOK.add_worksheet("int_status"[0:30])
    COLUMN = 0
    ROW = 0

    # Header
    WORKSHEET.write(ROW, 0, "PORT")
    WORKSHEET.write(ROW, 1, "DESCRIPTION")
    WORKSHEET.write(ROW, 2, "STATUS")
    WORKSHEET.write(ROW, 3, "SPEED")
    WORKSHEET.write(ROW, 4, "DUPLEX")
    WORKSHEET.write(ROW, 5, "VLAN_ID")
    WORKSHEET.write(ROW, 6, "TYPE")
    ROW += 1

    description = {}

    descr_parsed = net_conn.send_command(SHOW_INT_DESCR, use_textfsm=True)
    for descr_info in descr_parsed:
        if isinstance(descr_info, dict):
            description[descr_info["port"]] = descr_info["description"]

    status_parsed = net_conn.send_command(SHOW_INT_STATUS, use_textfsm=True)
    for status_info in status_parsed:
        if isinstance(status_info, dict):
            WORKSHEET.write(ROW, 0, status_info["port"])
            WORKSHEET.write(ROW, 1, description[status_info["port"]])
            WORKSHEET.write(ROW, 2, status_info["status"])
            WORKSHEET.write(ROW, 3, status_info["speed"])
            WORKSHEET.write(ROW, 4, status_info["duplex"])
            WORKSHEET.write(ROW, 5, status_info["vlan_id"])
            WORKSHEET.write(ROW, 6, status_info["type"])
            ROW += 1

    net_conn.disconnect()
    WORKBOOK.close()
    print(MY_FILENAME, " was created.\n")

print("\n\n")
