#!/usr/bin/env python3
"""Creates xlsx from show int status and show name"""

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
SHOW_INT_STATUS = "show interfaces status"
SHOW_NAME = "show name"

with open("switches.txt", encoding="utf-8") as file:
    switch_name_list = file.read()
file.close()

for switch_name in switch_name_list.splitlines():
    my_device = {
        "host": switch_name,
        "username": MY_USERNAME,
        "password": MY_PASSWORD,
        "device_type": "hp_procurve",
    }
    net_conn = Netmiko(**my_device)

    MY_FILENAME = switch_name + ".xlsx"
    WORKBOOK = xlsxwriter.Workbook(MY_FILENAME)

    # Excel WORKSHEET name must be <=31
    WORKSHEET = WORKBOOK.add_worksheet("int_status"[0:30])
    COLUMN = 0
    ROW = 0

    # Header
    WORKSHEET.write(ROW, 0, "PORT")
    WORKSHEET.write(ROW, 1, "NAME")
    WORKSHEET.write(ROW, 2, "STATUS")
    WORKSHEET.write(ROW, 3, "CONFIG-MODE")
    WORKSHEET.write(ROW, 4, "SPEED")
    WORKSHEET.write(ROW, 5, "TYPE")
    WORKSHEET.write(ROW, 6, "TAGGED")
    WORKSHEET.write(ROW, 7, "UNTAGGED")
    ROW += 1

    output = net_conn.send_command(SHOW_NAME)
    name_parsed = parse_output(platform="hp_procurve", command=SHOW_NAME, data=output)
    for name_info in name_parsed:
        if isinstance(name_info, dict):
            WORKSHEET.write(ROW, 1, name_info["name"])
            ROW += 1

    output = net_conn.send_command(SHOW_INT_STATUS)
    status_parsed = parse_output(
        platform="hp_procurve", command=SHOW_INT_STATUS, data=output
    )
    ROW = 1
    for status_info in status_parsed:
        if isinstance(status_info, dict):
            WORKSHEET.write(ROW, 0, status_info["port"])
            WORKSHEET.write(ROW, 2, status_info["status"])
            WORKSHEET.write(ROW, 3, status_info["mode"])
            WORKSHEET.write(ROW, 4, status_info["speed"])
            WORKSHEET.write(ROW, 5, status_info["type"])
            WORKSHEET.write(ROW, 6, status_info["tagged"])
            WORKSHEET.write(ROW, 7, status_info["untagged"])
            ROW += 1

    net_conn.disconnect()
    WORKBOOK.close()
    print(MY_FILENAME, " was created.\n\n\n")

print("\n\n")
