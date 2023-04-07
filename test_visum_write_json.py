#!/usr/bin/python
# -*- coding: utf-8 -*-

import pandas as pd
import json

output_data = {}

vehicle_id = 1001 # row of excel Stock_Configuration.xlsx (2-433) or electric vehicle (1000<)

xlsx = pd.ExcelFile(f"Stock_configuration.xlsx")
sheet_stock_conf = pd.read_excel(xlsx, sheet_name = "STOCK_CONFIGURATION")

output_data["COPERT_XLSX_ROW"] = vehicle_id

if (vehicle_id >=2) and (vehicle_id <= 433):
    output_data["CATEGORY"]        = sheet_stock_conf.loc[vehicle_id-2, "Category"]
    output_data["FUEL"]            = sheet_stock_conf.loc[vehicle_id-2, "Fuel"]
    output_data["SEGMENT"]         = sheet_stock_conf.loc[vehicle_id-2, "Segment"]
    output_data["EURO_STADARSD"]   = sheet_stock_conf.loc[vehicle_id-2, "Euro Standard"]
    output_data["ENERGY"]          = ""
elif (vehicle_id == 1001):
    output_data["CATEGORY"]        = ""
    output_data["FUEL"]            = "Electric"
    output_data["SEGMENT"]         = "MiniBike"
    output_data["EURO_STANDARD"]   = ""
    output_data["ENERGY"]          = 0.1
else:
    output_data["CATEGORY"]        = ""
    output_data["FUEL"]            = "Electric"
    output_data["SEGMENT"]         = "Medium"
    output_data["EURO_STANDARD"]   = ""
    output_data["ENERGY"]          = 30

output_data["STOCK"]                = 3
output_data["MEAN_ACTIVITY"]        = 142.3
output_data["URBAN_OFF_PEAK_SPEED"] = 30
output_data["URBAN_PEAK_SPEED"]     = 15
output_data["URBAN_OFF_PEAK_SHARE"] = 0.6
output_data["URBAN_PEAK_SHARE"]      = 0.4

with open('visum_output.json', 'w') as outfile:
    json.dump(output_data, outfile, indent = 2)
