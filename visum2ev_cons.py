#!/usr/bin/python
# -*- coding: utf-8 -*-

import pandas as pd
import json

input_data  = {}
output_data = {}

with open('visum_output.json') as infile:
    input_data = json.load(infile)

vehicle_id = input_data["COPERT_XLSX_ROW"]    # row of excel Stock_Configuration.xlsx
if vehicle_id < 1000:
     output_data["Stock"]     = 0
     output_data["energykwh"] = 0
else:
     distance                 = input_data["MEAN_ACTIVITY"]
     energy                   = input_data["ENERGY"]

     result = (energy/100) * distance * pow(3.6,10^(-6))

     output_data["Stock"]     = input_data["STOCK"]
     output_data["energykwh"] = result

df = pd.DataFrame(output_data, index = [0])
df.to_excel("EV_CONS.xlsx", index=False)

