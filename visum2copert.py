#!/usr/bin/python
# -*- coding: utf-8 -*-

import pandas as pd
import json

input_data  = {}
output_data = {}

filename = "COPERT.xlsx"
writer = pd.ExcelWriter(filename, engine='xlsxwriter')

#===========================================================
df_sheets = pd.DataFrame({
    "SHEET_NAME": ['=HYPERLINK("'+filename+'#STOCK!$A$1","STOCK")',
                   '=HYPERLINK("'+filename+'#MEAN_ACTIVITY!$A$1","MEAN_ACTIVITY")',
                   '=HYPERLINK("'+filename+'#URBAN_OFF_PEAK_SPEED!$A$1","URBAN_OFF_PEAK_SPEED")',
                   '=HYPERLINK("'+filename+'#URBAN_PEAK_SPEED!$A$1","URBAN_PEAK_SPEED")',
                   '=HYPERLINK("'+filename+'#URBAN_OFF_PEAK_SHARE!$A$1","URBAN_OFF_PEAK_SHARE")',
                   '=HYPERLINK("'+filename+'#URBAN_PEAK_SHARE!$A$1","URBAN_PEAK_SHARE")',
                   '=HYPERLINK("'+filename+'#MIN_TEMPERATURE!$A$1","MIN_TEMPERATURE")',
                   '=HYPERLINK("'+filename+'#MAX_TEMPERATURE!$A$1","MAX_TEMPERATURE")',
                   '=HYPERLINK("'+filename+'#HUMIDITY!$A$1","HUMIDITY")' ],
    "Unit": [ "[n]","[km]","[km/h","[km/h]","[%]","[%]","[C]","[C]","[%]" ]})

df_sheets.to_excel(writer, sheet_name='SHEETS', index = False)

#===========================================================
with open('visum_output.json') as infile:
    input_data = json.load(infile)

dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["STOCK"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='STOCK', index = False)

#---------------------------------------
dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["MEAN_ACTIVITY"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='MEAN_ACTIVITY', index = False)

#---------------------------------------
dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["URBAN_OFF_PEAK_SPEED"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='URBAN_OFF_PEAK_SPEAD', index = False)

#---------------------------------------
dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["URBAN_PEAK_SPEED"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='URBAN_PEAK_SPEAD', index = False)

#---------------------------------------
dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["URBAN_OFF_PEAK_SHARE"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='URBAN_OFF_PEAK_SHARE', index = False)

#---------------------------------------
dict_df = {}
dict_df["Category"]      = input_data["CATEGORY"]
dict_df["Fuel"]          = input_data["FUEL"]
dict_df["Segment"]       = input_data["SEGMENT"]
dict_df["Euro Standard"] = input_data["EURO_STANDARD"]
dict_df["2021"]          = input_data["URBAN_PEAK_SHARE"]

df_df = pd.DataFrame(dict_df, index = [0])
df_df.to_excel(writer, sheet_name='URBAN_PEAK_SHARE', index = False)


#===========================================================
with open('climate_bp.json') as infile:
    input_data = json.load(infile)

list_month    = list(input_data["MONTH"])
list_min_temp = list(input_data["MIN_TEMPERATURE"])
list_max_temp = list(input_data["MAX_TEMPERATURE"])
list_humidity = list(input_data["HUMIDITY"])

#---------------------------------------
dict_min_temp = {}
dict_min_temp["Month"] = list_month
dict_min_temp["2021"]  = list_min_temp

df_min_temp = pd.DataFrame(dict_min_temp)
df_min_temp.to_excel(writer, sheet_name='MIN_TEMPERATURE', index = False)

#---------------------------------------
dict_max_temp = {}
dict_max_temp["Month"] = list_month
dict_max_temp["2021"]  = list_max_temp

df_max_temp = pd.DataFrame(dict_max_temp)
df_max_temp.to_excel(writer, sheet_name='MAX_TEMPERATURE', index = False)

#---------------------------------------
dict_humidity = {}
dict_humidity["Month"] = list_month
dict_humidity["2021"]  = list_humidity

df_humidity = pd.DataFrame(dict_humidity)
df_humidity.to_excel(writer, sheet_name='HUMIDITY', index = False)

#===========================================================

writer.close()