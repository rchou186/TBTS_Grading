###############################################################################
### Grading program for TBTS                                                ###
### The test raw csv data is in the same folder of this program             ###
### DCD Chart output is in html file                                        ###
### Grading result table and printable result sheet is in the excel file    ###
### Module to install: pandas, openpyxl, plotly                             ###
###############################################################################

import datetime
import glob
import os
#import platform
import pathlib

import pandas as pd

from defines import *
from printout import print_result
from result_to_chart import result_to_chart
from result_to_excel import result_to_excel
from result_to_sql import result_to_sql

version = 'V22.0215.01'  # 'V20.1112.01'

# chaege time threshold, 2099 for DCD35, 2219 for DCD37
CHR_TIME_THRES = 2219

# VDT threshold, 24 seconds, 0 for not checking
VDT_THRES = 0

# Last class for extended grades of U1~U4, was L1~L4
LAST_CLASS = 0

# Grades, G[0], G[1], G[2], ...
G = ['0', 'D', 'F', 'G', 'N', 'Y', 'Z', 'U1', 'U2', 'U3', 'U4']


def pre_grading(dis_time_sec):
    if dis_time_sec >= 2190:        # 36.5 * 60
        return G[1]
    elif dis_time_sec >= 2160:      # 36 * 60
        return G[2]
    elif dis_time_sec >= 2130:      # 35.5 * 60
        return G[3]
    elif dis_time_sec >= 2100:      # 35 * 60
        return G[4]
    elif dis_time_sec >= 2070:      # 34.5 * 60
        return G[5]
    elif dis_time_sec >= 2040:      # 34 *60
        return G[6]
    else:
        return G[0]


# output table defination
table = pd.DataFrame(
    columns=['B_SN', 'Battery', 'Org_Module', 'M_SN', 'DIS0_Wh', 'Norm_D0', 'CHR_Wh', 'DIS_Wh',
             'Ratio_Wh', 'CHR_Time', 'DIS_Time', 'Ratio_Time', 'V15min', 'V20min', 'V25min', 'V30min',
             'V35min', 'DIS0_Time', 'RV10s', 'VDT', 'BIN', 'Out_Battery', 'Out_Module', 'FG_SN',
             'Station', 'Date', 'Time', 'Model', 'VPVD', 'TPtE', 'VEOC', 'VDeltaMax', 'RSN', 'GSN'])

# set Table column to desired dtype for following calculation
table = table.astype({'DIS0_Wh': 'float', 'Norm_D0': 'float', 'CHR_Wh': 'float', 'DIS_Wh': 'float', 'Ratio_Wh': 'float',
                      'Ratio_Time': 'float', 'V15min': 'float', 'V20min': 'float', 'V25min': 'float', 'V30min': 'float',
                      'V35min': 'float', 'RV10s': 'float', 'VDT': 'int', 'VPVD': 'float', 'TPtE': 'int',
                      'VEOC': 'float', 'VDeltaMax': 'float'})

# check platform and set pathes
# if platform.system() == 'Darwin':  # Mac
#    CID_path = pathlib.Path("/Volumes/Battery Test Data/CID")
# elif platform.system() == 'Windows':  # Windows
#    CID_path = pathlib.Path("Z:/Battery Test Data/CID")

# get current working path
current_path = pathlib.Path().resolve()

# get csv files
csv_files = sorted(glob.glob(str(current_path / 'M??.CSV')))
number_of_modules = len(csv_files)

# get B_SN from M01 for checking battery model
try:
    with open(csv_files[0], 'r', encoding='latin-1') as f:
        line1 = f.readline().strip().split(',')  # read the first line
        bsn = line1[9]
        battery = line1[1]

# if no CSV files in folder
except:
    print("No CSV files in folder.")
    exit()

print(f"Grading battery: {battery}")

# check battery model from number of modules and B_SN
# battery_model: 1:'34M',2:'28M',3:'18M+16M', 4:'32M', 5:'40M', 6:'34M+6M', 7:'20M+20M', 8:'20M'
if number_of_modules == 20:
    battery_model = 8
elif number_of_modules == 28:
    battery_model = 2
elif number_of_modules == 32:
    battery_model = 4
elif number_of_modules == 34:
    if bsn[6:8] == '48':        # B_SN for 18+16 is 2345P148xxx
        battery_model = 3
    else:                       # B_SN for 34 is 2345P133xxx or T22101001
        battery_model = 1
elif number_of_modules == 40:
    if bsn[6:11] in {'30090', '30091'}:             # B_SN for 34+6
        battery_model = 6
    elif bsn[6:11] in {'52030', '52031'}:           # B_SN for 20+20
        battery_model = 7
    elif bsn[6:11] in {'30011', '30030', '30060'}:  # B_SN for 40
        battery_model = 5
    else:                                           # others for 40
        battery_model = 5
else:
    print("Module count incorrect!!!")
    exit()

print(f"{battery} is a {get_model_result(battery_model)} battery.")

df_all = pd.DataFrame()

# process each csv file
for i, csv_file in enumerate(csv_files):
    row = [None] * 34
    df = pd.read_csv(csv_file, header=3, sep=',', encoding='latin-1')
    df.columns = ['Time', 'Voltage', 'Current', 'Ah', 'Wh', 'Temp', 'Phase', 'DCIR', 'VMOS', 'VDelta', 'TempMOS']

    # open the csv file to get the first and second line for parsing the background data
    with open(csv_file, 'r', encoding='latin-1') as f:
        line1 = f.readline().strip().split(',')  # read the first line
        line2 = f.readline().strip().split(',')  # read the second line
    row[0] = line1[9]       # B_SN
    row[1] = line1[1]       # Battery
    row[2] = csv_file[-7:-4]  # Org_Module
    row[3] = line1[3]       # M_SN
    row[24] = line1[13]     # Station
    try:
        row[32] = line1[15]  # RSN
    except:
        pass
    try:
        row[33] = line1[17]  # GSN
    except:
        pass
    # row[25] = str(datetime.datetime.strptime(line2[1][1:], "%m-%d-%Y"))  # Date
    # row[26] = str(datetime.datetime.strptime(line2[2][1:], "%H:%M:%S"))  # Time
    row[25] = line2[1][1:]  # Date
    row[26] = line2[2][1:]  # Time
    row[27] = '34M'  # Model

    # if ERROR occurs during test, write 0 value in table
    if (row[24] == 'ST-XX') or (4 not in df.loc[:, 'Phase'].values) or (df.loc[df['Phase'] == 4]['Voltage'].values[-1] >= 6.0):  # The error module
        row[4] = 0          # DIS0_Wh
        row[6] = 0          # CHR_Wh
        row[7] = 0          # DIS_Wh
        row[8] = 0          # Ratio_Wh
        row[9] = 0          # CHR_Time
        row[10] = 0         # DIS_Time
        row[11] = 0         # Ratio_Time
        row[17] = 0         # DIS0_Time
        row[18] = 0         # RV10s
        row[19] = 0         # VDT
        row[20] = 'ERR'     # BIN
        row[28] = 0         # VPVD
        row[29] = 0         # TPtE
        row[30] = 0         # VEOC
        row[31] = 0         # VDeltaMax
    else:
        # remove last few rows of phase 0
        if (df.iloc[-4]['Phase'] != 0):
            if (df.iloc[-3]['Phase'] == 0):
                df.drop(df.tail(3).index, inplace=True)
            elif (df.iloc[-2]['Phase'] == 0):
                df.drop(df.tail(2).index, inplace=True)
            elif (df.iloc[-1]['Phase'] == 0):
                df.drop(df.tail(1).index, inplace=True)

        # get Wh of each phases
        row[4] = -df.loc[df['Phase'] == 0]['Wh'].values[-1]                                     # DIS0_Wh
        row[6] = df.loc[df['Phase'] == 2]['Wh'].values[-1]                                      # CHR_Wh
        row[7] = -df.loc[df['Phase'] == 4]['Wh'].values[-1]                                     # DIS_Wh
        row[8] = row[7] / row[6]                                                                # Ratio_Wh

        # add 1 second of time in phase 2(CHR) and phase  4(DIS)
        df.loc[df['Phase'] == 2, 'Time'] = df.loc[df['Phase'] == 2, 'Time'] + 1
        df.loc[df['Phase'] == 4, 'Time'] = df.loc[df['Phase'] == 4, 'Time'] + 1

        # get Time of each phases
        row[17] = (df.loc[df['Phase'] == 0]['Time'].values[-1])                                 # DIS0_Time
        row[9] = (df.loc[df['Phase'] == 2]['Time'].values[-1])                                  # CHR_Time
        row[10] = (df.loc[df['Phase'] == 4]['Time'].values[-1])                                 # DIS_Time
        row[11] = row[10] / row[9]                                                              # Ratio_Time

        # get 5min voltages
        try:
            row[12] = df.loc[(df['Phase'] == 2) & (df['Time'] == 900)]['Voltage'].values[0]     # V15min
        except:
            pass
        try:
            row[13] = df.loc[(df['Phase'] == 2) & (df['Time'] == 1200)]['Voltage'].values[0]    # V20min
        except:
            pass
        try:
            row[14] = df.loc[(df['Phase'] == 2) & (df['Time'] == 1500)]['Voltage'].values[0]    # V25min
        except:
            pass
        try:
            row[15] = df.loc[(df['Phase'] == 2) & (df['Time'] == 1800)]['Voltage'].values[0]    # V30min
        except:
            pass
        try:
            row[16] = df.loc[(df['Phase'] == 2) & (df['Time'] == 2100)]['Voltage'].values[0]    # V35min
        except:
            pass
        row[18] = df.loc[df['Phase'] == 1]['Voltage'].values[9]                                 # RV10s
        try:
            row[19] = int(df.loc[(df['Phase'] == 0) & (df['Voltage'] < 6.0)]['Time'].values[0] -
                          df.loc[(df['Phase'] == 0) & (df['Voltage'] < 6.4)]['Time'].values[0])  # VDT
        except:
            pass

        vpeak_idx = df.loc[df['Phase'] == 2]['Voltage'].idxmax()
        row[28] = df.loc[vpeak_idx]['Voltage']                                                  # VPVD
        row[29] = int(df.loc[df['Phase'] == 2]['Time'].values[-1] - df.loc[vpeak_idx]['Time'])  # TPtE
        row[30] = df.loc[df['Phase'] == 2]['Voltage'].values[-1]                                # VEOC
        row[31] = df.loc[(df['Phase'] == 0) | (df['Phase'] == 4)]['VDelta'].max()               # VDeltaMax

    # add row to table
    table.loc[len(table)] = row

    # remove phase 1 and phase 3
    df = df[(df.Phase != 1) & (df.Phase != 3)]

    # insert module Mxx in the first column to differentiate module in df_all
    df.insert(0, 'Module', 'M{0:02d}'.format(i+1))

    # append df to df_all for charting
    df_all = pd.concat([df_all, df])

# if with ERR module, calculate and find the middle number of DIS0_Wh and replace to that ERR module

# sort DIS0_Wh to find the middle number
s = table['DIS0_Wh'].sort_values().reset_index(drop=True)

# find the middle number of DIS0_Wh
mid = s[(number_of_modules / 2) - 1]

# get the index of that middle number
idx = abs(table['DIS0_Wh'] - mid).idxmin()

for i in range(len(table)):
    # copy the middle DIS0_Wh module to ERR module
    if (table.loc[i, 'BIN'] == 'ERR'):
        table.loc[i, 'DIS0_Wh'] = table.loc[idx, 'DIS0_Wh']

    # check if sense errors
    elif ((table.loc[i, 'VPVD'] > 9.2) or ((table.loc[i, 'CHR_Time'] < CHR_TIME_THRES) and (table.loc[i, 'TPtE'] < 20)) or
         ((table.loc[i, 'CHR_Time'] >= CHR_TIME_THRES) and (table.loc[i, 'VEOC'] > 8.95)) or (table.loc[i, 'VDT'] > 250)):
          table.loc[i, 'BIN'] = 'SER'

# calculate Norm_D0
if battery_model == 3:      # 18+16
    middle_module = 18
elif battery_model == 6:    # 34+6
    middle_module = 34
elif battery_model == 7:    # 20+20
    middle_module = 20
else:
    middle_module = number_of_modules

# first half of the pack
table_half = table.iloc[0:middle_module]
avst = table_half['DIS0_Wh'].mean() + table_half['DIS0_Wh'].std(ddof=0)  # ddof=0 makes the stdevp instead od stdev
for i in range(0, middle_module):
    table.loc[i, 'Norm_D0'] = table.loc[i, 'DIS0_Wh'] / avst

# second half of the pack
table_half = table.iloc[middle_module:len(table)]
avst = table_half['DIS0_Wh'].mean() + table_half['DIS0_Wh'].std(ddof=0)  # ddof=0 makes the stdevp instead od stdev
for i in range(middle_module, len(table)):
    table.loc[i, 'Norm_D0'] = table.loc[i, 'DIS0_Wh'] / avst

# Bin grading
ColorList = pd.Series([], dtype='object')  # seperate color list states the color for BIN
good_module = 0
for i in range(len(table)):
    ColorList[i] = BlankFill
    if table.loc[i, 'BIN'] != 'ERR' and table.loc[i, 'BIN'] != 'SER':  # if not ERR or SER module
        pregrading = pre_grading(table.loc[i, 'DIS_Time'])
        if table.loc[i, 'DIS_Time'] >= 2040:                # >=34:00
            if table.loc[i, 'CHR_Time'] >= 2219:            # >=36:59
                if table.loc[i, 'VDT'] > VDT_THRES:
                    if table.loc[i, 'Norm_D0'] >= 0.5 and pregrading == G[1]:       # was 0.7 for DCD35
                        table.loc[i, 'BIN'] = G[1]
                        ColorList[i] = GreenFill
                        good_module += 1
                    elif table.loc[i, 'Norm_D0'] >= 0.5 and pregrading == G[2]:       # was 0.75 for DCD35
                        table.loc[i, 'BIN'] = G[2]
                        ColorList[i] = PurpleFill
                        good_module += 1
                    elif table.loc[i, 'Norm_D0'] >= 0.5:                            # was 0.8 for DCD35
                        if pregrading == G[3]:
                            table.loc[i, 'BIN'] = G[3]
                            ColorList[i] = OrangeFill
                            good_module += 1
                        elif pregrading == G[4]:
                            table.loc[i, 'BIN'] = G[4]
                            ColorList[i] = WhiteFill
                            good_module += 1
                        elif pregrading == G[5]:
                            table.loc[i, 'BIN'] = G[5]
                            ColorList[i] = BlueFill
                            good_module += 1
                        elif pregrading == G[6]:
                            table.loc[i, 'BIN'] = G[6]
                            ColorList[i] = RedFill
                            good_module += 1
                        else:
                            table.loc[i, 'BIN'] = 'NG'
                    else:
                        table.loc[i, 'BIN'] = 'NGD0'
                else:
                    if table.loc[i, 'Norm_D0'] >= 0.7 and pregrading == G[1]:       # not checking VDT if K
                        table.loc[i, 'BIN'] = G[1]
                        ColorList[i] = LGreenFill
                        good_module += 1
                    else:
                        table.loc[i, 'BIN'] = 'NGvdt'
            else:
                table.loc[i, 'BIN'] = 'PVD'  # if charge time < 36:59
        else:
            table.loc[i, 'BIN'] = 'NG'

        # PVD_H_E_grading
        if table.loc[i, 'BIN'] == 'PVD':
            if table.loc[i, 'VDT'] > VDT_THRES:
                if table.loc[i, 'DIS_Time'] >= 2070:    # >=34:30, YP=EPVD
                    table.loc[i, 'BIN'] = G[5]
                    ColorList[i] = LBlueFill
                elif table.loc[i, 'DIS_Time'] >= 2040:  # >=34:00, ZP=HPVD
                    table.loc[i, 'BIN'] = G[6]
                    ColorList[i] = LRedFill
            else:
                table.loc[i, 'BIN'] = 'NGvdt'

        # extended grading of U1~U4 grades
        if table.loc[i, 'BIN'] == "NG" and table.loc[i, 'Ratio_Time'] >= 0.92 and LAST_CLASS != 0:
            if table.loc[i, 'VDT'] > VDT_THRES:
                if table.loc[i, 'DIS_Time'] >= 2010 and LAST_CLASS >= 1:    # >=33:30
                    table.loc[i, 'BIN'] = G[7]
                elif table.loc[i, 'DIS_Time'] >= 1980 and LAST_CLASS >= 2:  # >=33:00
                    table.loc[i, 'BIN'] = G[8]
                elif table.loc[i, 'DIS_Time'] >= 1950 and LAST_CLASS >= 3:  # >=32:30
                    table.loc[i, 'BIN'] = G[9]
                elif table.loc[i, 'DIS_Time'] >= 1920 and LAST_CLASS >= 4:  # >=32:00
                    table.loc[i, 'BIN'] = G[10]
            else:
                table.loc[i, 'BIN'] = "NGvdt"

# if All module pass, check if discharge time ranging is less than 45 second for Camery and ES300h
if good_module == len(table) and battery_model == 1:
    DIS_time_list = table['DIS_Time']
    if DIS_time_list.max() - DIS_time_list.min() < 45:
        avg_dis_time = DIS_time_list.mean()
        if avg_dis_time > 2190:
            table.loc[:, 'BIN'] = G[1]
        elif avg_dis_time > 2160:
            table.loc[:, 'BIN'] = G[2]
        elif avg_dis_time > 2130:
            table.loc[:, 'BIN'] = G[3]
        elif avg_dis_time > 2100:
            table.loc[:, 'BIN'] = G[4]
        elif avg_dis_time > 2070:
            table.loc[:, 'BIN'] = G[5]
        elif avg_dis_time > 2040:
            table.loc[:, 'BIN'] = G[6]

# convert second to time format. before this point, the time is integer of second
table['CHR_Time'] = pd.to_datetime(table['CHR_Time'], unit='s').dt.strftime('%H:%M:%S')
table['DIS_Time'] = pd.to_datetime(table['DIS_Time'], unit='s').dt.strftime('%H:%M:%S')
table['DIS0_Time'] = pd.to_datetime(table['DIS0_Time'], unit='s').dt.strftime('%H:%M:%S')

result_filename = battery + get_model_suffix(battery_model)

# output result csv file, JxxTxxxx-(xx).csv
#print(f"Save CSV file: {result_filename}.csv")
#table.to_csv((result_filename+'.csv'), index=False)

# output result chart file, JxxTxxxx-(xx).html
print(datetime.datetime.now(), f"Saving chart file: {result_filename}.html")
result_to_chart(df_all, result_filename)

# output result excel file, JxxTxxxx-(xx).xlsx
print(datetime.datetime.now(), f"Saving result excel file: {result_filename}.xlsx")
result_to_excel(table, ColorList, battery_model, version, result_filename)

# store result on MySQL server
print(datetime.datetime.now(), "Saving result to MySQL database")
result_to_sql(table)

# print result sheet to printer
print(datetime.datetime.now(), f'Printing result: {result_filename}')
print_result(str(current_path / (result_filename+'.xlsx')))
