import os; import platform; import datetime
import pandas as pd; import numpy as np; import matplotlib.pyplot as pp
from tkinter import filedialog; from tkinter import *; from pathlib import Path

EXTEN = '.jpg'
#INIT_DIR = str(Path.home()) + '/Pictures/Deer/2020 Season/test'
#HOME = str(Path.home())
CWD = os.getcwd()


def main():
    pic = []; yr = []; mo = []; day = []; hr = []; mi = []
    print_header()
    #picture_dir = get_pic_directory()
    picture_dir = "/home/firekracker50/Documents/WIP Programs/DeerMovementTracker/Test Pics/"
    #excel_file = get_output_file()
    excel_file = "/home/firekracker50/Documents/WIP Programs/DeerMovementTracker/test.xlsx"
    clear_file(excel_file)
    for dirpath, dirnames, files in os.walk(picture_dir):
        for name in files:
            if name.lower().endswith(EXTEN):
                path = os.path.join(dirpath, name)
                time_data = get_date_time(path)
                yr.append(time_data.year)
                mo.append(time_data.month)
                day.append(time_data.day)
                if time_data.minute<15:
                    mi.append(0)
                    hr.append(time_data.hour)
                elif time_data.minute>=15 and time_data.minute <45:
                    mi.append(30)
                    hr.append(time_data.hour)
                elif time_data.minute>=45:
                    mi.append(0)
                    hr.append(time_data.hour +1)
    keys = ["Year", "Month", "Day", "Hour", "Minute", "Count"]
    pic_data_dict = dict(zip(keys,[yr,mo,day,hr,mi,0]))
    
    for i in range(24):
            print(pic_data_dict["Hour","Count"])
            if pic_data_dict["Hour"]=i
                Pass
            else
                pic_data_dict["Hour"=i, "Count"=0, "Year"=0, "Month"=0, "Day"=0, "Minute"=0]
                print(pic_data_dict["Hour", "Count"])
    
    pt_time = make_time_pt(pic_data_dict)
    make_excel(pt_time, excel_file)


def print_header():
    print('------------------------------------------------------------------')
    print("--                  Deer Movement Tracker                       --")
    print('------------------------------------------------------------------')

def get_pic_directory():
    pic_dir = filedialog.askdirectory(initialdir=CWD,
                                    title='Select Directory Containing Your Pictures')
    return pic_dir

def get_output_file():
    out_file = filedialog.asksaveasfilename(initialdir=CWD, title='Select Where to Save Output', 
                                            filetypes = [("Excel Files", "*.xlsx")])
    return out_file

def clear_file(clears_file):
    if os.path.isfile(clears_file):  # check for the existence of a spreadsheet
        os.remove(clears_file)

def get_date_time(pic_file):
    t_stamp = os.path.getmtime(pic_file)
    date_time = datetime.datetime.fromtimestamp(t_stamp)
    return date_time

def make_time_pt(data_list):
    df = pd.DataFrame(data_list, columns=["Year", "Month", "Day", "Hour", "Minute", "Count"])
    df.set_index('Year', drop=True, append=False, inplace=True)
    pt = pd.pivot_table(df, index=["Hour"], values=["Count"], 
                            aggfunc=len, margins=True)
    return pt

def make_excel(piv_table, xl_file):
    sheet_name = 'Pivot Table'
    writer = pd.ExcelWriter(xl_file, engine='xlsxwriter')
    piv_table.to_excel(writer, sheet_name=sheet_name)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    chart = workbook.add_chart({'type':'column'})
    #chart.add_series({'categories': [sheetname, first_row, first_col, last_row, last_col],
    #                  'values': [sheetname, first_row, first_col, last_row, last_col]})
    last_row = len(piv_table)-1
    chart.add_series({'values':['Pivot Table',1,1,last_row,1], 
                        'categories':['Pivot Table',1,0,last_row,0], 'gap':2})
    chart.set_x_axis({'name': "Hour", 'min':0, 'max':24})
    chart.set_y_axis({'major_gridlines':{'visible':True}})
    chart.set_legend({'position':'none'})
    worksheet.insert_chart('D2', chart)
    writer.save()


if __name__ == '__main__':
    main()