# -*- coding: utf-8 -*-
"""
Updated on April 29, 2019

What's new:
1. Added cancel message at the end the spreadsheet


@author: DanielYuan
"""

#!shipment_schedule5.py
"""Extract the data from Product Plan and write them to a new spreadsheet"""


from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles.colors import RED, BLUE
from openpyxl.comments import Comment
import datetime
import json, re, os
import pandas as pd


# Open the source file to read data from
def extract_data(src_file, *sheets):
    print('Start extracting data from the source file...')
    pattern = re.compile(r'NSO|Reol|Upgr|Relo|Refurb')
    frames = []
    for sheet in sheets:
        frame = pd.read_excel(src_file, sheet)
        frame = frame[['Month','Customer','Product Info','Project ID', 'Crated Date','Ship Date']]
        frame = frame.dropna(axis=0, how='any')
        frame = frame[-frame['Product Info'].str.contains(pattern)]        
        frames.append(frame)
    df = pd.concat(frames,ignore_index=True)
    df['Ship Date'] = df['Ship Date'].astype('str')
    df['Project ID'] = df['Project ID'].astype('int').astype('str') # convert to integer
    # strip '00:00:00' suffixed to the crate date
    crate_date = []
    for i in df['Crated Date'].astype('str'):
        i = i.split()[0]
        crate_date.append(i)
    df['Crated Date'] = crate_date
    return df


def reshape_cmp_data():
    try:
        with open('cmp_data.json') as f_obj:
            compare_file = json.load(f_obj)
    except FileNotFoundError:
        return {}
    else:
        cmp_dict = {}
        for i in compare_file.values():
            cmp_dict[i['Project ID']] = [i['Product Info'], i['Ship Date']]
        print('A compare data file has been created!')
        return cmp_dict
    

# Compare new ship data with existing ship data
def compare_data(df, cmp_dict):
    for i, k in enumerate(df['Project ID']):
        if k in cmp_dict.keys():            
            if df['Ship Date'][i] != cmp_dict[k][1]:
                df = df.replace(k, k+'!')
        else:
            df = df.replace(k, k+'*')
    return df


# Write data in the target file
def write_data(df, new_tar_file):
    print('Saving data to the excel file...')
    df.to_excel(new_tar_file, index=False)  # Save file to excel
    

# Color the specified fonts and add comments to changed schedules
def add_color_comment(new_tar_file, cmp_dict, msg):
    wb = load_workbook(new_tar_file)
    ws = wb.active
    for cell in list(ws.columns)[3]:
        if '*' in cell.value:
            cell.font = Font(color=BLUE)
        elif '!' in cell.value:
            cell.font = Font(color=RED)
            cmt_value = cmp_dict[cell.value.strip('!')][1]
            cell.comment = Comment('Last ship date:\n%s' % cmt_value, 'Author')
    
    ws.append(msg)
    wb.save(new_tar_file)


def cancelled_items(df1, cmp_dict, today):
    cancel_item = []
    for j, h in cmp_dict.items():
        if j not in df1['Project ID'].values and h[1] > today:  # check if a value exists
            print('%s was cancelled...' %j)
            cancel_item.append(j)
    return cancel_item



def file_names():
    '''Get the MFG Product Plan source file and 
    the existing Shipment Schedule file for comparison'''
    while True:
        prompt = ('Please enter the MFG Product Plan source file.\n')
        filename= input(prompt + '>>> ')
        src_file = 'c:\\users\\danielyuan\\desktop\\ProductPlanProject\\%s.xlsx' %filename
        if os.path.exists(src_file):
            break
    now_str = format_time('%m%d%Y')
    new_tar_file = 'c:\\users\\danielyuan\\desktop\\ProductPlanProject\\ShipmentSchedule_%s.xlsx' % now_str
    return src_file, new_tar_file


def format_time(fmt):
    now = datetime.datetime.now()
    return datetime.datetime.strftime(now, fmt)
    

def main():
    src_file, new_tar_file = file_names()
    df = extract_data(src_file, 'M-Plan', 'E-Plan')    
    df1 = df.copy()
    today = format_time('%Y-%m-%d')
    df = df[df['Ship Date'] > today] # Display the shipments after today
    df = df.reset_index(drop=True)    
    # convert dataframe to a compare dictionary
    cmp_dict = reshape_cmp_data()  
    # if cmp_dict has data, do data comparison 
    # or save to excel directly
    if cmp_dict:
        df = compare_data(df,cmp_dict)
        write_data(df, new_tar_file)
        cancel_item = cancelled_items(df1, cmp_dict, today)
        if cancel_item:
            msg = ['Note: The following items were cancelled: %s' % cancel_item]
        else:
            msg = ['Note: No shipments were cancelled!']
        add_color_comment(new_tar_file, cmp_dict, msg)                     
    else:
        write_data(df, new_tar_file)

        
    print('\nA total of %d items has been successfully created!' %len(df))
    # Save compare_data into cmp_data.json(compare_file)
    
    df1.to_json('cmp_data.json', orient='index')


if __name__ == '__main__':
    main()
