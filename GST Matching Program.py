#!/usr/bin/env python
# coding: utf-8

import re
import pandas as pd
import numpy as np
from datetime import datetime
import getpass
import time
from progress_bar import InitBar

def highlight(s, column):
    is_max = pd.Series(data = False, index = s.index)
    is_max[column] = s.loc[column] == 0
    return ['background-color: #B2FF59' if is_max.any() else '' for v in is_max]

check_password = 0
password = getpass.getpass('Enter Password: ')

if (password == 'ENTER_YOUR_PASSWORD_HERE'):
    print('*****PASSWORD VERIFIED SUCCESSFULLY*****\n')
    check_passowrd = 1
else:
    print('*****PASSWORD VERIFICATION FAILED*****\n')
    input('Press ENTER to exit')
    
while (check_passowrd == 1):
    pbar = InitBar()

    file_name = input('\nEnter File Name: ')
    
    pbar(0)
    time.sleep(0.5)

    my_data_colnames = ['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date',
                        'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess']

    my_govt_colnames = ['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date',
                        'Filling Date', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess']

    pbar(5)
    time.sleep(0.5)

    try:
        my_data = pd.read_excel(file_name + '.xlsx', sheet_name = 'Tally Data', usecols = my_data_colnames, engine = 'openpyxl')
    except:
        print('\n\n*****ERROR OCCURRED: FILE NOT FOUND*****\n\n')
        input('Press ENTER to exit')
        break
    
    try:
        if pd.api.types.is_string_dtype(my_data['GSTIN of Supplier']):
            my_data['GSTIN of Supplier'] = my_data['GSTIN of Supplier'].str.replace(' ', '')
        
        if pd.api.types.is_string_dtype(my_data['Invoice Number']):
            my_data['Invoice Number'] = my_data['Invoice Number'].str.replace(' ', '')
        
        if pd.api.types.is_string_dtype(my_data['Invoice Date']):
            my_data['Invoice Date'] = my_data['Invoice Date'].str.replace(' ', '')

        my_data = my_data.replace({"^\s*|\s*$":""}, regex = True)
    
        my_data['Invoice Date'] = my_data['Invoice Date'].dt.strftime('%d-%m-%Y')
    
        my_data['Integrated Tax'].fillna(0.0, inplace = True)
        my_data['Central Tax'].fillna(0.0, inplace = True)
        my_data['State/UT Tax'].fillna(0.0, inplace = True)
        my_data['Cess'].fillna(0.0, inplace = True)

        my_data = my_data[['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date',
                           'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess']]
    
        my_data = my_data.groupby(['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date']).sum().reset_index()
    
        pregna_drop_values = my_data.loc[
            (my_data['Integrated Tax'] == 0) & (my_data['Central Tax'] == 0) & 
            (my_data['State/UT Tax'] == 0) & (my_data['Cess'] == 0)
        ]

        my_data.drop(pregna_drop_values.index, inplace = True)

        my_data['Integrated Tax'] = my_data['Integrated Tax'].round(0)
        my_data['Central Tax'] = my_data['Central Tax'].round(0)
        my_data['State/UT Tax'] = my_data['State/UT Tax'].round(0)
        my_data['Cess'] = my_data['Cess'].round(0)

        my_data['Total'] = (my_data['Integrated Tax'] + my_data['Central Tax'] + my_data['State/UT Tax'] + my_data['Cess'])
    except:
        print('\n\n*****ERROR OCCURRED: Tally Data Problem*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(10)
    time.sleep(0.5)

    my_govt = pd.read_excel(file_name + '.xlsx', sheet_name = 'Government Data', usecols = my_govt_colnames, engine = 'openpyxl')
    
    try:
        if pd.api.types.is_string_dtype(my_govt['GSTIN of Supplier']):
            my_govt['GSTIN of Supplier'] = my_govt['GSTIN of Supplier'].str.replace(' ', '')

        if pd.api.types.is_string_dtype(my_govt['Invoice Number']):
            my_govt['Invoice Number'] = my_govt['Invoice Number'].str.replace(' ', '')
    
        if pd.api.types.is_string_dtype(my_govt['Invoice Date']):
            my_govt['Invoice Date'] = my_govt['Invoice Date'].str.replace(' ', '')
    
        if pd.api.types.is_string_dtype(my_govt['Filling Date']):
            my_govt['Filling Date'] = my_govt['Filling Date'].str.replace(' ', '')
        
        my_govt = my_govt.replace({"^\s*|\s*$":""}, regex = True)
    
        my_govt['Invoice Date'] = pd.to_datetime(my_govt['Invoice Date'])

        my_govt['Invoice Date'] = my_govt['Invoice Date'].dt.strftime('%d-%m-%Y')
    
        my_govt['Integrated Tax'].fillna(0.0, inplace = True)
        my_govt['Central Tax'].fillna(0.0, inplace = True)
        my_govt['State/UT Tax'].fillna(0.0, inplace = True)
        my_govt['Cess'].fillna(0.0, inplace = True)

        my_govt = my_govt[['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date',
                           'Filling Date', 'Integrated Tax', 'Central Tax', 'State/UT Tax', 'Cess']]
    
        my_govt = my_govt.groupby(['GSTIN of Supplier', 'Name of Supplier', 'Invoice Number', 'Invoice Date',
                                   'Filling Date']).sum().reset_index()
    
        portal_drop_values = my_govt.loc[
            (my_govt['Integrated Tax'] == 0) & (my_govt['Central Tax'] == 0) & 
            (my_govt['State/UT Tax'] == 0) & (my_govt['Cess'] == 0)
        ]

        my_govt.drop(portal_drop_values.index, inplace = True)
    
        my_govt['Integrated Tax'] = my_govt['Integrated Tax'].round(0)
        my_govt['Central Tax'] = my_govt['Central Tax'].round(0)
        my_govt['State/UT Tax'] = my_govt['State/UT Tax'].round(0)
        my_govt['Cess'] = my_govt['Cess'].round(0)

        my_govt['Total'] = (my_govt['Integrated Tax'] + my_govt['Central Tax'] + my_govt['State/UT Tax'] + my_govt['Cess'])
    except:
        print('\n\n*****ERROR OCCURRED: Government Data Problem*****\n\n')
        input('Press ENTER to exit')
        break
    
    pbar(20)
    time.sleep(0.5)

    try:
        final_data = my_data.merge(my_govt, on = ['GSTIN of Supplier', 'Invoice Number'], how = 'outer', sort = True, suffixes = ('_Tally', '_Govt'))
    except:
        print('\n\n*****ERROR OCCURRED: Could not Merge Data*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(25)
    time.sleep(0.5)

    try:
        final_data['Integrated Tax_Tally'].fillna(0.0, inplace = True)
        final_data['Integrated Tax_Govt'].fillna(0.0, inplace = True)
        final_data['Central Tax_Tally'].fillna(0.0, inplace = True)
        final_data['Central Tax_Govt'].fillna(0.0, inplace = True)
    except:
        print('\n\n*****ERROR OCCURRED*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(35)
    time.sleep(0.5)

    try:
        final_data['State/UT Tax_Tally'].fillna(0.0, inplace = True)
        final_data['State/UT Tax_Govt'].fillna(0.0, inplace = True)
        final_data['Cess_Tally'].fillna(0.0, inplace = True)
        final_data['Cess_Govt'].fillna(0.0, inplace = True)
        final_data['Total_Tally'].fillna(0.0, inplace = True)
        final_data['Total_Govt'].fillna(0.0, inplace = True)
    except:
        print('\n\n*****ERROR OCCURRED*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(55)
    time.sleep(0.5)

    try:
        final_data = final_data[['GSTIN of Supplier', 'Name of Supplier_Tally', 'Name of Supplier_Govt', 'Invoice Number',
                                 'Invoice Date_Tally', 'Invoice Date_Govt', 'Filling Date',
                                 'Integrated Tax_Tally', 'Integrated Tax_Govt', 'Central Tax_Tally',
                                 'Central Tax_Govt', 'State/UT Tax_Tally', 'State/UT Tax_Govt', 'Cess_Tally',
                                 'Cess_Govt', 'Total_Tally', 'Total_Govt']]
    except:
        print('\n\n*****ERROR OCCURRED*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(65)
    time.sleep(0.5)
    
    try:
        drop_values = final_data.loc[
            (final_data['Integrated Tax_Tally'] == 0) & (final_data['Integrated Tax_Govt'] == 0) &
            (final_data['Central Tax_Tally'] == 0) & (final_data['Central Tax_Govt'] == 0) &
            (final_data['State/UT Tax_Tally'] == 0) & (final_data['State/UT Tax_Govt'] == 0) &
            (final_data['Cess_Tally'] == 0) & (final_data['Cess_Govt'] == 0) &
            (final_data['Total_Tally'] == 0) & (final_data['Total_Govt'] == 0)
        ]
    
        final_data.drop(drop_values.index, inplace = True)

        final_data['Difference'] = (final_data['Total_Tally'] - final_data['Total_Govt'])
    except:
        print('\n\n*****ERROR OCCURRED*****\n\n')
        input('Press ENTER to exit')
        break

    pbar(80)
    time.sleep(0.5)

    now = datetime.now()
    datestr = now.strftime('%d-%m-%Y')
    timestr = now.strftime('%I %M %p')

    pbar(90)
    
    final_data.style.apply(highlight, column='Difference', axis=1).to_excel('Final_data(Date-' + datestr + ' Time-' + timestr + ').xlsx', engine = 'openpyxl', index = False)

    pbar(100)

    check_passowrd = False

    print('\n\n*****FILE SUCCESSFULLY GENERATED*****\n\n')
    input('Press ENTER to exit')

