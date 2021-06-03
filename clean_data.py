import pandas as pd
import glob, os
import re
from openpyxl import Workbook
from openpyxl import load_workbook
import openpyxl
import time
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows

#file_name = "prova.xlsx"
id_company = "identifiers.xlsx"

in_par_folder = "resources/parameters/"
in_data_folder = "resources/to_be_cleaned/"
out_folder="results/"

def clean_err(df):
    df = df.replace(to_replace=r'\$\$ER:(.*)', value='', regex=True)
    df = df.replace(np.nan, '', regex=True)
    return df

def locate_data(df):
    row_tobecleaned = df.index[(df[0] == 'CURRENCY')].tolist()
    row_tobecleaned = max(row_tobecleaned)
    start_row = row_tobecleaned + 1 
    end_row = df.index[-1]  
    return start_row, end_row

def create_excel(df,fname):
    df.to_excel(out_folder + fname, engine='xlsxwriter') 

def clean_data_lastestvalue(df, sheet_name):
    df = clean_err(df)
    start_row, end_row = locate_data(df)
    df = df.loc[(start_row - 2), :]
    df[0] = sheet_name
    Id = df.index.tolist()
    Id[0] = 'Id'
    d = {'col1': Id, 'col2': df}
    df = pd.DataFrame(d)
    return df

def clean_multiple_lastestvalue(fname):
    xlsx = pd.ExcelFile(in_data_folder+fname, engine='openpyxl')
    sheet_names = xlsx.sheet_names
    df = pd.read_excel(in_data_folder+fname, sheet_name=None, engine='openpyxl', header=None)
    df[sheet_names[0]]['Id'] = df[sheet_names[0]].index
    df[sheet_names[0]].rename(columns={ df[sheet_names[0]].columns[0]: df[sheet_names[0]].iloc[0,0] }, inplace = True)
    df[sheet_names[0]] = df[sheet_names[0]].loc[1:, :] 
    df[sheet_names[0]].to_csv(out_folder + 'cleaned/identifiers/'+ fname[:-5] + '_' +'identifiers.csv', index=None)    
    for sheet_name in sheet_names[1:]:
        df[sheet_name] = clean_data_lastestvalue(df[sheet_name], sheet_name)
        df[sheet_name].to_csv(out_folder + 'cleaned/data/' + fname[:-5] + '_' + sheet_name+ '_cleaned.csv', header=None, index=None)

def merge_lastestvalue(fname):
    listing = glob.glob(out_folder + 'cleaned/data/'+ fname[:-5] + '*.csv')
    n = 0
    for filename in listing:
        if n == 0 :
            merge = pd.read_csv(filename)
            n = 1
        else:
            df = pd.read_csv(filename)
            merge = merge.join(df.set_index(['Id']), on=['Id'])
    df_id = pd.read_csv(out_folder + 'cleaned/identifiers/'+ fname[:-5] + '_' +'identifiers.csv')
    df_id['Id'] = df_id.index + 1
    merge = df_id.join(merge.set_index(['Id']), on=['Id'])
    merge.to_csv(out_folder + 'merged/' + fname[:-5] + '_merged.csv', index=False)     

def clean_data_timeseries(df, sheet_name):
    df = clean_err(df)
    start_row, end_row = locate_data(df)
    df = df.loc[start_row:end_row, :]
    for x in range(len(df.columns)):
                        if x == 0:
                            df.rename(columns={ df.columns[0]: 'Date' }, inplace = True)   
                        elif x>0:
                            num = str(x)   
                            df.rename(columns={ df.columns[x]: sheet_name+num }, inplace = True)
    df = pd.wide_to_long(df, stubnames=sheet_name, i='Date', j="Id")
    return df

def clean_multiple_timeseries(fname):
    xlsx = pd.ExcelFile(in_data_folder+fname, engine='openpyxl')
    sheet_names = xlsx.sheet_names
    df = pd.read_excel(in_data_folder+fname, sheet_name=None, engine='openpyxl', header=None)
    df[sheet_names[0]]['Id'] = df[sheet_names[0]].index
    df[sheet_names[0]].rename(columns={ df[sheet_names[0]].columns[0]: df[sheet_names[0]].iloc[0,0] }, inplace = True)
    df[sheet_names[0]] = df[sheet_names[0]].loc[1:, :] 
    df[sheet_names[0]].to_csv(out_folder + 'cleaned/identifiers/'+ fname[:-5] + '_' +'identifiers.csv', index=None)  
    for sheet_name in sheet_names[1:]:
        df[sheet_name] = clean_data_timeseries(df[sheet_name], sheet_name)
        df[sheet_name].to_csv(out_folder + 'cleaned/data/' + fname[:-5] + '_' + sheet_name+ '_cleaned.csv')

def merge_timeseries(fname):
    listing = glob.glob(out_folder + 'cleaned/data/'+ fname[:-5] + '*.csv')
    n = 0
    for filename in listing:
        if n == 0 :
            merge = pd.read_csv(filename)
            n = 1
        else:
            df = pd.read_csv(filename)
            merge = merge.join(df.set_index(['Date','Id']), on=['Date','Id'])
    df_id = pd.read_csv(out_folder + 'cleaned/identifiers/'+ fname[:-5] + '_' +'identifiers.csv')
    df_id['Id'] = df_id.index + 1
    merge = df_id.join(merge.set_index(['Id']), on=['Id'])
    merge.to_csv(out_folder + 'merged/' + fname[:-5] + '_merged.csv', index=False) 

if __name__ == "__main__":
    for n, filename in enumerate(os.listdir(in_data_folder)):
        print("({})Processing {}...".format(str(n),filename))
        if filename != '.DS_Store':
            if 'Value' in filename:
                clean_multiple_lastestvalue(filename)
                merge_lastestvalue(filename)
            else:     
                clean_multiple_timeseries(filename)
                merge_timeseries(filename)
    print('Process Completed...')

