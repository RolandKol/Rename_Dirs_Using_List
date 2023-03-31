__author__ = "R.K."
# a simple script to rename Windows folders
# to use:
# create excel or comma delimited .csv file with columns: 'old_name', 'new_name', 'old_location', 'new_location'
# if 'new_location' of the file or folder is different from 'old_location' items will be renamed and Moved
# otherwise, - if they are the same, - renamed only.
# if any items renaming will fail (for example: in use by other processes, or not found),
# the script will create .csv file with the list of failures.

# my friend JOHN renames Only the copies of his files or folders!
# JOHN is a smart lad!!!
# be like a JOHN

import os
import tkinter as tk
from subprocess import Popen
from time import sleep
from tkinter import filedialog

import pandas as pd
from tqdm import tqdm

root = tk.Tk()
root.withdraw()
default_headers = ['old_name', 'new_name', 'old_location', 'new_location']

file_path = filedialog.askopenfilename(
    title='Selected CSV or Excel file with the list to Rename',
    filetypes=[('All Files', '*.*'), ('CSV files', '*.csv'), ('Excel Files', '*.xlsx')]
)

file_extension = os.path.splitext(file_path)

if len(file_path) ==0:
    print(f'Rename List has NOT been selected')
    print(f'The script was cancelled')
    exit()
elif file_extension[1] == '.csv':
    df = pd.read_csv(file_path, dtype=str)
elif file_extension[1] == '.xlsx' or file_extension[1] == '.xls' or file_extension[1] == '.xlsm':
    df = pd.read_excel(file_path, dtype=str)
else:
    print(f'File Not Supported')
    print(f'The script was cancelled')
    exit()

data_headers = df.columns.tolist()
data_headers = [item.lower() for item in data_headers]
df.columns = data_headers

missing_headers = set(default_headers) - set(data_headers)

if len(missing_headers) != 0:
    print('Please check the list headers!')
    for i in missing_headers:
        print(f'header: "{i}" missing')
    print(f'\nThe Script Cancelled!')
    exit()

missing_cells = df.isnull().values.sum()

if missing_cells!= 0:
    print(f'List has {missing_cells} missing data entries (empty cells)')
    print('Double check the list and do not leave empty cells in the table')
    print(f'\nThe script was cancelled')
    exit()

df['old_address'] = df[['old_location', 'old_name']].agg('\\'.join, axis=1)
df['new_address'] = df[['new_location', 'new_name']].agg('\\'.join, axis=1)

failed_old = []
failed_new = []
failed_error = []


for old, new in tqdm(zip(df.old_address, df.new_address), total=len(df)):
    try:
        os.rename(old, new)
    except Exception as e:
        print(f"Failed to rename: {old} ")
        failed_old.append(old)
        failed_new.append(new)
        failed_error.append(e)

if len(failed_old) != 0 or len(failed_new) != 0:
    df_failed = pd.DataFrame()
    df_failed['Old'] = failed_old
    df_failed['New'] = failed_new
    df_failed['Result'] = 'Failed to rename'
    df_failed['Reason'] = failed_error
    df_failed.to_csv('Failed_to_Rename.csv', index=False)
    print(f'Finished renaming all {len(df)} items in the list.')
    print(f'{len(df_failed)} items Failed')
    sleep(5)
    p = Popen('Failed_to_Rename.csv', shell=True)
else:
    print(f'\nFinished renaming all {len(df)} items in the list.')
    print('No Errors')
    print('Lucky You!')
