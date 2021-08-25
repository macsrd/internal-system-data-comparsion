from tkinter import filedialog
import os
import pandas as pd
import numpy as np 
import time

# script comparing if inbound data ties to system

inputFolder = os.getcwd()

# importing source files
filename   = filedialog.askopenfilename(title='Select CSV Export File', filetypes=(('csv files', '*csv'), ('all files', '*.*')), initialdir = inputFolder)
filename_2 = filedialog.askopenfilename(title='Select System Export File', filetypes=(('xlsx files', '*xlsx'), ('all files', '*.*')), initialdir = inputFolder)

df   = pd.read_csv(filename, index_col=None, header=0)
df_2 = pd.read_csv(filename_2, index_col=None, header=0)

# creating new dataframe with Ref IDs extracted from Status column
df_3 = pd.Dataframe()
df_3.insert(0, 'CSV', '')
df_3['CSV'] = df['Order ID'].str.split(',').apply(pd.Series, 1).stack()
df_3['CSV'].replace('', np.nan, inplace=True)
df_3.dropna(subset=['CSV'], inplace=True)

# checking if transactions from inbound are booked into system
df_3['Match'] = df_3['CSV'].isin(df_2['Ref ID'])
df_3 = df_3.rename(columns=('CSV': 'Ref ID'))
df_3 = pd.merge(df_3, df_2[['Ref ID', 'Last Update User']], on='Ref ID', how='left')
df_3 = df_3.rename(columns=('Ref ID': 'CSV'))
df_3.loc[df_3['Last Update User'] == '', 'Last Update User'] = 'Check'

# exporting final file
ExcelWriter = pd.ExcelWriter
timestr = time.strftime("%Y%m%d-%H%M%S")
file = ('Analyze_' + timestr + '.xlsx')

with ExcelWriter(file) as writer:
	df.to_excel(writer, sheet_name='CSV', index=False)
	df_2.to_excel(writer, sheet_name='CRIMS', index=False)
	df_3.to_excel(writer, sheet_name='Compare', index=False)