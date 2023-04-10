import os
import pandas as pd
from datetime import date

# Retrieve list of all files in network and delete unwanted files
root = r"\\ugv.corp\seismic_data\well_documents\LGPU"

list_of_dates = []
for path, subdirs, files in os.walk(root):
    for filename in files:
        date_created = date.fromtimestamp(os.path.getmtime(path))
        list_of_dates.append(date_created)

list_of_dates.pop(0)
list_of_dates.pop(0)

# Save all files modification dates
df = pd.DataFrame(list_of_dates)
df.to_excel('dates.xlsx', sheet_name='new_sheet')

# Read excel file with dates
input_file = pd.read_excel('dates.xlsx')
df = pd.DataFrame(input_file)
df.drop(df.columns[0], axis=1, inplace=True)

# Create column names for output file
df.rename(columns={df.columns[0]: 'Дата'}, inplace=True)
column_date = df.columns[0]
column_year = 'Рік'
column_week = 'Номер тижня'
column_files_per_week = 'Кількість файлів за тиждень'
column_files_per_day = 'Кількість файлів за день'

# Parse date and get necessary information
df[column_date] = pd.to_datetime(df[column_date], format='%Y-%m-%d')
df[column_year] = pd.DatetimeIndex(df[column_date]).year
df[column_week] = df[column_date].dt.isocalendar().week
df[column_date] = pd.to_datetime(df[column_date]).dt.date

# Count amount of files per date
df[column_files_per_week] = 0
df[column_files_per_week] = df.groupby([column_year, column_week])[column_date].transform('count')
df.drop_duplicates()

df[column_files_per_day] = 0
columns_result = [column_year, column_week, column_files_per_week, column_date]
df = df.groupby(columns_result)[column_files_per_day].agg('count')

df.to_excel('result.xlsx')
