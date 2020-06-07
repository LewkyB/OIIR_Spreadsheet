import pandas as pd
import numpy as np

# options
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 8)
pd.set_option('display.width', 1000)

# open spreadsheet with pandas
data = pd.ExcelFile('339LabReport.xlsx')

# create data frame
df = data.parse(sheetname='Sheet1', skiprows=7)

# saving unedited original data to use as last page in workbook 
df_unedited_data = data.parse(sheetname='Sheet1', skiprows=7)

# set variable to columns
col1 = list(df.columns)

# remove empty columns and rename existing ones
df_unedited_data, df = df.drop(["Unnamed: 0","Unnamed: 3","Unnamed: 5","Unnamed: 7"],
             axis=1).iloc[4:,:].rename(columns=
                                       {'Title/Meeting Name':'Class Name',
                                        'Course/Reservation #':'Course #',
                                        'Subject/Customer':'Department',
                                        'Instructor/Contact':'Instructor',
                                        'Date':'Date',
                                        'Software':'Software'
                                        })

# drop all NaN rows
df.dropna(how='all', inplace=True)

# correct data type of Date to datetime64[ns]
df['Date'] = pd.to_datetime(df['Date'])
print(df.info())

# create mask so that data frames only include certain date range
fall_start_date = '2019-08-21'
fall_end_date = '2019-12-14'
mask_fall = (df['Date'] >= fall_start_date) & (df['Date'] <= fall_end_date)

winter_start_date = '2019-12-15'
winter_end_date = '2020-01-11'
mask_winter = (df['Date'] >= winter_start_date) & (df['Date'] <= winter_end_date)

spring_start_date = '2020-01-15'
spring_end_date = '2020-04-16'
mask_spring = (df['Date'] >= spring_start_date) & (df['Date'] <= spring_end_date)

# sort by Department, Course #, and Date
df_fall = df.sort_values(by=['Department', 'Course #', 'Date'], axis=0, ascending=True).loc[mask_fall]
df_winter = df.sort_values(by=['Department', 'Course #', 'Date'], axis=0, ascending=True).loc[mask_winter]
df_spring = df.sort_values(by=['Department', 'Course #', 'Date'], axis=0, ascending=True).loc[mask_spring]

# set names for sheets
dfs = {'Fall_Semester': df_fall,
       'Winter_Semester': df_winter,
       'Spring_Semester': df_spring,
       'Unedited_Data': df_unedited_data}

# writer object
writer = pd.ExcelWriter('OIIR_Astra_Summary.xlsx',
                        engine = 'xlsxwriter',
                        date_format = 'yyyy-mm-dd', # set date format
                        datetime_format = 'yyyy-mm-dd') # set date time format to leave out time

# auto size columns
for sheetname, df in dfs.items():  # loop through `dict` of dataframes
    df.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
        )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

# save file
writer.save()
