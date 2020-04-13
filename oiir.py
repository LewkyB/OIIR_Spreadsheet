"""
1. Figure out what software for each class and insert it into the appropriate cell.
2. Export each data frame as a different page to an .xlsx file
3. Add input prompts so that this can be applied to future sheets
4. What to do with weird course #'s?
5. github
"""


import pandas as pd
# options
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 8)
pd.set_option('display.width', 1000)

# open spreadsheet with pandas
data = pd.ExcelFile('339LabReport.xlsx')

# create data frame
df = data.parse(sheetname='Sheet1', skiprows=7)

# set variable to columns
#col1 = list(df.columns)

# remove empty columns and rename existing ones
df = df.drop(["Unnamed: 0","Unnamed: 3","Unnamed: 5","Unnamed: 7"],
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

print(df_fall.head(50))
print(df_winter.head(50))
print(df_spring.head(50))

#print("\n INFO: \n")
#print(df.info())

