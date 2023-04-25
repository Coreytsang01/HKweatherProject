#%%
import requests
import os
import pandas as pd
import openpyxl
import csv

#%%
downloadUrl = [
'https://data.weather.gov.hk/weatherAPI/hko_data/regional-weather/latest_10min_wind.csv', 
'https://data.weather.gov.hk/weatherAPI/hko_data/regional-weather/latest_1min_temperature.csv',
'https://data.weather.gov.hk/weatherAPI/hko_data/regional-weather/latest_1min_pressure.csv',
'https://data.weather.gov.hk/weatherAPI/hko_data/regional-weather/latest_1min_humidity.csv'
]
downloadfile_directory = 'temp file'

for url in downloadUrl:
    response = requests.get(url)
    if response.status_code == 200:
        file_path = os.path.join(downloadfile_directory, os.path.basename(url))
        with open(file_path, 'wb') as f:
            f.write(response.content)

#%%

workbook_Master = openpyxl.load_workbook('Master Data.xlsx')
workbook_Master.sheetnames[0]
df_workbook_Master = pd.read_excel('Master Data.xlsx')
print(df_workbook_Master)

df_wind = pd.read_csv('temp file\latest_1min_pressure.csv')
print(df_wind)
wind_column_list = list(df_wind.columns)
print(wind_column_list)

# %%
#   Date cleaning
wind_data_cols = [
    'Date',
    'Location',
    'Wind Direction',
    'Wind Mean Speed',
    'Maximum Gust'
]
new_df_wind = df[wind_data_cols]

date = df_wind['Date time']
for i in date:
    x = str(i)
    x1 = [x[6:8],'/',x[4:6],'/',x[:4],' ',x[-4:-2],':',x[-2:]]
    x1 = ''.join(x)
    print(x1)
# %%
df_wind = pd.read_csv('temp file\latest_10min_wind.csv')
df_pressure = pd.read_csv('temp file\latest_1min_pressure.csv')
df_temperature = pd.read_csv('temp file\latest_1min_temperature.csv')
df_humidity = pd.read_csv('temp file\latest_1min_humidity')
concat_df = df_temperature.merge(df_wind, how = 'left', on = ['Date time', 'Automatic Weather Station'])

print(concat_df)
concat_df.to_excel("test.xlsx", index = False )
# %%
