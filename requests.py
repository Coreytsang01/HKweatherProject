## %%=====================================================================================================

import requests
import os
import pandas as pd
import openpyxl
import numpy as np
import time
import datetime
## %%=====================================================================================================

#Download the files from the web
start_time = time.time()
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
print('Finished downloading the files.')
download_time = round(time.time() - start_time, 2)
print("It uses %s seconds to execute in downloading stage." % download_time )
## %%=====================================================================================================
#   Combine all temporary file into one excel

#   read all files
df_wind = pd.read_csv('temp file\latest_10min_wind.csv')
df_pressure = pd.read_csv('temp file\latest_1min_pressure.csv')
df_temperature = pd.read_csv('temp file\latest_1min_temperature.csv')
df_humidity = pd.read_csv('temp file\latest_1min_humidity.csv')

#   combine into one dataframe
concat_df1 = df_temperature.merge(df_wind, how = 'left', on = ['Date time', 'Automatic Weather Station'])
concat_df2 = concat_df1.merge(df_pressure, how = 'left', on = ['Date time', 'Automatic Weather Station'])
final_combine_df = concat_df2.merge(df_humidity, how = 'left', on = ['Date time', 'Automatic Weather Station'])

#   formatting the date
final_combine_df["Date time"] = final_combine_df["Date time"].astype(str)
date = final_combine_df['Date time']
for i in date:
    x = str(i)
    x1 = [x[6:8],'/',x[4:6],'/',x[:4],' ',x[-4:-2],':',x[-2:]]
    x2 = ''.join(x1)
    #print(x2)
    final_combine_df['Date time'] = final_combine_df['Date time'].replace(to_replace=i, value=x2)

#   reorder the column header
final_combine_df = final_combine_df.reindex(columns=[  'Automatic Weather Station',
                                    'Date time', 
                                    'Air Temperature(degree Celsius)', 
                                    'Relative Humidity(percent)', 
                                    'Mean Sea Level Pressure(hPa)', 
                                    '10-Minute Mean Wind Direction(Compass points)', 
                                    '10-Minute Mean Speed(km/hour)',
                                    '10-Minute Maximum Gust(km/hour)'
                                    ])
final_combine_df = final_combine_df.rename(columns={ 
                                    'Date time':'Date', 
                                    'Air Temperature(degree Celsius)':'Temperature (deg. C)', 
                                    'Relative Humidity(percent)':'Relative Humidity (%)', 
                                    'Mean Sea Level Pressure(hPa)':'Pressure (hPa)', 
                                    '10-Minute Mean Wind Direction(Compass points)':'Wind Direction (Deg.)', 
                                    '10-Minute Mean Speed(km/hour)':'Wind Speed (km/h)',
                                    '10-Minute Maximum Gust(km/hour)':'Maxiumn Gust (km/h)'
                                    })
#print(final_combine_df)
final_combine_df.to_excel("test.xlsx", index = False )

print('Finished combining and cleaning the files.')
process_time = round(time.time() - start_time - download_time, 2)
print("It uses %s seconds to execute in downloading stage." % process_time)
## %%=====================================================================================================

# Insert data into master data.xlsx

combine_excel_df = pd.read_excel('test.xlsx')
workbook_Master = openpyxl.load_workbook('Master Data.xlsx')
#print(combine_excel)

# number of station: 39
Station_location = [
"Chek Lap Kok","Cheung Chau","Clear Water Bay",
"Happy Valley","HK Observatory","HK Park",
"Kai Tak Runway Park","Kau Sai Chau","King's Park",
"Kowloon City","Kwun Tong","Lau Fau Shan",
"Ngong Ping","Pak Tam Chung","Peng Chau",
"Sai Kung","Sha Tin","Sham Shui Po",
"Shau Kei Wan","Shek Kong","Sheung Shui",
"Stanley","Ta Kwu Ling","Tai Lung",
"Tai Mei Tuk","Tai Mo Shan","Tai Po",
"Tate's Cairn","The Peak","Tseung Kwan O",
"Tsing Yi","Tsuen Wan Ho Koon","Tsuen Wan Shing Mun Valley",
"Tuen Mun","Waglan Island","Wetland Park",
"Wong Chuk Hang","Wong Tai Sin","Yuen Long Park"
]
workbook_Master = openpyxl.load_workbook('Master Data.xlsx')

for locations in combine_excel_df['Automatic Weather Station']:
    selected_row = combine_excel_df.loc[combine_excel_df['Automatic Weather Station'] == locations]
    selected_row = selected_row.drop('Automatic Weather Station', inplace=False, axis=1)    #cancel the location column
    #print(selected_row)


    workbook_Master.active = workbook_Master[locations]
    df_workbook_Master = pd.read_excel('Master Data.xlsx',sheet_name=locations)

    #print(df_workbook_Master)

    concated_df = pd.concat([df_workbook_Master,selected_row], axis=0)
    concated_df = concated_df.drop_duplicates(subset=['Date'])

    with pd.ExcelWriter(
    "Master Data.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
    ) as writer:
        concated_df.to_excel(writer, sheet_name=locations,index=False)
print('Finished transforming the data from temporary files to Master Data.xlsx')

transforming_time = round(time.time() - start_time - download_time - process_time, 2)
print("It uses %s seconds to execute in transforming stage." % transforming_time)


# Get Now time
now = datetime.datetime.now()
currentTime = now.strftime("%m/%d/%Y, %H:%M:%S")
txt = 'Previous Update Time: ' + str(currentTime)

# change to df
df_time = pd.DataFrame([txt], index=['UpdateTime'])

# save to csv
df_time.to_csv('log.csv', header=False)
## %%=====================================================================================================

