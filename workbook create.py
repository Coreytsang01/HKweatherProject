#%%
import openpyxl
import os

worksheetname = ["Chek Lap Kok","Cheung Chau","Clear Water Bay","Happy Valley","HK Observatory","HK Park","Kai Tak Runway Park","Kau Sai Chau","King's Park","Kowloon City","Kwun Tong","Lau Fau Shan","Ngong Ping","Pak Tam Chung","Peng Chau","Sai Kung","Sha Tin","Sham Shui Po","Shau Kei Wan","Shek Kong","Sheung Shui","Stanley","Ta Kwu Ling","Tai Lung","Tai Mei Tuk","Tai Mo Shan","Tai Po","Tate's Cairn","The Peak","Tseung Kwan O","Tsing Yi","Tsuen Wan Ho Koon","Tsuen Wan Shing Mun Valley","Tuen Mun","Waglan Island","Wetland Park","Wong Chuk Hang","Wong Tai Sin","Yuen Long Park",
]

#%%
workbook = openpyxl.load_workbook("Master Data.xlsx")

#%%
for worksheet in worksheetname:
    if workbook.sheetnames[worksheet]==False:
        workbook.create_sheet(worksheet)


# %%
workbook.save("Master Data.xlsx")
# %%
