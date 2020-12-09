# -*- coding: utf-8 -*-

# Read each Column
# print(df[['Name', 'Type 1', 'HP']])

# Read Each Row
# print(df.iloc[0:4])
# for index, row in df.iterrows():
#     print(index, row['Name'])
# df.loc[df['Type 1'] == "Grass"]

# Read a specific location (R,C)
# print(df.iloc[2,1])
import os
import xlsxwriter
import pandas as pd
import numpy as np

# Pandas and XR
df = pd.read_excel('input.xls')

print(df.columns)

# Concise Code

ends = {'21': 'c21',
        '57': 'c57',
        '58': 'c58',
        '59': 'c59',
        '70': 'c70',
        '71': 'c71',
        '74': 'c74',
        '75': 'c75'
        }

df['Industry Code'] = df['Industry Code'].astype(str)
masko = df['Code'].str.endswith('000') == True
#If Mask statement is true, Code','Corporate Name -> c1','c2
df.loc[masko, ['OC','ON']] = df.loc[masko,['Code','Corporate Name']].to_numpy()
df.loc[~masko, ['CC','CN']] = df.loc[~masko, ['Code','Corporate Name']].to_numpy()

for k, v in ends.items():
    mask = df['Industry Code'].str.endswith(k, na=False)
    df.loc[mask, v] = df.loc[mask, 'End Date']
    # if need append inverse mask
    #df.loc[mask, v] = df.loc[~mask, 'End Date'] 

for k, v in ends.items():
    mask = df['Industry Code'].str.endswith(k, na=False)
    #df.loc[mask, v + ' loc'] = 'Ningbo'
    #df.loc[mask, v + ' loc'] = 'Changshu'
    #df.loc[mask, v + ' loc'] = 'Kuncai'
    df.loc[mask, v + ' loc'] = 'Songjiang'
    #df.loc[mask, v + ' loc'] = 'Huijin'
    #df.loc[mask, v + ' loc'] = 'CDP'
    #df.loc[mask, v + ' loc'] = 'Zhabei'

# Write to column

# dfo[['B', 'C']] = df[['F', 'G']]
# dfo[['B', 'C']] = df[['F', 'G']]

# Write to file

df.to_excel('output1.xlsx', index=False)

# os.remove("output1.xlsx")
# df.to_excel("output1.xlsx")
