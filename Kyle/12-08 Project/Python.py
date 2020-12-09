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

df = pd.read_excel('Result - Organization.xlsx')

df = (df.assign(Ocode = df['Ocode'].fillna('nan'),Ccode = df['Ccode'].fillna('nan'))
        .groupby(['Ocode','Ccode'])
        .last()
        .reset_index()
        .replace({'Ocode': {'nan':np.nan}, 'Ccode':{'nan':np.nan}}))

# Initialized

df = (df.groupby(['Ocode'])
        .last()
        .reset_index())

# Write to file

df.to_excel('Out.xlsx', index=False)

# os.remove("output1.xlsx")
