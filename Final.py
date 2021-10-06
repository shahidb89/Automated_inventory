#!/usr/bin/env python
# coding: utf-8

# In[6]:


import datetime
import pandas as pd
import numpy as np

bmt_df = pd.read_excel('bmt_pharmacy.xls')
drug_store_df = pd.read_excel('drug_store.xls')
estimation_df = pd.read_excel('estimation_work.xlsx')
pharmacy_df = pd.read_excel('hiwa_pharmacy.xls')
in_patient_df = pd.read_excel('in_patient.xls')
out_patient_df = pd.read_excel('out_patient.xls')

# Clean dataframes:

bmt_df = bmt_df.iloc[:,2:7]
drug_store_df = drug_store_df.iloc[:,2:7]
pharmacy_df = pharmacy_df.iloc[:,2:7]
in_patient_df = in_patient_df.iloc[:,2:7]
out_patient_df = out_patient_df.iloc[:,2:7]

list_of_dfs = [bmt_df, drug_store_df, pharmacy_df, in_patient_df, out_patient_df]

for i, ls in enumerate(list_of_dfs):
    if i == 0:
        ls["Source"] = "bmt"
    elif i == 1:
        ls["Source"] = "drug_store"
    elif i == 2:
        ls['Source'] = "pharmacy"
    elif i == 3:
        ls ['Source'] = "in_patient"
    elif i == 4:
        ls['Source'] = "out_patient"


for index, row in bmt_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

for index, row in pharmacy_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

for index, row in in_patient_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)
    
for index, row in out_patient_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

    
drug_store_df['Expiry Date'] = pd.to_datetime(drug_store_df['Expiry Date'])

drug_store_df.head(5)

df_new = drug_store_df.groupby(['Item Name', 'Expiry Date'])['Quantity in Stock'].sum()




df_new.to_csv('all_in_one.csv')

all_in_one_df = pd.read_csv('all_in_one.csv')

# estimation_work_df = estimation_df[['Item Name', 'AVG QTY Dispensed Daily']]
estimation_df ['system est'] = estimation_df ['system est']/30
estimation_df ['manual est'] = estimation_df ['manual est']/30

work_df = pd.merge(all_in_one_df, estimation_df, on = 'Item Name', how = 'left')

expiry_list = all_in_one_df['Expiry Date'].tolist()

work_df['Expiry Date'] = pd.Series(expiry_list)

today = pd.to_datetime("today")

work_df['today'] = today

work_df['Expiry Date'] = pd.to_datetime(work_df['Expiry Date'])

work_df['Difference'] = (work_df['Expiry Date'] - work_df['today']).dt.days

work_df_sys = work_df
work_df_man = work_df
# system dataframe

work_df_sys['supply in months sys'] = np.nan
work_df_sys['total supply in months sys'] = np.nan
work_df_sys['will expire sys'] = np.nan

for index, row in work_df_sys.iterrows():
    a = work_df_sys.loc[index, 'Quantity in Stock'] / work_df_sys.loc[index, 'system est']
    b = work_df_sys.loc[index, 'Difference'] - a
    
    if b < 0:
        work_df_sys.loc[index, 'will expire sys'] = b * -1
        work_df_sys.loc[index, 'supply in months sys'] = work_df_sys.loc[index, 'Difference']
        work_df_sys.loc[index, 'total supply in months sys'] = work_df_sys.loc[index, 'Difference']
    else:
        work_df_sys.loc[index, 'supply in months sys'] = a
        work_df_sys.loc[index, 'total supply in months sys'] = a

for index, row in work_df_sys.iterrows():
    if work_df_sys.iloc[index, 0] == work_df_sys.iloc[index-1, 0]:
        work_df_sys.iloc[index, 7] = work_df_sys.iloc[index, 7] - work_df_sys.iloc[index-1, 9]
        a = work_df_sys.loc[index, 'Quantity in Stock'] / work_df_sys.loc[index, 'system est']
        b = work_df_sys.loc[index, 'Difference'] - a
    
        if b < 0:
            work_df_sys.loc[index, 'will expire sys'] = b * -1
            work_df_sys.loc[index, 'supply in months sys'] = work_df_sys.loc[index, 'Difference'] 
            work_df_sys.loc[index, 'total supply in months sys'] = work_df_sys.loc[index, 'Difference'] + work_df_sys.loc[index-1, 'total supply in months sys']
        else:
            work_df_sys.loc[index, 'supply in months sys'] = a
            work_df_sys.loc[index, 'total supply in months sys'] = a + work_df_sys.loc[index-1, 'total supply in months sys']


work_df_sys['will expire sys'] = work_df_sys['will expire sys'] * work_df_sys['system est']


work_df_sys['supply in months sys'] = work_df_sys['supply in months sys']/30
work_df_sys['total supply in months sys'] = work_df_sys['total supply in months sys']/30
work_df_sys = work_df_sys.round(1)
work_df_sys['will expire sys'] = work_df_sys['will expire sys'].round(0)

# Manual estimation dataframe

bmt_df = pd.read_excel('bmt_pharmacy.xls')
drug_store_df = pd.read_excel('drug_store.xls')
estimation_df = pd.read_excel('estimation_work.xlsx')
pharmacy_df = pd.read_excel('hiwa_pharmacy.xls')
in_patient_df = pd.read_excel('in_patient.xls')
out_patient_df = pd.read_excel('out_patient.xls')

# Clean dataframes:

bmt_df = bmt_df.iloc[:,2:7]
drug_store_df = drug_store_df.iloc[:,2:7]
pharmacy_df = pharmacy_df.iloc[:,2:7]
in_patient_df = in_patient_df.iloc[:,2:7]
out_patient_df = out_patient_df.iloc[:,2:7]

list_of_dfs = [bmt_df, drug_store_df, pharmacy_df, in_patient_df, out_patient_df]

for i, ls in enumerate(list_of_dfs):
    if i == 0:
        ls["Source"] = "bmt"
    elif i == 1:
        ls["Source"] = "drug_store"
    elif i == 2:
        ls['Source'] = "pharmacy"
    elif i == 3:
        ls ['Source'] = "in_patient"
    elif i == 4:
        ls['Source'] = "out_patient"


for index, row in bmt_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

for index, row in pharmacy_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

for index, row in in_patient_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)
    
for index, row in out_patient_df.iterrows():
    drug_store_df = drug_store_df.append(pd.Series(row, index=drug_store_df.columns), ignore_index=True)

    
drug_store_df['Expiry Date'] = pd.to_datetime(drug_store_df['Expiry Date'])

df_new = drug_store_df.groupby(['Item Name', 'Expiry Date'])['Quantity in Stock'].sum()

drug_store_df.head(5)


df_new.to_csv('all_in_one.csv')

all_in_one_df = pd.read_csv('all_in_one.csv')

# estimation_work_df = estimation_df[['Item Name', 'AVG QTY Dispensed Daily']]
estimation_df ['system est'] = estimation_df ['system est']/30
estimation_df ['manual est'] = estimation_df ['manual est']/30

work_df = pd.merge(all_in_one_df, estimation_df, on = 'Item Name', how = 'left')

expiry_list = all_in_one_df['Expiry Date'].tolist()

work_df['Expiry Date'] = pd.Series(expiry_list)

today = pd.to_datetime("today")

work_df['today'] = today

work_df['Expiry Date'] = pd.to_datetime(work_df['Expiry Date'])

work_df['Difference'] = (work_df['Expiry Date'] - work_df['today']).dt.days


work_df_man = work_df


work_df_man['supply in months man'] = np.nan
work_df_man['total supply in months man'] = np.nan
work_df_man['will expire man'] = np.nan

for index, row in work_df_man.iterrows():
    a = work_df_man.loc[index, 'Quantity in Stock'] / work_df_man.loc[index, 'manual est']
    b = work_df_man.loc[index, 'Difference'] - a
    
    if b < 0:
        work_df_man.loc[index, 'will expire man'] = b * -1
        work_df_man.loc[index, 'supply in months man'] = work_df_man.loc[index, 'Difference']
        work_df_man.loc[index, 'total supply in months man'] = work_df_man.loc[index, 'Difference']
    else:
        work_df_man.loc[index, 'supply in months man'] = a
        work_df_man.loc[index, 'total supply in months man'] = a

for index, row in work_df_man.iterrows():
    if work_df_man.iloc[index, 0] == work_df_man.iloc[index-1, 0]:
        work_df_man.iloc[index, 7] = work_df_man.iloc[index, 7] - work_df_man.iloc[index-1, 9]
        a = work_df_man.loc[index, 'Quantity in Stock'] / work_df_man.loc[index, 'manual est']
        b = work_df_man.loc[index, 'Difference'] - a
    
        if b < 0:
            work_df_man.loc[index, 'will expire man'] = b * -1
            work_df_man.loc[index, 'supply in months man'] = work_df_man.loc[index, 'Difference']
            work_df_man.loc[index, 'total supply in months man'] = work_df_man.loc[index, 'Difference'] + work_df_man.loc[index-1, 'total supply in months man']
        else:
            work_df_man.loc[index, 'supply in months man'] = a
            work_df_man.loc[index, 'total supply in months man'] = a + work_df_man.loc[index-1, 'total supply in months man']
            


work_df_man['will expire man'] = work_df_man['will expire man'] * work_df_man['manual est'] 
work_df_man['manual est'] = work_df_man['manual est'] * 30

work_df_man['supply in months man'] = work_df_man['supply in months man']/30
work_df_man['total supply in months man'] = work_df_man['total supply in months man']/30
work_df_man = work_df_man.round(1)
work_df_man['will expire man'] = work_df_man['will expire man'].round(0)


# Final dataframes


f_df = work_df_sys
f_df['supply in months man'] = work_df_man['supply in months man']
f_df['total supply in months man']= work_df_man['total supply in months man']
f_df['will expire man'] = work_df_man['will expire man']


estimation_df_2 = pd.read_csv('estimation_work.csv')
final_df = union = pd.merge(f_df, estimation_df_2, how='outer',on=['Item Name', 'Item Name'])

edit_df_man = final_df.groupby(['Item Name'])['supply in months man'].sum()
edit_df_sys = final_df.groupby(['Item Name'])['supply in months sys'].sum()

edit_df_man.to_csv('edit_df_man.csv')
edit_df_sys.to_csv('edit_df_sys.csv')

edit_df_man_1 = pd.read_csv('edit_df_man.csv')
edit_df_sys_1 = pd.read_csv('edit_df_sys.csv')

edit_df_man_1 = edit_df_man_1.rename(columns={'supply in months man': 'supply_man_months Grand total'})
edit_df_sys_1 = edit_df_sys_1.rename(columns={'supply in months sys': 'supply_sys_months Grand total'})

final_df = union = pd.merge(final_df, edit_df_sys_1, how='outer',on=['Item Name', 'Item Name'])
final_df = union = pd.merge(final_df, edit_df_man_1, how='outer',on=['Item Name', 'Item Name'])

edit_qty_df = final_df.groupby(['Item Name'])['Quantity in Stock'].sum()


edit_qty_df.to_csv('edit_qty.csv')
edit_qty_df_1 = pd.read_csv('edit_qty.csv')
edit_qty_df_1 = edit_qty_df_1.rename(columns={'Quantity in Stock': 'Quantity Grand total'})
final_df = union = pd.merge(final_df, edit_qty_df_1, how='outer',on=['Item Name', 'Item Name'])

final_df.to_csv('FINAL.csv')
final_df.to_html('FINAL.html')
import os
os.remove("all_in_one.csv")
os.remove("edit_df_man.csv")
os.remove("edit_df_sys.csv")
os.remove('edit_qty.csv')
