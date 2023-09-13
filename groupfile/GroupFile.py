#!/usr/bin/env python
# coding: utf-8


from nptdms import TdmsFile
import os, shutil, re, csv
import pandas as pd
import numpy as np

fullpath = os.getcwd()
fileall = os.listdir(fullpath)
if 'groupfile.xlsx' in fileall:
    os.remove('groupfile.xlsx')
    print('remove previous groupfile.xlsx')
_nsre = re.compile('([0-9]+)')
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(_nsre, s)]

excelfile = []
wt_data = pd.DataFrame()
a5_data = pd.DataFrame()
p_bout= pd.DataFrame()
group_hour = pd.DataFrame()
group_avg = pd.DataFrame()

# remove previous file
for File in os.listdir(fullpath):
        if File.endswith('.xlsm'):
            #print(File)
            excelfile.append(File)

# looping all files to exctract data and group them as dataframe            
for file in excelfile:       
    print("data extracting from : ",file)
    
    # 6 Cycles concating---
    print('6 Cycles data processing---')
    six_Cycles = pd.read_excel(fullpath+'/'+file, sheet_name = '6 Cycles')
    # concat target columns
    wt_data = pd.concat([wt_data,six_Cycles.loc[:,'WT':'Cycle-5']],axis = 1)
    # select rows unitl the empty row appears
    wt_data = wt_data.iloc[:wt_data.isnull().any(axis=1).idxmax(),:]
    a5_data = pd.concat([a5_data,six_Cycles.loc[:,'A5':'Cycle-5.1']],axis = 1)
    a5_data = a5_data.iloc[:a5_data.isnull().any(axis=1).idxmax(),:]
    
    # sleep bouts concating ---
    sleepbouts = pd.read_excel(fullpath+'/'+file, sheet_name = 'Sleep Bouts')
    print('Sleep Bouts data processing---')
    
    # source sheets have no column name,system automatively add column name as 'Unnamed: 0'
    # since one sheet includes too many contents, 
    # find the target content by keyword:Data-1
    id_data1 = sleepbouts[sleepbouts["Unnamed: 0"]== "Data-1"].index.tolist()[0]
    sleepbouts = sleepbouts.iloc[id_data1+1:,:]
    id_data1_end = sleepbouts[sleepbouts["Unnamed: 0"]== "Data-1"].index.tolist()[0]
    #select data with non-null columns and rows
    sleepbouts = sleepbouts.iloc[:id_data1_end-id_data1-1,:]
    sleepbouts = sleepbouts.dropna(axis=1, how = 'all')
    sleepbouts = sleepbouts.dropna(axis=0, how = 'all')
    #print(sleepbouts.shape)    
    p_bout = pd.concat([p_bout,sleepbouts],axis = 1)
    
    # group count data concating ---
    print('groupcount data processing---')
    groupcount = pd.read_excel(fullpath+'/'+file, sheet_name ='GroupCount Data')
    #extract column index
    hour_hour = list(groupcount).index("Hour-1 After Correction by Hour")
    hour_avg = list(groupcount).index('Hour-1 After Correction by Average')
    group_hour = pd.concat([group_hour,groupcount.iloc[:,hour_hour:hour_hour+3]],axis = 1)
    group_avg = pd.concat([group_avg,groupcount.iloc[:,hour_avg:hour_avg+3]],axis = 1)

# final concating after looping done
six_cycles_total = pd.concat([wt_data,a5_data],axis = 1)
group_hour.reset_index(drop = True,inplace = True)
group_avg.reset_index(drop = True,inplace = True)
groupc_total = pd.concat([group_hour,group_avg],axis = 1)

# data frame write to sheets
with pd.ExcelWriter(fullpath +'/groupfile.xlsx', engine = 'xlsxwriter') as writer:
    print('data to sheet processing ---')
    six_cycles_total.to_excel(writer ,sheet_name = '6 Cycles',index = False ,header = True)
    p_bout.to_excel(writer ,sheet_name = 'Sleep Bouts',index = False ,header = False)
    groupc_total.to_excel(writer ,sheet_name = 'Group Count Data',index = False ,header = True)
    
    # set column width
    worksheet = writer.sheets['6 Cycles']
    worksheet.set_column("A:Z", 10)
    worksheet = writer.sheets['Sleep Bouts']
    worksheet.set_column("A:Z", 15)
    worksheet = writer.sheets['Group Count Data']
    worksheet.set_column("A:Z", 15)
    
writer.save()
#writer.close()
print('Groupfile DONE!')




