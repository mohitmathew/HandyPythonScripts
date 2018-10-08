'''
Created by Mohit Mathew

This script can compare to excel sheets. 
It requires you to provide a key for the table which could be a combination of many columns
Please note that the key should be able to uniquly identify a record and should not have duplpicates values.

'''

import pandas as pd
import numpy as np

def getDataFrame(xlPath, sheetName, HeaderRow,idxcol):
    df = pd.read_excel(xlPath,sheetName,header = headerRow,na_values='')
    df = df.fillna(-9999)
    df.set_index(idxcol,verify_integrity=True,drop=False,inplace=True)
    return df



def dataframe_getMissing(df1, df2, msg,idxcol):
    indx_d1_d2 = df1.index.isin(df2.index)
    df_notin_d2 = df1.loc[~indx_d1_d2][idxcol]
    df_notin_d2['CompareComment'] = msg;
    return df_notin_d2


# inputs
Oldexcel = 'Old.xlsx'
Newexcel = 'New.xlsx'
Sheet = 'LB'
indexCols = ['USUBJID' , 'LBTESTCD' , 'LBCAT' ,'LBDY']
headerRow = 0




print('Reading ' + Oldexcel)
df_Old = getDataFrame(xlPath=Oldexcel,sheetName=Sheet,HeaderRow=headerRow,idxcol=indexCols)

print('Reading ' + Newexcel)
df_New = getDataFrame(xlPath=Newexcel,sheetName=Sheet,HeaderRow=headerRow,idxcol=indexCols)

print('Finding missing in New')
df_notInNew = dataframe_getMissing(df1=df_Old,df2=df_New,msg='Missing in New',idxcol=indexCols)

print('Finding missing in Old')
df_notInOld = dataframe_getMissing(df2=df_Old,df1=df_New,msg='Missing in Old',idxcol=indexCols)

print('Finding data changed data')
cmnindx = df_Old.index.isin(df_New.index)
df_Old = df_Old[cmnindx]
matchResult= df_Old.isin(df_New)

df_DataChange = df_Old[~matchResult]

df_DataChange = df_DataChange.dropna(how='all')
df_DataChange = df_DataChange[indexCols]

df_DataChange['CompareComment'] = 'data changed'


df_allDiff = pd.concat([df_notInNew,df_notInOld,df_DataChange])

#print(df_notInNew.shape)
#print(df_notInOld.shape)
#print(df_DataChange.shape)
#print(df_allDiff.shape)

df_allDiff = df_allDiff[['CompareComment']]

#indx = df_Old.index[notInNew]
if(df_allDiff.shape[0] > 0):
    writer = pd.ExcelWriter('Report.xlsx')
    df_allDiff.to_excel(writer,'Report',index=True)
    writer.save()
    
    print(str(df_allDiff.shape[0]) + ' difference(s) written to Report.xlsx')
    
else:
    print('No difference detected')
    

print('-----------Done-----------')
