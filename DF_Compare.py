'''
Created by Mohit Mathew

This script can compare to excel sheets. It requires you to provide a key for the table which 
could be a combination of many columns
Please note that the key should be able to uniquly identify a record and should not lead to duplpicates.

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


# inputs --------------------------------
leftexcel = 'Left.xlsx'
rightexcel = 'Right.xlsx'
Sheet = 'LB'
indexCols = ['USUBJID' , 'LBTESTCD' , 'LBCAT' ,'LBDY']
headerRow = 0
# inputs --------------------------------


print('Reading ' + leftexcel)
df_left = getDataFrame(xlPath=leftexcel,sheetName=Sheet,HeaderRow=headerRow,idxcol=indexCols)

print('Reading ' + rightexcel)
df_right = getDataFrame(xlPath=rightexcel,sheetName=Sheet,HeaderRow=headerRow,idxcol=indexCols)

print('Finding missing in right')
df_notInRight = dataframe_getMissing(df1=df_left,df2=df_right,msg='Missing in right',idxcol=indexCols)

print('Finding missing in left')
df_notInLeft = dataframe_getMissing(df2=df_left,df1=df_right,msg='Missing in left',idxcol=indexCols)

print('Finding data changed data')
cmnindx = df_left.index.isin(df_right.index)
df_left = df_left[cmnindx]
matchResult= df_left.isin(df_right)

df_DataChange = df_left[~matchResult]

df_DataChange = df_DataChange.dropna(how='all')
df_DataChange = df_DataChange[indexCols]

df_DataChange['CompareComment'] = 'data changed'


df_allDiff = pd.concat([df_notInRight,df_notInLeft,df_DataChange])

#print(df_notInRight.shape)
#print(df_notInLeft.shape)
#print(df_DataChange.shape)
#print(df_allDiff.shape)

df_allDiff = df_allDiff[['CompareComment']]

#indx = df_left.index[notInRight]
if(df_allDiff.shape[0] > 0):
    writer = pd.ExcelWriter('Report.xlsx')
    df_allDiff.to_excel(writer,'Report',index=True)
    writer.save()
    
    print(str(df_allDiff.shape[0]) + ' difference(s) written to Report.xlsx')
    
else:
    print('No difference detected')
    

print('-----------Done-----------')
