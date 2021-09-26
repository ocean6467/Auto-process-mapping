# -*- coding: utf-8 -*-
"""
Created on Thu Sep 23 22:07:21 2021

@author: user
"""
import os
print(os.getcwd())

import pandas as pd
import numpy as np
from datetime import datetime
df=pd.read_excel('dplot.xlsx',usecols=['AMK_Lot','Device','DP_SAW/C_MC','DP_SAW/C_OP'\
,'DP_SAW/C_In','WaferQty','Z1_SID','Z1_Blade','Z2_SID','Z2_Blade','AOI/2_Yield','PnP_Yield'],nrows=5000)

df['DP_SAW/C_In']=pd.to_datetime(df['DP_SAW/C_In'])
df.sort_values('DP_SAW/C_In',inplace=True)

lotNB=input('Which lot you want to mapping....')
findlot=df.index[df['AMK_Lot']==lotNB][0]
print(findlot)

MCsort=(df.at[findlot,'DP_SAW/C_MC'])
print(MCsort)

filt=(df['DP_SAW/C_MC']==MCsort)
sort=df.loc[filt]
sort.reset_index(drop=True,inplace=True)
print(sort)

findlot2=sort.index[sort['AMK_Lot']==lotNB][0]
print(findlot2)

print(sort.loc[(findlot2-3):(findlot2+4),:])
result=sort.loc[(findlot2-3):(findlot2+4),:]
result.to_excel('process mapping_saw MC.xls')

