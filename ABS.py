# -*- coding: utf-8 -*-
## ABS
## Haifeng XU
## QQ: 78112407
## Email: hfxu@wxtrust.com

import numpy as np
import pandas as pd
import datetime

df = pd.read_excel('abs.xlsx')
print ( df )

df_date = df.iloc[:,4:]
print ( df_date )

tmp1 = ( df_date.iloc[0,1] - df_date.iloc[0,0] ).days * 100000 - 0.0
tmp1 = format( tmp1, ',')
print ( tmp1 )


## excel 模版是什么 ？

