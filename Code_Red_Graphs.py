
# coding: utf-8

# In[67]:


import pandas as pd
import numpy as np 
import matplotlib.pyplot as plt
import seaborn as sns
import random
import plotly as py
get_ipython().run_line_magic('matplotlib', 'inline')
from plotly import __version__
from plotly.offline import download_plotlyjs, init_notebook_mode, plot, iplot
import cufflinks as cf
init_notebook_mode(connected=True)
cf.go_offline()
from openpyxl import load_workbook


# In[68]:



pd.set_option('display.height', 500)
pd.set_option('display.max_rows', 500)
wb = load_workbook("CORP Code Red Report 2018.03.01.xlsx")


# In[69]:


sheets = []
for sheet in wb.worksheets:
   
    sheets.append(sheet)


# In[70]:


for sheet in range(0, len(sheets)):
    
    print (sheets[sheet])
    xl = pd.read_excel("CORP Code Red Report 2018.03.01.xlsx", sheet_name=sheet)
    xl.dropna(axis=0, how='all', thresh=None)
    xl1 = xl.fillna(value=0)
    xl2 = xl1.loc['CF after dist']
    xl3 = pd.DataFrame(data=xl.loc['CF after dist'])
    xl4 = xl.loc["CF after dist"].dropna()
    xl4 = xl4.to_frame()
    xl5 = xl.loc["Distributions"].dropna()
    xl5 = xl5.to_frame()
    xl6 = xl.loc["Property Cash"].dropna()
    xl6 = xl6.to_frame()
    xl7 = xl.loc["Total Income"].dropna()
    xl7 = xl7.to_frame()
    xl8 = xl.loc["Total OPEX"].dropna()
    xl8 = xl8.to_frame()
    xl9 = xl.loc["CF after debt"].dropna()
    xl9 = xl9.to_frame()
    mm = joined = xl4.join([xl5,xl6,xl7,xl8,xl9])
    frames = [mm]
    df1 = pd.concat(frames)
    df1.iplot()
    sheet = sheet+1

