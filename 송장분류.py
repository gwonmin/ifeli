#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np
import os
import time

# In[3]:


base_dir = 'C:/Users/user/Desktop'
excel_file = 'ESA005E.xlsx'
excel_dir = os.path.join(base_dir, excel_file)

df = pd.read_excel(excel_dir, engine='openpyxl', skiprows = [0], 
                   converters = {'주문번호':str, '묶음주문번호':str})

df = df.drop(df.index[df['주문번호'].count()])

single_list = []
multi_list = []

grouped = df.groupby(['묶음주문번호'])
for i, j in grouped:
    if j['주문번호'].count() == 1:
        single_list.append(i)
    else:
        multi_list.append(i)


# In[4]:



df1 = df.copy() 

df1 = df1[df1['묶음주문번호'].isin(single_list)]

df1 = df1.sort_values(['품목코드(ERP)','수량'], ascending = True)


# In[5]:



df1_1 = df1.copy()

df1_num_1 = df1_1[df1_1['수량'] < 2] # 수량 1개만 분류
df1_else = df1_1[df1_1['수량'] >= 2] # 나머지

df1_num_1 = df1_num_1.sort_values(['품목코드(ERP)'], ascending = True)

df1_else = df1_else.sort_values(['품목코드(ERP)', '수량'], ascending = True)

df1 = pd.merge(df1_num_1, df1_else, how = 'outer')


# In[6]:



df2 = df.copy()

df2 = df2[df2['묶음주문번호'].isin(multi_list)]

df2 = df2.sort_values(['묶음주문번호','품목코드(ERP)'], ascending = True)


# In[7]:


df2_1 = df2.copy().reset_index()

a = df2_1.groupby(['묶음주문번호'])
df2_2_list = []
df2_else_list = []

for i,j in a:
        if j['주문번호'].count() == 2:
            df2_2_list.append(i)
        else:
            df2_else_list.append(i)


# In[8]:


df2_2 = df2_1[df2_1['묶음주문번호'].isin(df2_2_list)]

df2_else = df2_1[df2_1['묶음주문번호'].isin(df2_else_list)]


# In[78]:


NL_code = ['A0001', 'A0002']
NL_list = []
else_list = []

for i in range(df2_2['묶음주문번호'].count()):
    sort_list = []
    if i%2 == 0:
        sort_list.append(df2_2.iloc[i]['품목코드(ERP)'])
        sort_list.append(df2_2.iloc[i+1]['품목코드(ERP)'])
        if 'A0001' in sort_list and 'A0002' in sort_list:
            NL_list.append(df2_2.iloc[i]['묶음주문번호'])
        else:
            else_list.append(df2_2.iloc[i]['묶음주문번호'])
    else:
        else_list.append(df2_2.iloc[i]['묶음주문번호'])


# In[84]:


df2_2_else = df2_2[df2_2['묶음주문번호'].isin(else_list)]

df2_2_NL = df2_2[df2_2['묶음주문번호'].isin(NL_list)]


# In[86]:


df2_2_NL.sort_values(['묶음주문번호', '수량' ], ascending = True)


# In[88]:


df2_2 = pd.merge(df2_2_NL, df2_2_else, how = 'outer')


# In[90]:


df2 = pd.merge(df2_2, df2_else, how = 'outer')


# In[91]:


result = pd.merge(df1, df2, how = 'outer')
result = pd.merge(result, df, how = 'outer')


# In[92]:


del result['index']


# In[93]:


now = time.strftime('%Y.%m.%d', time.localtime(time.time()))

base_dir = 'C:/Users/user/Desktop'
file_name = '%s.xlsx'%(now)
xlsx_dir = os.path.join(base_dir, file_name)
result.to_excel(xlsx_dir, engine = 'xlsxwriter', index=False)

