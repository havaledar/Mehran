#!/usr/bin/env python
# coding: utf-8

# In[171]:


import pandas as pd
df=pd.read_excel("report.xls", skiprows=5, skipfooter=2, header=0)

#Deleting 3rd row of each participant
df.dropna(subset=['کد ملی'], inplace=True)
df.reset_index(drop=True, inplace=True)

#Creating a df from even rows
df1 = df[df.index%2 == 0]
df1.iloc[0,[8,9,11]]=["Birth_En","Mobile","Name_En"]
df1.columns=df1.iloc[0]
df1=df1.drop(df1.index[0])
df1.reset_index(drop=True, inplace=True)

#Creating a df from odd rows
df2 = df[df.index%2 == 1]
df2.reset_index(drop=True, inplace=True)
df2.columns.values[[6,8,9,11]]=["Prov","Birth_Fa","Nat_ID","Name_Fa"]

#Prearing a neat df
df_concat = pd.concat([df1, df2], axis=1)
df_concat=df_concat.iloc[:,[26,11,21,23,8,9,24]]


# In[181]:


df_dict=df_concat.to_dict(orient='series')

Errors=""

for value in df_dict['Name_Fa'].items():
    try:
        not(str(value[1]).isdigit() and str(value[1]).isascii())
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not Persian." + "\n"
    
for value in df_dict['Name_En'].items():
    try:
        str(value[1]).title().isascii()
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not English." + "\n"
        
for value in df_dict['Prov'].items():
    try:
        not(str(value[1]).isdigit() and str(value[1]).isascii())
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not Persian." + "\n"

for value in df_dict['Birth_Fa'].items():
    try:
        str(value[1]).replace(" ","").isdigit() 
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not a valid date like 1399/01/01." + "\n"

for value in df_dict['Birth_En'].items():
    try:
        str(value[1]).replace(" ","").isdigit()
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not a valid date." + "\n"
        
for value in df_dict['Mobile'].items():
    try:
        str(value[1]).isdigit() and len(str(value[1]))==11
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not a valid Mobile Phone." + "\n"
        
for value in df_dict['Nat_ID'].items():
    try:
        str(value[1]).isdigit() and len(str(value[1]))==10
    except:
        Errors+= str(value[1]) + " in row number " + str(value[0]+1) + " is not a valid National ID." + "\n"


# In[194]:


class Besmela:
    def __init__(self, Dict):
        for k, v in Dict.items():
            setattr(self, k, v)

df_class=Besmela(df_dict)
dir(df_class)


# In[193]:


import sqlite3 as sq
conn= sq.connect('competition.db')
cur= conn.cursor()
cur.execute(""" CREATE TABLE athletes_info5 (Name_Fa, Name_En, Prov, Birth_Fa, Birth_En, Mobile, Nat_ID)""")
cur.execute("""INSERT INTO athletes_info VALUES (?,?,?,?,?,?,?)""", df_class)
conn.com


# In[ ]:


#open window
#RE

