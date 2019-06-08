#!/usr/bin/env python
# coding: utf-8

# In[32]:


import pandas as pd

data = pd.read_csv(r"C:\Users\kukumar\Documents\GitHub\LearningPython\Sample_data.csv","~")



#print(data)

#print("\n","type of data","\n")

#print(type(data))

#print(data.loc[data.COST>120000])

#print(data.loc[data.COST>120000])

#for i in range(len(data)):
    #print(data.iloc[i]) 
#    print(data.loc[data.cost>10000])


data2=data.loc[(data.COST>120000) & (data.TITLE == "K14")]

print(data2)


# In[ ]:




