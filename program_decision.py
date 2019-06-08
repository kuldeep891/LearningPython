#!/usr/bin/env python
# coding: utf-8

# In[3]:


#You have to make the decision making program which will ask user the answers of 4 questions..
#And based on answers of 4 questions, user have to make the combination of conditions to decide 
#whether a child will go out for playing or not.

#Question can be of this type:

#1. Are you physically fit?
#2. Is it raining outside or sunny?
#3. Are u interested to play football?
#4. Do u have friends outside to play?

#Above are example questions u can make use of  them or use ur own questions

print('Are you Physically fit ? {Y/N}')
fit=input()

print('Is it raining outside? {Y/N}')
rain=input()

print ('Are you interested in playing football ? {Y/N}')
fb = input()

print('Do you have friends interested in playing football ? {Y/N}')
frnd=input()



def outcome( fit,rain,fb,frnd ):
    if(fit == 'Y'):
        if( rain == 'Y' ):
            return "Cant Play due to Rain"
    
str=outcome(fit,rain,fb,frnd)
print(str)    


# In[ ]:




