#Input 5 values
#Name
#Address
#Contact
#Gender
#DateofBirth
#Eligibility DL, if eligible
#    print your GEnder and you are eligible

#Take input of 5 people
#print number of males and number of females

import datetime


male=1
print("Males number ",male)

now = datetime.datetime.now()
curr=now.year
#print(curr)

#num_arr = list()
n = list()
a = list()
c = list()
g = list()
d = list()

i=0
male=0
fm=0
other=0
while i<2:
    n1=input("Enter Name      : ")
    a1=input("Enter Address   : ")
    c1=input("Enter Contact   : ")
    g1=input("Enter Gender    : ")
    d1=input("Enter DOB in YYYY-MM-DD format : ")
#   num_arr.append(n1)
    n.append(n1)
    a.append(a1)
    c.append(c1)
    g.append(g1)
    yr1,mn1,day1 = map(int,d1.split('-'))
    d1=datetime.date(yr1,mn1,day1);
    d.append(d1)
    if(g1 == 'M'):
        male+=1
    elif(g1 == 'F'):
        fm+=1
    else:
        other+=1
    if(curr-d1.year >=18):
        print(n1+" is eligible for DL with contact "+c1)
    else:
        print("Not applicable")
    i+=1

print("Number of Males : ",male)
print("Number of Females : ",fm)
print("Number of Others : ",other)

    






