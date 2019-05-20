
#Program 2
#Matrimony eligibility
#Enter Name
#Enter Gender
#Enter Year of Birth (Date of Birth ( Year of Birth ) find out Age)

import datetime

name=input("Enter Name         : ")
gender=input("Enter Gender (M/F) : ")
c=input("Enter Year of Birth: ")
yr=int(c)

#curr=now.year
curr=2019

if(gender == "M"):
    if((curr-yr) >= 21):
        print(name + " is eligible for marriage")
    else:
        print(name + " is NOT eligible for marriage")
elif(gender == "F"):
    if((curr-yr) >= 18):
        print(name + " is eligible for marriage")
    else:
        print(name + " is eligible NOT for marriage")
else:
    print("Invalid Input");
    
