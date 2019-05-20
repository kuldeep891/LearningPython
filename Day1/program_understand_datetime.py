import datetime


a=input("Enter Date in YYYY-MM-DD format")
yr,month,day = map(int, a.split('-'))
dt1=datetime.date(yr,month,day)
print(dt1)
print(type(dt1))


now = datetime.datetime.now()

print(now)
#2019-05-13 12:42:49.641991

print(type(now))

print(now.year)
#2019

print(type(now.year))
#<class 'int'>

print(now.strftime("%A"))
#Monday

x=datetime.datetime(2019,5,20,)
print(x)
#2019-05-20 00:00:00

print(now.strftime("%B"))
#May



