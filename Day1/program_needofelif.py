#Check use of if elif and why elif is required

print("Enter a number : ")
a=int(input())

if (a>0):
    print("a is positive")
else:
    print("a is negative")
print("a is analysed")


#if(a>0):
#    print("Positive")
#if(a<0):
#    print("Negative")
#if(a==0):
#    print("Neutral")


if(a>0):
    print("Positive")
elif(a<0):
    print("Negative")
elif(a==0):
    print("Neutral")
else:
    print("Invalid Input")
