#ARRAY & List
#Memory allocation is static/fixed

#List()
#Datatype is heterogeneous
#Memory Allocation is dynamic

n=5
groclist=[]
print(type(groclist))
#<class 'list'>

i=1
while i<=5:
    elem = input("Enter grocery elements : ")
    groclist.append(elem)
    i+=1
    
#list can have heterogeneous data
#But since it PY3.7 we have to explicitly define the datatype
groclist.append(123)

#groclist[10]="Kuldeep"
#above be "IndexError: list assignment index out of range"

print("\nGrocery list as per input ")
print(groclist)

groclist[3] = "Wheat"
print("\n Changing 3rd value to Wheat:  ")
print(groclist)

#Remove any item from the list
toremove = input("\nInput which item to remove : ")
if(toremove in groclist):
    groclist.remove(toremove)
else:
    print("\n",toremove," doest exist in the list ")
print("\nGrocery list after removing a value ")
print(groclist)

#If any value given out of list will throw an error.
#groclist.remove(toremove)
#ValueError: list.remove(x): x not in list


i=0
for item in groclist:
    #print (item)
    if(item == 'Wheat'):
        groclist[i] = 'Rice'
    i+=1

print("\nGrocery list after all operations ")
print(groclist)



#to find the index of a value
# groclist.indexof("Rice")
