# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
#Remove item Null Handling
"""

def main():
    print("\n1.Add Items")
    print("2.Remove Item")
    print("3.Update Item")
    print("4.Display Item")
    print("5.Exit")
    choice = int(input())
    #print(choice)
    return choice;

def add_items():
    r = int(input("Enter Number of items to add with the cost : \n"))
    for i in range(r):
        prod = input("Enter product name ")
        value = input ("ENter the cost of product " + prod)
        elem[prod] = value
        

def remove_item():
    item = input("Enter the element to remove from the list : \n")
    if(item in elem):
        del elem[item]
    else:
        print("Item doesnt exist")


def update_item():
    prod = input("ENter the product name to update : ")
    value = input ("Enter the value : ")
    if prod in elem:
        print ("Item Exists updating the value")
        elem[prod] = value
    else:
        print ("Item doesnt exist, Appending to the list : ")
        elem[prod] = value
    
def display():
    #for i in range(len(elem)):
        #print("Cost of ", elem.keys," is ",elem[""])
    for key in elem.keys():
        print(key)
    
    print(elem)
            
choice = 1
elem = {}

while choice in (1,2,3,4):
    choice = main()
    if choice == 1:
        add_items()
    elif choice == 2:
        remove_item()
    elif choice == 3:
        update_item()
    elif choice == 4:
        display()
    else:
        print("Thank you")


"""

for i in range(len(slist)):
    slist = l1[i]
    if (prod in slist):
        operations;

"""

