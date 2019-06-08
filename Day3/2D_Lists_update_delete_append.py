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
        new = []
        item = input("Enter the Item Name : \n")
        new.append(item)
        #elem[i].append(item)
        cost = float(input("Enter the cost for "+ item +"  :\n"))
        new.append(cost)
        #elem[i+1].append(cost)
        elem.append(new)
        #print(elem)
        #print(elem[i][i+1])

def remove_item():
    item = input("Enter the element to remove from the list : \n")
    flag = 0
    #length = len(elem)
    #print(len(elem))
    for i in range(len(elem)):
        #print(elem[i-1])
        if item in elem[i-1][0]:
            elem.remove(elem[i-1])
        else:
            flag += 1
    
    #if flag == length-1:
    #    print("Item doesnt Exist : IF")

    for i in range(len(elem)):
        slist = elem[i]
        if(item not in slist):
            flag = 0
        else:
            flag = 1

    if (flag == 0):
        print ("Item Doesnt Exist")

def update_item():
    choice = int(input("Enter if you want to update : \n 1.Item \n 2.Cost\n 3.Both \n"))
    #length = len(elem)
    if choice == 1:
        item = input("Enter the Item to update : ")
        final_item = input("Enter the Item to updated Value : ")
        for i in range(len(elem)):
            if item in elem[i-1][0]:
                elem[i-1][0] = final_item
    elif choice == 2:
        item = input("Enter the Item to update : ")
        cost = input("Enter the updated cost : ")
        for i in range(len(elem)):
            if item in elem[i-1][0]:
                elem[i-1][1] = float(cost)
    elif choice == 3:
        item = input("Enter the Item to update : ")
        new_item = input("Enter the new Item Name : ")
        cost = input("Enter the new Items cost : ")
        for i in range(len(elem)):
            if item in elem[i-1][0]:
                elem[i-1][0] = new_item
                elem[i-1][1] = float(cost)
    else:
        print ("Invalid Choice")
    
def display():
    print(elem)
    for i in range(len(elem)):
        print("Element at position ",i," is Item : ",elem[i][0],"\t with Cost :",elem[i][1])
    #print(elem[1])
            
choice = 1
elem = list()

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

