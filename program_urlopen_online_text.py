#try running from home network
#Error when working in office

from urllib.request import urlopen

#with urlopen("C:\Users\kukumar\Documents\GitHub\LearningPython\c.txt") as story:

#Below doesnt work since the escape character \U starts an eight character Unicode Escape
#Hence raw string is used with prefix r before quotes
#with open(r"C:\Users\kukumar\Documents\GitHub\LearningPython\c.txt","r") as story:


def fetch_words():
    with open(r"C:\Users\kukumar\Documents\GitHub\LearningPython\c.txt","r") as story:
        words = []
        for line in story:
            line_words = line.split()
            for word in line_words:
                words.append(word)
    for word in words:
        print(word)
    
fetch_words()
