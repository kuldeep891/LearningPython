from urllib.request import urlopen as uo

with uo('http://sixty-north.com/c/t.txt') as story:
    words = []
    for line in story:
        
