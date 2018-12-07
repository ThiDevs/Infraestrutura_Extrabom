import string
a = string.ascii_uppercase
lst = []
for i in a:
    lst.append(i)
lst2 = []
j = 0
for j in range(5):
    for i in a:
        lst2.append(lst[j]+i)
lst += lst2
print(lst)