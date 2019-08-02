import os

c = os.popen("ipconfig /all")
print(c.read())
d = os.popen("dir c:\\")
print(c.read())

print(d.read())