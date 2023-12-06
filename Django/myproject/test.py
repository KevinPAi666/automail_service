import os

t = '../../automail/my-project-auto-mail-404814.json'

with open(t, 'r') as r:
    for i in r:
        print(i)
print(os.listdir('../../automail/'))
