
from quicktranslate import get_translate_baidu as trans



filmpath="1.txt"
fin=open(filmpath,mode='r')
lines=fin.readlines()
print(lines)
for line in lines:
    print(line," :",trans(line.strip()))
