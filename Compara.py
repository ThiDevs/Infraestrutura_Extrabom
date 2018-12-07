
arq = open("Ferias.txt",'rt')
lines = arq.readlines()

arq2 = open("Usuario Bloqueado no AD com usuario no Fluig ativo.txt",'rt')
lines2 = arq2.readlines()

dic = {}
for line in lines:
    dic.update({line.lower().strip(): 0})

i = 0
arq3 = open("NewUsers.txt",'wt')
write = ''



for line in lines2:
    if dic.get(line.lower().strip()) != 0:
        print(line)
        write += line.strip() + "\n"
        i += 1
print(i)
arq3.write(write)