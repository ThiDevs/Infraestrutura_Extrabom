arq = open("UsersAD2.txt", "rt", encoding="utf-8")
lines = arq.readlines()
a = 3

dic = {}
for i in range(len(lines)):
    try:
        if lines[a].split(":")[1].strip() == 'False':
            dic.update({lines[a + 2 + 3].split(":")[1][1:].strip(): 0})
            print(lines[a + 2 + 3].split(":")[1][1:].strip())
    except Exception:
        pass
    a += 11

arq.close()


arq = open("usuariofluig.csv", "rt")
lines = arq.readlines()[1:]
i = 0
write = open("Usuario Bloqueado no AD com usuario no Fluig ativo.txt", "wt")
write_now = ""
for line in lines:
    if dic.get(line.split(";")[0]) == 0:
        if line.split(";")[3].upper() == 'ACTIVE':
            #print(line.split(";")[1])
            write_now += line.split(";")[1] + "\n"
            i += 1
print(i)
arq.close()
write.write(write_now)
write.close()
