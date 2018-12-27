arq = open('Lista Pdvs.csv','rt')
LojaPdv = arq.readline().split(";")

writearq = open("PDVS.csv","wt")
write = ''

while LojaPdv[0] != '':
    write += LojaPdv[0]
    for i in range(int(LojaPdv[1])):
        write +=  ";" + str(i+1)
    LojaPdv = arq.readline().split(";")
    write += "\n"
print(write)

writearq.write(write)