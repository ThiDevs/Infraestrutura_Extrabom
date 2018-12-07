txt = ""

for i in range(30):
    for j in range(25):
        try:
            arq = open("C:\\Users\\thiago.alves.EXTRABOM\\Desktop\\vendas novembro\\"+str(i)+"\\"+str(j)+".csv",'rt')
            line = arq.readline()
            txt += line
            while line != "":
                line = arq.readline()
                txt += line

            arq.close()

        except Exception:
            pass

arq = open("Lojas.csv","wt")
arq.write(txt)