txt = "Loja;PDV;Data da venda;Quantidade \n"
lojas = ""
for i in range(33):
    for j in range(25):
        try:
            arq = open("C:\\Users\\thiago.alves.EXTRABOM\\Documents\\arq\\"+str(i)+"\\"+str(j),'rt')
            print(i,j)
            lojas += str(i)+";"+ str(j) + "/"
            line = arq.readline()
            txt += line
            while line != "":
                line = arq.readline()
                txt += line

            arq.close()

        except Exception:
            pass
print(lojas)
arq = open("lojas.txt","wt")
arq.write(lojas)
arq.close()

arq = open("Lojas.csv","wt")
arq.write(txt)