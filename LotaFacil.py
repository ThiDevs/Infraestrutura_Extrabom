def main():
    arq = open("LoteriaLimpa.csv","rt")
    lines = arq.readlines()
    dic = {}
    q = 0
    for i in lines:
        bola = i.strip()
        if q != 0:
            for bolas in bola.split(";"):
                if bolas != "":
                    try:
                        dic.update({bolas: dic[bolas]+1})
                    except Exception:
                        dic.update({bolas:1})
        q += 1

    arq = open("NovaLoteria.csv","wt")
    write = ""
    for key in dic.keys():
        print(key)
        write += key + ";" + str(dic[key]) + "\n"
    arq.write(write)

def limpa():
    arq = open("Pasta1.csv", "rt")
    lines = arq.readlines()
    q = 0
    arq = open("LoteriaLimpa.csv", "wt")
    write = ""
    for i in lines:
        bola = i.strip()
        if bola != ";;;;;;;;;;;;;;":
            print(bola)
            write += bola+"\n"
            q+=1
    arq.write(write)
    arq.close()

def UltimosXJogos():
    ultimos = 15
    arq = open("LoteriaLimpa.csv", "rt")
    lines = arq.readlines()
    q = 0
    dic = {}
    write2 = ""
    for i in lines:
        bola = i.strip()
        if q >= len(lines) - 15:
            print(bola)
            for bolas in bola.split(";"):
                if bolas != "":
                    try:
                        dic.update({bolas: dic[bolas]+1})

                    except Exception:
                        dic.update({bolas:1})
            write2 += str(sorted(bola.split(";"), key=int)) + "\n"

        q+=1
    print(dic)
    arq = open("NovaLoteria2.csv","wt")
    write = ""
    for key in dic.keys():
        print(key)
        write += key + ";" + str(dic[key]) + "\n"
    arq.write(write)
    arq.close()
    print(write2)

UltimosXJogos()