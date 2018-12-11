import openpyxl
import datetime
wb = openpyxl.load_workbook("Resultado.xlsx", read_only=False)
ws = wb.active


for i in range(2,len(ws['B'])):
    classificacao = str(ws['AH'+str(i)].value).strip()


    exece = str(ws['Y'+str(i)].value).strip()
    acomexec = str(ws['DU' + str(i)].value).strip()
    finalizada = str(ws['C' + str(i)].value).strip()
    solicitacao = str(ws['A' + str(i)].value).strip()




    if finalizada == "Finalizada":
        if classificacao != "" and classificacao != "None" and classificacao != " ":

            dt1 = classificacao
            DATASAIDAEXEC = str(ws['AL' + str(i)].value).strip()  # DATASAIDAEXECUCAOTEC - solicitacao_infraestrutura

            if DATASAIDAEXEC != " " and DATASAIDAEXEC != "" and DATASAIDAEXEC != "None":
                dt2 = DATASAIDAEXEC
                pass
            else:
                DATASAIDAACOMPANHAMENTO = str(ws['AO' + str(i)].value).strip() #DATASAIDAACOMPANHAMENTO - solicitacao_infraestrutura
                dt2 = DATASAIDAACOMPANHAMENTO
        else:
            dt1 = str(ws['DI' + str(i)].value).strip() #Atividade - solicitacao_infraestrutura - Aprovar - Conclusão

            DATASAIDAEXEC = str(ws['AL' + str(i)].value).strip()  # DATASAIDAEXECUCAOTEC - solicitacao_infraestrutura

            if DATASAIDAEXEC != " " and DATASAIDAEXEC != "" and DATASAIDAEXEC != "None":
                dt2 = DATASAIDAEXEC
                pass
            else:
                DATASAIDAACOMPANHAMENTO = str(
                    ws['AO' + str(i)].value).strip()  # DATASAIDAACOMPANHAMENTO - solicitacao_infraestrutura
                dt2 = DATASAIDAACOMPANHAMENTO

        print(dt1)
        print(dt2)


from datetime import datetime
s = '2015/08/05 08:12:23'
t = '2015/08/09 08:13:23'
f = '%Y/%m/%d %H:%M:%S'
dif = (datetime.strptime(t, f) - datetime.strptime(s, f)).total_seconds()
print(round(dif/60/60))