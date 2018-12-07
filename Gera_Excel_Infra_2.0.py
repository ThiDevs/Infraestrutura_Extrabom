import datetime
from openpyxl import load_workbook
from datetime import datetime

dic = {}


def gerarColunas():
    import string
    a = string.ascii_uppercase
    lst = []
    for i in a:
        lst.append(i)
    lst2 = []
    for j in range(5):
        for i in a:
            lst2.append(lst[j] + i)
    lst += lst2
    return lst


def calc(hour, hours2):
    date_format = "%m-%d-%Y %H:%M:%S"
    time1 = datetime.strptime('8-01-2008 ' + hour, date_format)
    time2 = datetime.strptime('8-01-2008 ' + hours2, date_format)
    diff = time2 - time1
    minutes = diff.seconds / 60
    print(str(minutes) + ' Minutes')


NAO = 0
SIM = 0
TOTAL = 0
CANCELADAS = 0
ABERTOS_HJ = 0
APROVAR = 0
hour = str(datetime.now().hour) + ":" + str(datetime.now().minute) + ":" + str(datetime.now().second)

wb = load_workbook(filename='Resultado.xlsx', read_only=True)
ws = wb['Resultado da consulta de solici']

colunas = gerarColunas()

for i in colunas:
    if ws[i + "1"].value == "Solicitação":
        coluna_solicitacao = i

    elif ws[i + "1"].value == "Situação":
        coluna_situacao = i

    elif ws[i + "1"].value == "Localização":
        coluna_localizacao = i

    elif ws[i + "1"].value == "Início":
        coluna_inicio = i

    elif ws[i + "1"].value == "Fim":
        coluna_fim = i

    elif ws[i + "1"].value == "SLACONSUMIDOEXECUCAOTEC - solicitacao_infraestrutura":
        coluna_slaexe = i

    elif ws[i + "1"].value == "SLACONSUMIDOACOMPANHAMENTO - solicitacao_infraestrutura":
        coluna_slaacom = i

    elif ws[i + "1"].value == "DESCRICAOITEM - solicitacao_infraestrutura":
        coluna_item = i

    elif ws[i + "1"].value == "TOTALHOURSSLA - solicitacao_infraestrutura":
        coluna_sla = i

    elif ws[i + "1"].value == "CODIGOTECNICO - solicitacao_infraestrutura":
        coluna_tec = i


def main():
    arq = open("Tecnicos.txt", 'rt')
    lines = arq.readlines()
    for line in lines:
        dic.update({line.split(";")[0]: line.split(";")[1].strip()})
    arq.close()

    from openpyxl import Workbook
    wb_new = Workbook(write_only=True)

    ws_new = wb_new.create_sheet(title="Dados")

    ws_new.append(
        ['Solicitações', 'Situação', 'Localização', 'Técnico', 'Inicio', 'Fim', 'Item', 'SLA', 'SLA Consumido',
         "Dentro do prazo?"])

    wb_new.save("Relatorio/Relatorio_InternoVsExterno.xlsx")
    wb_new.close()

    import threading
    percorre = 10
    for i in range(1, 5, 1):
        T1 = threading.Thread(target=GeraExcel, args=(percorre * i, i, percorre))
        T1.start()


Cont = 0



def GeraExcel(percorre, thread, percorre2):
    wb = load_workbook(filename='Resultado.xlsx', read_only=True, data_only=True)
    ws = wb['Resultado da consulta de solici']
    Lista_Salvar = []
    if thread == 1:
        j = percorre - percorre2 + 2
    else:
        j = percorre - percorre2
    for i in range(j, percorre, 1):

        try:
            global CANCELADAS, APROVAR, ABERTOS_HJ
            solicitacoes = str(int(ws[coluna_solicitacao + str(i)].value)).strip()

            localizacao = str(ws[coluna_localizacao + str(i)].value).strip()
            if localizacao == "Classificação" or localizacao == "Acompanhamento da Execução" or localizacao == "Acompanharmento da Execução" or localizacao == "Execução da Solicitação":
                situacao = "Em aberto"
            elif localizacao == "Aprovar":
                situacao = "Aprovação"
                APROVAR += 1
            elif localizacao == "Cancelada":
                situacao = "Cancelada"
                CANCELADAS += 1
            else:
                situacao = "Fechado"

            try:
                tecnico = dic[ws[coluna_tec + str(i)].value]
            except Exception:
                tecnico = ws[coluna_tec + str(i)].value

            item = ws[coluna_item + str(i)].value
            SLA = ws[coluna_sla + str(i)].value
            inicio = ws[coluna_inicio + str(i)].value
            if inicio.split("/")[1] == "12":
                ABERTOS_HJ += 1

            fim = ws[coluna_fim + str(i)].value
            SLA_Consumido_TXT = ""
            SLA_Consumido_TXT2 = ""
            SLA_Consumido_aux = ""
            SLA_Consumido_TXT_aux = ""

            if situacao == "Fechado":
                try:
                    SLA_Consumido_TXT = ws[coluna_slaexe + str(i)].value
                    SLA_Consumido = SLA_Consumido_TXT.split(":")[0]  # SLA EXECUÇÃO#
                    SLA_Consumido_aux = SLA_Consumido
                    SLA_Consumido_TXT_aux = SLA_Consumido_TXT
                    if SLA_Consumido_TXT.split(":")[1] != "":
                        SLA_Consumido += ":" + SLA_Consumido_TXT.split(":")[1]

                except Exception:
                    SLA_Consumido = "00:00"

                try:
                    SLA_Consumido_TXT2 = ws[coluna_slaacom + str(i)].value  # SLA ACOMPANHAMENTO#
                    SLA_Consumido2 = SLA_Consumido_TXT2.split(":")[0]
                    if SLA_Consumido_TXT2.split(":")[1] != "":
                        SLA_Consumido2 += ":" + SLA_Consumido_TXT2.split(":")[1]

                except Exception:
                    SLA_Consumido2 = "00:00"

                if int(SLA_Consumido.split(":")[0]) <= int(SLA_Consumido2.split(":")[0]):
                    SLA_Consumido_aux = SLA_Consumido2
                    SLA_Consumido_TXT_aux = SLA_Consumido_TXT2
                    try:
                        if int(SLA_Consumido.split(":")[1]) <= int(SLA_Consumido2.split(":")[1]):
                            SLA_Consumido_aux = SLA_Consumido2
                            SLA_Consumido_TXT_aux = SLA_Consumido_TXT2
                        else:
                            SLA_Consumido_aux = SLA_Consumido
                            SLA_Consumido_TXT_aux = SLA_Consumido_TXT
                    except Exception:
                        pass
            else:
                SLA_Consumido_aux = "00:00"
            SLA_Consumido = SLA_Consumido_aux
            SLA_Consumido_TXT = SLA_Consumido_TXT_aux
            global NAO
            global SIM
            global TOTAL
            if SLA_Consumido != "" and int(SLA_Consumido.split(":")[0]) <= int(
                    SLA.split(":")[0]) and SLA_Consumido != "00:00":

                if int(SLA_Consumido.split(":")[0]) > int(SLA.split(":")[0]):
                    Lista_Salvar.append(
                        [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, SLA_Consumido_TXT,
                         "Não"])
                    NAO += 1

                elif int(SLA_Consumido.split(":")[0]) == int(SLA.split(":")[0]):
                    if int(SLA_Consumido_TXT.split(":")[1]) <= 0:
                        Lista_Salvar.append(
                            [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, SLA_Consumido_TXT,
                             "Sim"])
                        SIM += 1
                    else:
                        Lista_Salvar.append(
                            [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, SLA_Consumido_TXT,
                             "Não"])
                        NAO += 1
                else:

                    Lista_Salvar.append(
                        [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, SLA_Consumido_TXT,
                         "Sim"])
                    SIM += 1

            elif SLA_Consumido != "" and int(SLA_Consumido.split(":")[0]) > int(SLA.split(":")[0]):

                Lista_Salvar.append(
                    [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, SLA_Consumido_TXT, "Não"])
                NAO += 1
            else:
                Lista_Salvar.append(
                    [solicitacoes, situacao, localizacao, tecnico, inicio, fim, item, SLA, "-", "-"])
            TOTAL += 1

        except Exception:
            pass

    print(str(thread) + " thread terminado")

    Finaliza(thread, Lista_Salvar)


lst = []

Lista_Salvar2 = []


def Finaliza(thread, Lista_Salvar):
    global Lista_Salvar2
    lst.append(thread)

    Lista_Salvar2 += Lista_Salvar

    if len(lst) == 4:
        wb_new = load_workbook(filename='Relatorio/Relatorio_InternoVsExterno.xlsx', read_only=False)
        ws_new = wb_new["Dados"]
        global NAO
        global SIM
        global TOTAL

        for list in Lista_Salvar2:
            ws_new.append(list)

        from openpyxl.styles import Alignment
        from openpyxl.styles import Color, PatternFill, Font, Border
        from openpyxl.styles import colors
        from openpyxl.cell import Cell

        redFill = PatternFill(start_color='e2fe13',
                              end_color='e2c813',
                              fill_type='solid')

        ws_new = wb_new.create_sheet(title="Informações")

        ws_new.merge_cells('A1:H1')
        cell = ws_new.cell(row=1, column=1)
        from openpyxl.styles.borders import Border, Side
        from openpyxl import Workbook
        cell.value = 'Geral'
        cell.fill = redFill

        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
        for i in range(1,9,1):
            ws_new.cell(row=1, column=i).border = thin_border
            ws_new.cell(row=2, column=i).border = thin_border
        lst2 = ["Total ", "Total Ativos", "Abertos em Dezembro", "Total de solicitações finalizadas",
                       "Percentual de Atendimento", "Finalizadas dentro do prazo", "Finalizadas fora do prazo",
                       "Canceladas"]
        z = 1
        ft = Font(color=colors.WHITE)

        blackFill = PatternFill(start_color='010204',
                              end_color='010204',
                              fill_type='solid')

        for name in lst2:
            ws_new.cell(row=2, column=z).value = name
            ws_new.cell(row=2, column=z).font = ft
            ws_new.cell(row=2, column=z).fill = blackFill
            z += 1


        cell.alignment = Alignment(horizontal='center', vertical='center')

        ws_new.append(
            [TOTAL, TOTAL - CANCELADAS, ABERTOS_HJ, SIM + NAO, str(round(((SIM + NAO) * 100) / TOTAL - CANCELADAS)) + " %", SIM,
             NAO, CANCELADAS])

        dims = {}
        for row in ws_new.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
        for col, value in dims.items():
            ws_new.column_dimensions[col].width = value

        wb_new.save('Relatorio/Relatorio_InternoVsExterno.xlsx')
        wb_new.close()
        hour2 = str(datetime.now().hour) + ":" + str(datetime.now().minute) + ":" + str(datetime.now().second)
        calc(hour, hour2)


main()
