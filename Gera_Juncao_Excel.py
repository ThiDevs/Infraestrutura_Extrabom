import datetime
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors


def Juncao():
    wb = load_workbook(filename='Relatorio/Relatorio_InternoVsExterno.xlsx', read_only=False)
    ws = wb['Dados']
    dic_newExcel = {}
    percorre = 200
    for i in range(2, percorre, 1):

        solicitacoes = str(ws['A' + str(i)].value)
        tecnico = ws['D' + str(i)].value
        DentroPrazo = str(ws['J' + str(i)].value).strip()
        situacao = str(ws['B' + str(i)].value).strip()
        if situacao == "Fechado":
            try:
                lst = dic_newExcel.get(tecnico)[0]
                lst.append(solicitacoes)
                prazo = dic_newExcel.get(tecnico)[1]

            except Exception:
                lst = [solicitacoes]
                prazo = 0

            if DentroPrazo == "Sim":
                prazo += 1
            dic_newExcel.update({tecnico: (lst, prazo)})

    print(dic_newExcel)

    ws_Juntos = wb.create_sheet(title="Junção")

    ft = Font(color=colors.WHITE)

    blackFill = PatternFill(start_color='010204',
                            end_color='010204',
                            fill_type='solid')

    lista_names = ['Técnico', 'Total de chamados', 'Dentro do prazo', 'Porcetagem', 'Fora do prazo']
    ws_Juntos.auto_filter.ref = 'A1:E9'
    z = 1
    for name in lista_names:
        ws_Juntos.cell(row=1, column=z).value = name
        ws_Juntos.cell(row=1, column=z).font = ft
        ws_Juntos.cell(row=1, column=z).fill = blackFill
        z += 1

    ws = wb['Informações']

    ws.cell(row=6, column=1).value = "Técnico"
    ws.cell(row=6, column=2).value = "Solicitações"

    names = ['Paulo Tavares', 'Bruno Soares', 'Romildo Carvalho', 'Aires Mendonça', 'Marcos Aurélio', 'Valcleide Silva',
             'Daniel Ferreira',
             'Julio Pulcher', 'Andre Melo', 'Sandro Geraldino', 'Claudio']
    dic_Civil = {}
    for key in dic_newExcel.keys():
        for name in names:
            if name in key:
                dic_Civil.update({key: len(dic_newExcel.get(key)[0])})

        len_soli = len(dic_newExcel.get(key)[0])
        ws_Juntos.append(
            [key, len_soli, dic_newExcel.get(key)[1], str(round((dic_newExcel.get(key)[1] * 100) / len_soli)) + " %",
             len_soli - dic_newExcel.get(key)[1]])
    row = 7
    ordena = list(reversed(sorted(dic_Civil.values())))
    for i in range(len(dic_Civil.keys())):
        for names in dic_Civil.keys():
            if dic_Civil[names] == ordena[i]:
                print(names)
                print(ordena[i])
                ws.cell(row=row, column=1).value = names
                ws.cell(row=row, column=2).value = ordena[i]
                row =+ 1
                break
        del dic_Civil[names]

    wb.save('Relatorio/Relatorio_InternoVsExterno.xlsx')
    wb.close()


Juncao()
