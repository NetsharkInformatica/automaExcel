from openpyxl import Workbook

#1 criando workbook

wb= Workbook()
name= 'files/test.xlsx'

#2= utilizando o worksheet
ws1=wb.active
ws1.title='Planilha1'

#3 adicionando os dados
data=[
    ['Ano','Lucro','Custos'],
    [2021,'25%','40%'],
    [2022,'30%','20%'],
    [2023, '45%','25%'],
    [2024,'43%','17%'],
    [2025,'67%','22%'],
    [2026,'15%','30%']
    ]
for line in data:
    ws1.append(line)

#para criar outra folha de trabalho  
#ws2=wb.create_sheet(title='teste')
#ws2['D2'] = 'qualquer coisa que desejo inserir na celula D2'

wb.save(filename = name)