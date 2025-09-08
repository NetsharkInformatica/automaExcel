##Criação de graficom com Python
from openpyxl import Workbook
from openpyxl.chart import AreaChart,Reference,series

wb= Workbook()
ws= wb.active

data= [
    ['Ano','Lucro','Custos'],
    [2001, 25, 12000],
    [2002, 26, 13000],
    [2003, 28, 14000],
    [2004, 30, 15000],
    [2005, 31, 16000],
    [2006, 32, 17000],
    [2007, 33, 18000],
    [2008, 34, 19000],
    [2009, 35, 20000],
    [2010, 35, 21000],
    [2011, 36, 22000],
    [2012, 37, 23000],
    [2013, 37, 24000],
    [2014, 38, 25000],
    [2015, 38, 26000],
    [2016, 38, 27000],
    [2017, 39, 28000],
    [2018, 39, 29000],
    [2019, 40, 30000],
    [2020, 40, 31000]
]
for d in data:
    ws.append(d)
    
chart = AreaChart()

chart.title=' Lucro x Custos por ano'
chart.style=16
chart.x_axis.title='Ano'
chart.y_axis.title=' Porcentagem %'

categorias = Reference(
    ws,
    min_col=1,  # Coluna A (Ano)
    min_row=2,  # Começa da linha 2 (pula o cabeçalho)
    max_row=21  # Até a linha 21 (20 registros + cabeçalho)
)

# CORREÇÃO AQUI: Referência para os dados (Lucro e Custos)
dados = Reference(
    ws,
    min_col=2,  # Coluna B (Lucro)
    min_row=1,  # Inclui o cabeçalho na linha 1
    max_col=3,  # Até coluna C (Custos)
    max_row=21  # Até linha 21
)

# CORREÇÃO AQUI: Use titles_from_data=True para pegar os títulos da planilha
chart.add_data(dados, titles_from_data=True)
chart.set_categories(categorias)

# Adiciona o gráfico na planilha
ws.add_chart(chart, 'A23')

# Salva o arquivo
wb.save('files/chart.xlsx')
print("Gráfico criado com sucesso!")

