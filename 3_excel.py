from openpyxl import Workbook

print('lendo dados do arquivo')
#1 importando arquivo txt
file_txt= open('files/gastos.txt', 'r',encoding='utf-8')
file= file_txt.read()
list_data= file.splitlines()

#iterando valores
for i in range(0,len(list_data)):
    list_data[i] = list_data[i].split(',')

#criando planilha

wb= Workbook()
# nome= 'files/gastos.xlsx'


ws1=wb.active
ws1.title="PLGastos"

for row in list_data:
     ws1.append(row)
    
    
wb.save(filename='files/gastos.xlsx')


