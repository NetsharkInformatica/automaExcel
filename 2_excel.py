from openpyxl import load_workbook

#1 lendo workbook e buscando planilha
wb= load_workbook(filename='files/test.xlsx')
planilha= wb['Planilha1']

#2 acessando um determinado valor
#iterando valores a partir de um loop

for i in range(2,7):
    #print(f' {planilha['A[i]}'].value}')
    ano=planilha['A%s' %i].value
    lucro=planilha['B%s' %i].value
    custos=planilha['C%s' %i].value
    
    print(f'o ano {ano} teve {lucro} de lucro e {custos} de custos')

#print(planilha['B2'].value)

