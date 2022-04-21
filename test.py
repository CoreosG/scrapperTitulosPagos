import xlrd

arquivo = xlrd.open_workbook('tabela.xls')

tabela = arquivo.sheet_by_index(0)


listTitulos = {}
for x in range(11, tabela.nrows-1): 
    listTitulos[tabela.cell_value(rowx=x, colx=1)] = tabela.cell_value(rowx=x, colx=2) 

print(listTitulos)
