import xlsxwriter as opcoesDOXL

import os

nomeArquivo = 'C:\\Users\\SylviaMissio\\Desktop\\RPA1\\xlsx\\Grafico.xlsx'
workbook = opcoesDOXL.Workbook(nomeArquivo)

sheetDados = workbook.add_worksheet("Resumo")

bold = workbook.add_format({'bold':1})

Titulos = ['Vendedores', 'Total Vendas']
dadosTabela = [
    ["Ana", "Pedro", "Allan", "Francisco", "Rosa", "Amanda"],
    [700,500,450,400,380,370],
]

sheetDados.write_row('A1', Titulos, bold)
sheetDados.write_column('A2',dadosTabela[0])
sheetDados.write_column('B2',dadosTabela[1])

graficoColunas = workbook.add_chart({'type': 'column'})

graficoColunas.add_series({
    'name': '=Resumo!$B$1',
    'categories': '=Resumo!$A$2:$A$7',
    'values': '=Resumo!$B$2:$B$7',
})

graficoColunas.set_title({'name': 'Gráfico Total de vendas'})
graficoColunas.set_x_axis({'name': 'Vendores'})
graficoColunas.set_y_axis({'name': 'Vendas'})

graficoColunas.set_style(11)

sheetDados.insert_chart('D2', graficoColunas, {'x_offset': 25, 'x_offset':10})

#########################################################

graficoEmpilhado = workbook.add_chart({'type': 'area','subtype': 'stacked'})

graficoEmpilhado.add_series({
    'name': '=Resumo!$B$1',
    'categories': '=Resumo!$A$2:$A$7',
    'values': '=Resumo!$B$2:$B$7',
})

graficoEmpilhado.set_title({'name': 'Gráfico Empilhado'})
graficoEmpilhado.set_x_axis({'name': 'Vendedores'})
graficoEmpilhado.set_y_axis({'name': 'Vendas'})

graficoEmpilhado.set_style(12)

sheetDados.insert_chart('L2', graficoEmpilhado, {'x_offset': 25, 'x_offset':10})

workbook.close()

os.startfile(nomeArquivo)