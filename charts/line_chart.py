from openpyxl import Workbook
from openpyxl.chart import (LineChart, Reference)
from datetime import date

wb = Workbook()

ws = wb.active

rows = [
    ['Date', 'Batch 1', 'Batch 2', 'Batch 3'],
    [date(2015,9, 1), 40, 30, 25],
    [date(2015,9, 2), 40, 25, 30],
    [date(2015,9, 3), 50, 30, 45],
    [date(2015,9, 4), 30, 25, 40],
    [date(2015,9, 5), 25, 35, 30],
    [date(2015,9, 6), 20, 40, 35],
]

for row in rows:
    ws.append(row)
# Adicionando os dados

c1 = LineChart()
# instanciando o gráfico

c1.title = 'Gráfico de linha' # título do gráfico
c1.x_axis.title = 'Número Teste' # Títulos de eixo
c1.y_axis.title = 'Tamanho'

data = Reference(ws, min_col=2, max_col=4, min_row=1, max_row=7)
# Instanciando a referência de dados, onde começa e onde termina

c1.add_data(data, titles_from_data=True) # Adicionando os dados, o parâmetro titles_from data define se há cabeçalho na tabela

# para o openpyxl, cada linha do gráfico é uma series, então, podemos editar desse modo:

c1.series[0].marker.symbol = 'triangle' # alterando o marcador
c1.series[0].marker.graphicalProperties.solidFill = 'FF0000' # alterando a cor do marcador
c1.series[0].graphicalProperties.line.noFill = True # removendo a linha


ws.add_chart(c1, "A10")
# Adicionado o gráfico ao sheet

wb.save('line.xlsx')
