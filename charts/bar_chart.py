from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

wb = Workbook()

ws = wb.active

rows = [
    ('Number', 'Batch 1', 'Batch 2'),
    (2, 10, 30),
    (3, 40, 60),
    (4, 50, 70),
    (5, 20, 10),
    (6, 10, 40),
    (7, 50, 30),
]

for row in rows:
    ws.append(row)

c1 = BarChart()
c1.type = 'col' # permite o empilhamento vertical, type 'bar' tem o empilhamento horizontal

c1.title = 'Gráfico de Coluna'
c1.y_axis.title = 'Número Teste'
c1.x_axis.title = 'Tamanho da amostra (mm)'

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)
cats = Reference(ws, min_col=1, min_row=2, max_row=7)
# criando categorias

c1.add_data(data, titles_from_data=True)
c1.set_categories(cats)
# adicionando as categorias

ws.add_chart(c1, "A10")

from copy import deepcopy

c2 = deepcopy(c1) # com o deepcopy, criamos uma cópia ao invés do puro apontamento

c2.type = 'bar'
c2.title = 'Gráfico de barras'
ws.add_chart(c2, "G10")


wb.save('bar.xlsx')