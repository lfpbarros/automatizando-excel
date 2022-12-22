from openpyxl import Workbook

wb = Workbook()
# instanciando um workbook vazio

sheet = wb.active
# seleciona a primeira sheet, a ativa

sheet['A1'].value = 100
sheet['A2'].value = 200

formula = '=SUM(A1:A2)'

sheet['A3'].value = formula
# a função passada tem que ser em inglês, e os separadores de argumento tem que ser vírgula

from openpyxl.formula.translate import Translator

sheet['B1'].value = 300
sheet['B2'].value = 250

sheet['B3'] = Translator(formula, origin='A3').translate_formula('B3')
# traduz a fórmula que está numa outra célula e adapta para a que selecionarmos

wb.save('formula.xlsx')

from openpyxl.utils import FORMULAE
FORMULAE
# retorna a lista de fórmulas