from openpyxl import load_workbook

# 1 lê pasta de trabalho e planilha

wb = load_workbook("data/pivot_table.xlsx")
sheet = wb ["Relatorio"]
print(sheet)

# 2 acessando um valor expecifico

print(sheet["A3"].value)
print(sheet["B3"].value)


# 3 interando valores por meio de loop

for i in range(2, 6):
    ano = sheet["A%s" %i].value
    an = sheet["B%s" %i].value
    bt = sheet["C%s" %i].value
    print("{0} o Aston martin vendeu {1} eo Bentley vendeu {2}".format(ano,an,bt))
    
