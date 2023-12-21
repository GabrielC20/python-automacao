import pandas as pd 

# importando dados
data = pd.read_excel("data/VendaCarros.xlsx")
print(data)

#2 listar os primeiros registros

print(data.head())

#listar os utimos registros
print(data.tail())

# contagem de valores com o frabricante

print(data["Fabricante"].value_counts())