import pandas as pd

# 1-Importando Dados
data = pd.read_excel("C:/Users/User/Desktop/python_estudos/minicurso_Onebitecode/data/VendaCarros.xlsx")

# 2-Selecionando colunas específicas do dataframe
df = data[["Fabricante", "ValorVenda", "Ano"]]
print (df)

# 3-Criando tabela pivô
pivot_table = df.pivot_table(index="Ano", columns="Fabricante", values="ValorVenda", aggfunc="sum")
print (pivot_table)

# 4-Exportando tabela pivô em arquivo excel
pivot_table.to_excel("C:/Users/User/Desktop/python_estudos/minicurso_Onebitecode/data/pivot_table.xlsx","Relatorio")  # Exporta para arquivo excel
