import pandas as pd
import os

rel = os.getcwd()

df_caminho = os.path.join(os.path.realpath(os.getcwd()), "dados.xlsx")

df = pd.read_excel(df_caminho)

for index, row in df.iterrows():
    foro_pesquisa = row[0]
    print(foro_pesquisa)
    for column in row[1:]:
        cnpj = column
        print(cnpj)