import pandas as pd

def maior_gastos(df):
    maior_gasto = 0
    categoria = ''
    for j in df['categoria'].drop_duplicates():
        df_temp = df.loc[df['categoria'] == j]
        soma = df_temp['Gasto'].sum()
        if soma > maior_gasto:
            maior_gasto = soma
            categoria = j
        
    return maior_gasto, categoria 


def gastos_por_categoria(df):
    categoria = []
    valor = []
    for j in df['categoria'].drop_duplicates():
        df_temp = df.loc[df['categoria'] == j]
        soma = df_temp['Gasto'].sum()
        categoria.append(j)
        valor.append(soma)
    
    return valor, categoria