import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from sklearn.linear_model import LinearRegression
import openpyxl

#criando book
book = openpyxl.Workbook()

guias = ['MENOR FRETE', 'Hapag-Lloyd', 'Hapag-LLoyd SPOT', 'MSC', 'CMA-CGM', 'MAERSK LINES']

for guia in guias:
    print (guia)
    #adicionar uma guia
    book.create_sheet(guia)
    #adicionar cabeçalho ao book
    book[guia].append(['País Cliente', 'Incot.', 'PORTO DE DESTINO', 'Destino', '25.09', '02.10', '09.10', '16.10', '23.10'])

    # Carregue o arquivo Excel em um DataFrame
    caminho_arquivo = r'C:\Users\LucasCerqueiraGalvao\Documents\Python programs\Previsão\Comparativos.xlsx'

    df = pd.read_excel(caminho_arquivo, sheet_name=guia)
    


    # Especifique as colunas de recursos (X) pelos números das colunas (E a H)
    colunas = [4, 5, 6, 7]  # Suponha que estas são as colunas E a H (índices 4 a 7)
    dados = df.iloc[:, colunas]

    # Valores de entrada (X)
    x = np.array([1, 2, 3, 4]).reshape(-1, 1)

    for i in range (76):
        # Valores de saída (y) - Use o índice 35 para selecionar uma linha
        y = dados.iloc[i].values
        
        todos_zeros = all(elemento == 0 for elemento in y)

        if not todos_zeros:
            media_sem_zeros = np.mean([valor for valor in y if valor != 0])
            # Substituir os zeros pela média dos valores não nulos
            y = [media_sem_zeros if valor == 0 else valor for valor in y]
    
        # Criar um modelo de regressão linear
        modelo = LinearRegression()
        modelo.fit(x, y)

        # Obter os coeficientes do modelo
        a_coeff = modelo.coef_
        l_coeff = modelo.intercept_
        prev = l_coeff + 5*a_coeff
        #print (l_coeff, a_coeff, prev)
        df.at[i, '23.10'] = prev
        #manipulação de book, adicionar a linha do df no book 
        book[guia].append(df.iloc[i].values.tolist())

    print ("Foi a:", guia)

novo_caminho_arquivo = r'C:\Users\LucasCerqueiraGalvao\Documents\Python programs\Previsão\Comparativos_atualizado.xlsx'
book.save(novo_caminho_arquivo)

print ('Deu tudo certo!')    
