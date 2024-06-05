import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException
import os
import pandas as pd

# configurando o Chrome Options
chrome_options = Options()

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico, options=chrome_options)

# Define uma parte fixa do nome do arquivo e uma lista de ações
parte_fixa = '.SA.csv'
açao = ['AZUL4', 'ITUB4', 'ABEV3', 'MGLU3', 'VALE3', 'PETR4', 'CYRE3', 'BBAS3']
açao.sort()
diretorio = 'C:\\Users\\Cliente\\Downloads'

# Define as listas que serão utilizadas para armazenar os dados
açoes = list()
date2 = list()
open2 = list()
high2 = list()
low2 = list()
close2 = list()
adj2 = list()
volume2 = list()
valorizaçao_12meses = list()
valorizaçao_mes = list()


# Função para aplanar uma lista de listas em uma única lista
def aplanar_lista_de_listas(lista_de_listas):
    lista = list()
    for a in lista_de_listas:
        lista.extend(a)
    return lista

# Loop para acessar e baixar dados históricos de ações
for i in açao:
    try:
        driver.get(f'https://br.financas.yahoo.com/quote/{i}.SA/history/')
        nome_arquivo = i + parte_fixa
        caminho_arquivo = os.path.join(diretorio, nome_arquivo)

        time.sleep(1)

        # captura o link de instalação do arquivo
        elm = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[1]/div/div[3]/div[1]/div/div[2]/div/div/section/div[1]/div[2]/span[2]/a')
        href = elm.get_attribute('href')

        # Verifica se o arquivo já existe, e se sim, o remove antes de baixar a nova versão
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            driver.get(href)
            print(f'{nome_arquivo} ja existe, porém está sendo substituído por uma versão mais recente')

        else:
            driver.get(href)
            print(f'{nome_arquivo} está sendo baixado')

        # Acessa a página de statusinvest para pegar a valorização mensal e anual
        driver.get(f'https://statusinvest.com.br/acoes/{i}')

        time.sleep(1)

        # Extrai a valorização de 12 meses e do mês
        valorizaçao_12meses1 = driver.find_element(By.XPATH, '//*[@id="main-2"]/div[2]/div/div[1]/div/div[5]/div/div[1]/strong').text
        valorizaçao_12meses.append(valorizaçao_12meses1)
        valorizaçao_mes1 = driver.find_element(By.XPATH, '//*[@id="main-2"]/div[2]/div/div[1]/div/div[5]/div/div[2]/div/span[2]/b').text
        valorizaçao_mes.append(valorizaçao_mes1)

    except NoSuchElementException:
        print('Elemento não encontrado')


    except Exception as e:
        print(f'Erro {e}')


# Loop para ler os arquivos CSV baixados e extrair os dados
for i in açao:
    try:
        time.sleep(0.5)
        caminhodf = f'C:\\Users\\Cliente\\Downloads\\{i}.SA.csv'
        df = pd.read_csv(caminhodf)
        date1 = df['Date'].tolist()
        open1 = df['Open'].tolist()
        high1 = df['High'].tolist()
        low1 = df['Low'].tolist()
        close1 = df['Close'].tolist()
        adj1 = df['Adj Close'].tolist()
        volume1 = df['Volume'].tolist()
        date2.append(date1)
        open2.append(open1)
        high2.append(high1)
        low2.append(low1)
        close2.append(close1)
        adj2.append(adj1)
        volume2.append(volume1)

    except Exception as e:
        print(f'Erro: {e}')

# Aplanar as listas de listas em listas simples
date = aplanar_lista_de_listas(date2)
open = aplanar_lista_de_listas(open2)
high = aplanar_lista_de_listas(high2)
low = aplanar_lista_de_listas(low2)
close = aplanar_lista_de_listas(close2)
adj = aplanar_lista_de_listas(adj2)
volume = aplanar_lista_de_listas(volume2)

# Criar uma lista de ações repetida para cada data disponível, que sera anexada como uma nova coluna no arquivo final
for i in açao:
    for e in range(len(date1)):
        açoes.append(i)


# Criar DataFrames a partir das listas
df_açoes = pd.DataFrame({'Açoes': açoes})
df_date = pd.DataFrame({'Date': date})
df_open = pd.DataFrame({'Open': open})
df_high = pd.DataFrame({'High': high})
df_low = pd.DataFrame({'Low': low})
df_close = pd.DataFrame({'Close': close})
df_adj = pd.DataFrame({'Adj Close': adj})
df_volume = pd.DataFrame({'Volume': volume})
df_valorizaçao12 = pd.DataFrame({'Valorização 12 meses': valorizaçao_12meses})
df_valorizaçaomes = pd.DataFrame({'Valorização mes': valorizaçao_mes})
df_açao = pd.DataFrame({'Açoes': açao})

# Concatenar os DataFrames em dois DataFrames principais
df_concatenado = pd.concat([df_açoes, df_date, df_open, df_high, df_low, df_close, df_adj, df_volume], axis=1)
df_concatenado1 = pd.concat([df_açao, df_valorizaçaomes, df_valorizaçao12], axis=1)

# Define o nome do arquivo
nome_planilha = 'planilhadf.xlsx'

# Salvar os DataFrames concatenados em um arquivo Excel com diferentes abas
with pd.ExcelWriter(nome_planilha, engine='openpyxl') as writer:
    df_concatenado.to_excel(writer, sheet_name='Ações', index=False)
    df_concatenado1.to_excel(writer, sheet_name='Valorização', index=False)
