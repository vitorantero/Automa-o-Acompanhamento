from time import sleep as t
import os
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as df
import Gerar_Ordem as OS
import datetime

warnings.filterwarnings("ignore", category=DeprecationWarning)
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
drv = webdriver.Chrome(executable_path='chromedriver.exe', options=chrome_options)

user = os.getlogin()
pasta_download = os.path.expanduser('~/Downloads')
caminho_download = os.listdir(pasta_download)
arquivo_procurado = 'arquivo.csv'
arquivo_procurado_filiais = 'arquivo.csv'

os.system('CLS')
tb_frota = os.path.join(pasta_download, arquivo_procurado)
tb_filais = os.path.join(pasta_download, arquivo_procurado_filiais)
tb_rastreamento = 'arquivo.xlsx'

data_atual = datetime.date.today()
data_formatada = data_atual.strftime('%d/%m/%Y')
print(data_formatada)

for arquivo in caminho_download:
    if arquivo_procurado in arquivo or arquivo_procurado_filiais in arquivo:
        caminho_arquivo = os.path.join(pasta_download, arquivo)
        os.remove(caminho_arquivo)
        print(f'Arquivo excluído: {arquivo}')

print('Baixando Arquivo')
drv.get(f'https://site.com.br/{data_formatada}')

while True:
    if os.path.exists(os.path.join(pasta_download, arquivo_procurado)):
        print("Arquivo Relação de frotas Baixado!")
        break

drv.get('https://site.com.br')    
while True:
    if os.path.exists(os.path.join(pasta_download, arquivo_procurado_filiais)):
        print("Arquivo Filiais Baixado!")
        break

# Base Frota
tb_frota = df.read_csv(tb_frota, delimiter=';', encoding='ISO-8859-1')
tb_frota = tb_frota[['Frota', 'Placa', 'Data de Desinstalação']]
tb_frota_filtrada = tb_frota[(tb_frota['Data de Desinstalação'].isna())]
#print(tb_frota_filtrada)

tb_filais = df.read_csv(tb_filais, delimiter=';', encoding='ISO-8859-1')
tb_filais_resumo = tb_filais[['Loja','CEP', 'Endereço', 'Número', 'Bairro', 'Fone']]

tb_rastreamento = df.read_excel(tb_rastreamento)
tb_rastreamento.rename(columns={'PLACA': 'Placa'}, inplace=True)

tb_base_tratada = df.merge(tb_frota_filtrada, tb_rastreamento, on='Placa')

tb_base_tratada.to_excel('Base.xlsx', index=False)
print('Base Criada!')


tb_filiais_filtrada = tb_filais[(tb_filais['Loja']=='CURITIBA CENTRO APTA')] 

loja = tb_filiais_filtrada['Loja'].iloc[0]
cep = tb_filiais_filtrada['CEP'].iloc[0]
endereco = tb_filiais_filtrada['Endereço'].iloc[0]
numero = tb_filiais_filtrada['Número'].iloc[0]
bairro = tb_filiais_filtrada['Bairro'].iloc[0]
cidade = tb_filiais_filtrada['Cidade'].iloc[0]
uf = tb_filiais_filtrada['UF'].iloc[0]
telefone = tb_filiais_filtrada['Fone'].iloc[0]

print(f'Loja: {loja}, Cep: {cep}, Endereço:{endereco}, Bairro: {bairro}, Cidade: {cidade}, UF: {uf}, Telefone: {telefone}')

OS.Criar_OS()
