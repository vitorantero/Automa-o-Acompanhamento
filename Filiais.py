from time import sleep as t
import os
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

placa = 'AAA1A11'
pasta_download = os.path.expanduser('~/Downloads')
arquivo_procurado_filiais = 'arquivo.csv'
arquivo_procurado = 'arquivo.csv'

tb_filais = os.path.join(pasta_download, arquivo_procurado_filiais)
tb_filais = pd.read_csv(tb_filais, delimiter=';', encoding='ISO-8859-1')
tb_frota = os.path.join(pasta_download, arquivo_procurado)
tb_frota = pd.read_csv(tb_frota, delimiter=';', encoding='ISO-8859-1')

for index, row in tb_frota.iterrows():
    placa_frota = row['Placa']
    filial = row['Filial Atual']

    if placa in placa_frota:
        print(filial)
        break

for index, row in tb_filais.iterrows():
    filial_f = row['Loja']

    if filial in filial_f:
        print(filial)
        break
tb_filiais_filtrada = tb_filais[(tb_filais['Loja'] == 'Loja III São Paulo - Brasil')] 
print(tb_filiais_filtrada)
loja = row['Loja'].iloc[0]
cep = row['CEP'].iloc[0]
endereco = row['Endereço'].iloc[0]
numero = row['Número'].iloc[0]
bairro = row['Bairro'].iloc[0]
cidade = row['Cidade'].iloc[0]
uf = row['UF'].iloc[0]
telefone = row['Fone'].iloc[0]
print(f'Loja: {loja}, Cep: {cep}, Endereço:{endereco}, Numero: {numero}, Bairro: {bairro}, Cidade: {cidade}, UF: {uf}, Telefone: {telefone}')
