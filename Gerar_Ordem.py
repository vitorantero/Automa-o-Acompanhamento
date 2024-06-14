from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
from time import sleep as t
from datetime import datetime
import re
import os
import numpy as np
import Email as E

chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = "chromedriver.exe"
service = Service(chrome_driver)
drv = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(drv, 10)


def Criar_OS():
    tab = pd.read_excel('Base.xlsx') 
    tab['O.S Gerada'] = None
    num_linha = 1
    Os_Gerada = 1
    placa_anterior = None
    

    for index, row in tab.iterrows():
        linha = index
        drv.get("https://site.com.br")

        ######## Variaveis Declaradas
        data_hoje = datetime.now().strftime('%d/%m/%Y')
        data_18h = str(f'{data_hoje} 18:00 ')
        placa = row['Placa']
        cnpj1 = str(row['CNPJ1']).zfill(14)
        cnpj1 = cnpj1.split('.')[0]
        cnpj2 = str(row['CNPJ2'])
        cnpj2 = cnpj2.split('.')[0]
        cnpj3 = str(row['CNPJ2'])
        cnpj3 = cnpj3.split('.')[0]
        prestador1 = row['EMPRESA 1\n\n']
        prestador2 = row['EMPRESA 2\n\n']
        prestador3 = row['EMPRESA 3\n\n']  
        email1 = row['EMPRESA 1 E-MAIL'] 
        email2 = row['EMPRESA 2 E-MAIL'] 
        email3 = row['EMPRESA 3 E-MAIL'] 
        Valor1 = row['VALOR 1\n\n']
        Valor2 = row['VALOR 2\n\n']
        Valor3 = row['VALOR 3\n\n']

        ######## Arquivos
        pasta_download = os.path.expanduser('~/Downloads')
        arquivo_procurado_filiais = 'arquivo.csv'
        arquivo_procurado = 'arquivo.csv'

        tb_filais = os.path.join(pasta_download, arquivo_procurado_filiais)
        tb_filais =pd.read_csv(tb_filais, delimiter=';', encoding='ISO-8859-1')
        tb_frota = os.path.join(pasta_download, arquivo_procurado)
        tb_frota = pd.read_csv(tb_frota, delimiter=';', encoding='ISO-8859-1')

        for index, row in tb_frota.iterrows():
            placa_frota = row['Placa']
            filial = row['Filial Atual']

            if placa in placa_frota:
                break

        for index, row in tb_filais.iterrows():
            filial_f = row['Loja']

            if filial in filial_f:
                break
        tb_filiais_filtrada = tb_filais[(tb_filais['Loja'] == filial_f)] 
        loja = tb_filiais_filtrada['Loja'].iloc[0]
        cep = tb_filiais_filtrada['CEP'].iloc[0]
        endereco = tb_filiais_filtrada['Endereço'].iloc[0]
        numero = tb_filiais_filtrada['Número'].iloc[0]
        bairro = tb_filiais_filtrada['Bairro'].iloc[0]
        cidade = tb_filiais_filtrada['Cidade'].iloc[0]
        uf = tb_filiais_filtrada['UF'].iloc[0]
        telefone = tb_filiais_filtrada['Fone'].iloc[0]
        print(f'Loja: {loja}, Cep: {cep}, Endereço: {endereco}, Numero: {numero}, Bairro: {bairro}, Cidade: {cidade}, UF: {uf}, Telefone: {telefone}')


        def Validar_OS():
            print(f'Validando Placa {placa}')
            drv.get("https://site.com.br")
            t(3)
            drv.execute_script(f"LoadFrota('{placa}')")
            Desinstalacao = drv.execute_script('return document.querySelector("#DataDesinstalacao").value')
            if Desinstalacao == '':
                drv.get("https://site.com.br")
                drv.execute_script(f"LoadFrota('{placa}')")
                t(3)
                
                try:
                    alert = drv.switch_to.alert
                    mensagem = alert.text
                    print(mensagem)
                    if 'Já existe a O.S.' in mensagem:
                        alert.accept()
                        t(1.5)
                        Os_Gerada = drv.execute_script('return document.querySelector("#ServicoID").value')
                        Tipo = drv.execute_script('return document.querySelector("#F018_TipoID").value')
                        t(1.5)
                        if Tipo == "132":
                            drv.execute_script(f"DoCancelarOS()")
                            alert = drv.switch_to.alert
                            mensagem = alert.text
                            print(mensagem)
                            alert.accept()
                            t(1.5)
                            print("Os Em Triagem cancelada para abertura de OS de Rastreador")
                            t(2)
                            return 'OK'
                        elif Tipo == '194':
                            print(Tipo)
                            drv.get("https://site.com.br")
                            drv.execute_script(f"LoadFrota('{placa}')")
                            Desinstalacao = drv.execute_script('return document.querySelector("#DataDesinstalacao").value')
                            if Desinstalacao == '':
                                E.Email(placa, Os_Gerada, email)
                                print("Enviando email questionando a retirada do rastreador.")
                            else:
                                print(Desinstalacao)

                            tab.at[index, 'O.S Gerada'] = str(Os_Gerada)
                            tab.to_excel('teste.xlsx', index=False)
                            return "Aberto" 
                except:
                    print("Placa Validada")
                    t(2.5)
                    validacao = 0
                    return 'OK'

        def Gerar_OS():
            while True:
                drv.get("https://site.com.br")
                t(2)
                drv.execute_script(f"LoadFrota('{placa}')")
                t(3)
                drv.execute_script(f'document.querySelector("#CNPJ").value="{cnpj}"')
                drv.execute_script('document.querySelector("#RazaoSocial").click()')
                # drv.execute_script(f"LoadPessoa('{cnpj}')")
                t(4)
                print("CNPJ Carregado")
                razao_social = drv.execute_script('document.querySelector("#RazaoSocial").value')
                warning = wait.until(EC.presence_of_element_located((By.ID, "__msg__")))
                texto_warning = warning.text
                print(texto_warning)
                t(3)
                if 'bloqueado.' in texto_warning:
                    status = texto_warning
                    print(texto_warning) 
                    tab.at[index, 'O.S Gerada'] = str(texto_warning)
                    tab.to_excel('teste.xlsx', index=False) 
                    validacao = 1
                    break
                elif razao_social != '':
                    validacao = 0
                    break

            if validacao == 0:

                drv.execute_script(""" const campoEntrada = document.querySelector("#Email");
                                        campoEntrada.value = "";
                                                """)
                drv.execute_script(f'document.querySelector("#Email").value = "{email}"')
                KM_entrada = drv.execute_script('return document.querySelector("#KMEntrada").value')
                drv.execute_script(""" const campoEntrada = document.querySelector("#KMAtual");
                                        campoEntrada.value = "";
                                                """)
                drv.find_element(By.CSS_SELECTOR, '#KMAtual').send_keys(str(KM_entrada))
                print("KM Inserido")
                data_retorno = drv.find_element(By.ID, "Saida")
                data_retorno.send_keys(data_18h)
                tipo = 'document.querySelector("#F018_TipoID").value = "194";'
                drv.execute_script(tipo)
                print("Tipo Selecionado")
                Data = drv.find_element(By.CSS_SELECTOR, "#DataAgendamento")
                Data.send_keys(data_18h)
                print("Data Inserida")
                while True:
                    drv.execute_script(""" const campoEntrada = document.querySelector("#MatriculaConsultor");
                                        campoEntrada.value = "";
                                                """)
                    Consultor = drv.find_element(By.CSS_SELECTOR, "#MatriculaConsultor")
                    Consultor.send_keys('12345678')
                    Consultor.send_keys(Keys.TAB)
                    nome_consultor = drv.execute_script('return document.querySelector("#NomeConsultor").value')
                    print(nome_consultor)
                    if nome_consultor == "Vitor Antero da Cruz":
                        break
                
                drv.implicitly_wait(10)
                Descricao = drv.find_element(By.CSS_SELECTOR, "#Descricao")
                Descricao_Contato = drv.find_element(By.CSS_SELECTOR, "#DadosFornecedor")
                Descricao.send_keys(f'Local de Retirada do Rastreador.: Loja: {loja}, Cep: {cep}, Endereço: {endereco}, Numero: {numero}, Bairro: {bairro}, Cidade: {cidade}, UF: {uf}, Telefone: {telefone}')
                Descricao_Contato.send_keys(f'Local de Retirada do Rastreador.: Loja: {loja}, Cep: {cep}, Endereço: {endereco}, Numero: {numero}, Bairro: {bairro}, Cidade: {cidade}, UF: {uf}, Telefone: {telefone}')
                print("Descrições Inseridas")
                t(2)
                Salvar = drv.find_element(By.CSS_SELECTOR, '#OK')
                Salvar.click()
                t(1)
                Os_Gerada = drv.execute_script('return document.querySelector("#ServicoID").value')
                print(f'Salvo com sucesso OS gerada {Os_Gerada}')
                teste = 'OK'
        
                if Os_Gerada == '0': 
                    tab.at[index, 'O.S Gerada'] = 'Erro ao Gerar OS'
                    tab.to_excel('teste.xlsx', index=False)
                else:
                    tab.at[index, 'O.S Gerada'] = str(Os_Gerada)
                    tab.to_excel('teste.xlsx', index=False)
                return teste

        ## LANÇAR PEÇAS
        def Lancar_Pecas():    
            Placa_Sistema = drv.execute_script('return document.querySelector("#Placa").value')
            Os_Gerada = drv.execute_script('return document.querySelector("#ServicoID").value')
            for index, row in tab.iterrows():
                placa = row['Placa']
                if Os_Gerada == '0':
                    continue
                else:
                    drv.get(f"https://site.com.br/{Os_Gerada}")
                print(Os_Gerada)
                if placa == Placa_Sistema:            
                    print('Lançando Peças...') 
                    drv.execute_script('document.querySelector("#Garantia").click()')
                    Pmo = drv.find_element(By.ID, "PMO")
                    drv.execute_script('arguments[0].value = "1";', Pmo)
                    drv.execute_script(f'return DoThis(38868,"Instalacao/Desinstalacao-do-Rastreador",6000001836);')
                    drv.execute_script(f'return document.querySelector("#Quantidade").value="1"')
                    rateio = drv.find_element(By.CSS_SELECTOR, "#FinalidadeID")
                    rateio.send_keys(" AVARIAS - SERV. PROF. CONTRAT/PJ")
                    tipo_item = drv.find_element(By.CSS_SELECTOR, "#ItemTipo")
                    tipo_item.send_keys("CORRETIVO")
                    drv.execute_script(""" document.querySelector("#hidden > input.CodigoFIPE").value = 1
                                    DoInsert()
                                    """)
                    drv.implicitly_wait(10)
                    break

            # String dos valores
            string = drv.execute_script("return document.querySelector('html > body > table > tbody > tr > td > table > tbody > tr > td > span > span').textContent")
            # Expressões regulares para extrair os valores
            mao_de_obra_regex = r"Mão de obra: R\$ (\d+,\d+)"
            peca_regex = r"Peça: R\$ (\d+,\d+)"
            # Extrair o valor de mão de obra
            mao_de_obra_match = re.search(mao_de_obra_regex, string)
            valor_mao_obra = mao_de_obra_match.group(1) if mao_de_obra_match else "0,00"
            # Extrair o valor da peça
            peca_match = re.search(peca_regex, string)
            valor_peca = peca_match.group(1) if peca_match else "0,00"
            valor_mao_obra = valor_mao_obra.replace(',', '.')
            valor_peca = valor_peca.replace(',', '.')
            valor_total = float(valor_mao_obra) + float(valor_peca)
            print(valor_total)
            drv.execute_script(""" const campoEntrada = document.querySelector("#ValorOrcado");
                                        campoEntrada.value = "";
                                                """
                               )
            Valor_Orcado = drv.find_element(By.CSS_SELECTOR, "#ValorOrcado")
            Valor_Orcado.send_keys(valor_total)
            drv.execute_script(f'DoSalvarOrcamento({Os_Gerada})')

            
            drv.get("https://site.com.br")
            drv.execute_script(f"LoadOS('{Os_Gerada}')")
            drv.execute_script('document.querySelector("#ui-id-3").click()')
            drv.execute_script(f'DoAutorizar({Os_Gerada}, $("#Saida", parent.document).val())')
            alert = drv.switch_to.alert
            mensagem = alert.text
            print(mensagem)
            if 'Aguarde envio ao SAP' in mensagem:
                alert.accept()
            E.Email(placa, Os_Gerada, email)

        def Os_complementar():
            Os_Gerada = drv.execute_script('return document.querySelector("#ServicoID").value')
            drv.get("https://site.com.br")
            drv.execute_script(f"LoadOS('{Os_Gerada}')")
            drv.execute_script('document.querySelector("#frm_os > table > tbody > tr:nth-child(11) > td > span:nth-child(1) > input:nth-child(2)").click()')
            t(2)

        cont = 0
        while True:
            if cont == 0:
                Prestador = prestador1
                Valor = Valor1
                email = email1
                cnpj = cnpj1
            elif cont == 1:
                Prestador = prestador2
                Valor = Valor2
                email = email2
                cnpj = cnpj2.zfill(14)
            else:
                Prestador = prestador3  
                Valor = Valor3 
                email = email3
                cnpj = cnpj3.zfill(14)
            
            try:
                if np.isnan(Prestador):
                    break
            except Exception as e:
                print(e)

            if cont != 0:
                Os_complementar()


            Status = Validar_OS()
            if Status == 'OK':
                Gerar_OS()
                Lancar_Pecas()
                cont += 1
            else:
                break
        
            if cont == 2:
                break

            if Os_Gerada == '0': 
                tab.at[index, 'O.S Gerada'] = 'dsaasd' 
            else:
                tab.at[index, 'O.S Gerada'] = str(Os_Gerada)
            tab.to_excel('teste.xlsx', index=False)
    tab.to_excel('teste.xlsx', index=False) 
    print('inserido com sucesso na base')
Criar_OS()