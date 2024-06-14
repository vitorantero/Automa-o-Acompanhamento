import win32com.client
import os

def Email(placa, Os_Gerada, email):
    user = os.getlogin()
    diretorio_atual = os.getcwd()
    pasta_anexo = os.path.join(diretorio_atual, 'Anexo')
    anexo_logo = os.path.join(pasta_anexo, 'logo.png')

    conta = 'seuemail@gmail.com.br'
    destinatario = email
    assunto = 'Placa'
    placa = placa
    chamado = Os_Gerada
    print(placa)
    print(chamado)
    data_previsao = 'Digite a Data aqui!'

    outlook = win32com.client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    account = namespace.Accounts.Item(conta)
    mail = outlook.CreateItem(0)
    mail.SendUsingAccount = account
    mail.HTMLBody = str("""
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        background-color: #f5f5f5;
                        padding: 20px;
                    }
                    .movida{
                        width: 50%;
                        height: 40%;
                        margin-bottom: 5px;
                    }
                    #info{
                        background: linear-gradient(to bottom, #fffffff3, #f2f2f2d3);
                        border: 1px solid #dddddd;
                        padding: 10px;
                        margin-bottom: 5px;
                        text-align: center;
                        font-size: 1.3rem;
                        }
                    .container {
                        max-width: 800px;
                        margin: 0 auto;
                        background-color: #ffffff;
                        padding: 20px;
                        border-radius: 5px;
                        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
                        text-align: center;
                        margin-top: 25px;
                    }
                    .question {
                        font-size: 18px;
                        margin-bottom: 35px;
                        text-align: center;
                    }
                    .options {
                        display: flex;
                        flex-direction: column;
                        max-width: 800px;
                        margin: 0 auto;
                        background-color: #ffffff;
                        padding: 20px;
                        border-radius: 5px;
                        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
                        text-align: center;
                        margin-top: 25px;
                    }
                    .whatsapp {
                        margin-top: 20px;
                        text-align: center;
                    }
                    .whatsapp a {
                        display: inline-block;
                        padding: 12px 43px;
                        background-color: #25D366;
                        color: #ffffff;
                        text-decoration: none;
                        border-radius: 4px;
                        font-size: 16px;
                    }

                    .whatsapp a:hover {
                        background-color: #00c749;
                        padding: 12px 50px;
                    }

                    .message {
                        margin-top: 20px;
                        text-align: center;
                        font-size: 14px;
                        color: #888888;
                    }
                    .space{
                        display: table-column;
                        margin: 12px;
                    }
                    .red{
                        background-color: #ff3b30;
                        color: white;
                        padding: 15px 53px;
                        border-radius: 3px;
                        text-align: center;
                        text-decoration: none;
                        display: inline-block;
                        font-size: 16px;
                    }
                    .red:hover{
                        padding: 15px 60px;
                    }
                    .green{
                        background-color: rgb(5, 192, 5);
                        color: white;
                        padding: 15px 53px;
                        text-align: center;
                        border-radius: 3px;
                        text-decoration: none;
                        display: inline-block;
                        font-size: 16px;
                    }
                    .green:hover{
                        padding: 15px 60px;
                    }
                    .orange{
                        background-color: orange;
                        color: white;
                        padding: 15px 45px;
                        border-radius: 3px;
                        text-align: center;
                        text-decoration: none;
                        display: inline-block;
                        font-size: 16px;
                    }
                    .orange:hover{
                        padding: 15px 50px;
                    }
                    .box{
                    background: linear-gradient(to bottom, #ffffff, #f2f2f2);
                    border: 1px solid #dddddd;
                    padding: 10px;
                    margin-bottom: 5px;
                    font-size: 14px;
                    text-align: center;
                    font-size: 17px;
                    }
                    .nome_movida{
                        color:rgb(255, 136, 0);
                        font-size: 100px;
                        -bottom: 15px;
                        font-family:cursive;
                    }
                    .movida{
                    margin-bottom: 50px;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                <img src="logo.png" alt="Movida" class="movida">
                <div id="info">Placa: """+str(placa)+""" - OS: """+str(chamado)+"""</div>
                <h1 id="titulo">Qual a situação do veículo?</h1>              
                    <h3>Caro fornecedor, por favor selecione uma das opções abaixo para fornecer o status atualizado da solicitação em andamento</h4>
                        <br> 
                        <a class="orange" href="mailto:seuemail@gmail.com.br?subject=Situação do veículo PLACA: """+str(placa)+"""  -  CHAMADO: """+str(chamado)+""" / DATA PREVISTA """+ str(data_previsao) +""" PRECISA SER ALTERADA GTF2709 """"&body=Situação do veículo PLACA: """+str(placa)+"""  -  CHAMADO: """+str(chamado)+""" / DATA PREVISTA """+ str(data_previsao) +""" PRECISA SER ALTERADA"
                        >Previsão de Retirada!</a>
                        <br>
                        <br>
                        <a class="green" href="mailto:seuemail@gmail.com.br?subject=Situação do veículo PLACA: """+str(placa)+"""  -  CHAMADO: """+str(chamado)+" """"&body=Olá! Venho lhe informar que a retirada do rastreador do veículo com placa: """+str(placa)+"""  -  referente à Ordem de Serviço (OS) número """+str(chamado)+""" foi concluída com sucesso."
                        >Retirada concluída!</a>
                        <br>
                        <br>
                        <a class="red" href="mailto:seuemail@gmail.com.br?subject=Situação do veículo PLACA: """+str(placa)+"""  -  CHAMADO: """+str(chamado)+""" / DATA PREVISTA """+ str(data_previsao) +""" PERMANECE A MESMA GTF2709 """"&body=Informamos que a retirada do rastreador do veículo com placa """+str(placa)+"""  - referente à Ordem de Serviço (OS) número """+str(chamado)+""" / ainda não foi concluída e está prevista para """+ str(data_previsao) +""" "
                        >Retirada Frustrada!</a>
                        <br>
                    <div class="whatsapp">
                        <a class="button" href="https://api.whatsapp.com/send?phone=5511999999999">Contato via WhatsApp</a>
                    </div>
                    <div class="message">
                        Lembre-se de não alterar nenhum caractere do e-mail padrão ao clicar no botão, exceto se o veículo não estiver pronto.
                    </div>
                </div>
            </body>
            </html>


        """)

    mail.To = destinatario
    mail.Subject = assunto
    mail.Sender = outlook.Session.CurrentUser.Address
    attachment_image = mail.Attachments.Add(anexo_logo)
    mail.Send()
    print('Email Enviado com Sucesso')

