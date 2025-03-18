import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 
import pandas as pd 

webbrowser.open('https://web.whatsapp.com/')
sleep(5)

# Vai ler a planilha e puxar nome e telefone para enviar a mensagem
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['DM25']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone
    nome = linha[0].value
    telefone = linha[1].value
    
    mensagem = f'Opa, {nome} Tudo bem? Passando para lembrar da mensalidade de Março. Valor: R$ 40,00. Pix: jcibalneariocamboriu@jci.org.br. Peço que me envivem o comprovante de pagamento por aqui, com a informação MENSALIDADE na descrição. Aos que já pagaram, favor desconsiderar mensagem. Qualquer dúvida estou à disposição.'

    # Links personalizados do whatsapp e enviar mensagens para cada pessoa
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.click( x=4150, y=1025)
        sleep(2)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    

