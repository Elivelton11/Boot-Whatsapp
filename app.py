"""
Preciso automatizar minhas mensgens p/ meus clientes e gostaria de saber os valores 
Mensagens de cobrança em determinado dia com cliente com vencimento diferente

"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
# ler a planilha e guardar as informações

webbrowser.open("https://web.whatsapp.com/")
sleep (30)
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']
for linha in pagina_clientes.iter_rows(min_row=2):
# Criar links personalizados do whatsapp e enviar mensagens para cada cliente
#com base nos dados de cada planilha 
    NOME = linha[0].value
    TELEFONE = linha[1].value
    VENCIMENTO = linha[2].value
    try:
        mensagem = f'Olá {NOME} seu boleto venceu dia {VENCIMENTO} para realizar o pagamento clique no link www.google.com'
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={TELEFONE}&text={quote(mensagem)}'
        webbrowser.open (link_mensagem_whatsapp)
        sleep(15)
        seta = pyautogui.locateCenterOnScreen('seta.PNG')
        sleep(15)
        pyautogui.click(seta[0], seta[1])
        sleep(15)
        pyautogui.hotkey('Ctrl','w')
        sleep(15)
    except:
        print(f'Não foi possível mandar mensagem para {NOME}')     
        with open('erros.csv','a',newline='', encondig='utf-8') as arquivo:
            arquivo.write(f'{NOME},{TELEFONE}')