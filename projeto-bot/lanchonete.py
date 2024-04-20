import webbrowser
import openpyxl
from time import sleep
from pyautogui import *


#webbrowser.open('https://web.whatsapp.com/')
#sleep(30)

#acessando a planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
#acessando sheet1/contatos nos colchetes e salvando na variável pagina_clientes
pagina_clientes = workbook['Contatos']
#lendo informações de cada linha da planilha a partir da segunda linha
for linha in pagina_clientes.iter_rows(min_row=2):
    #pegando os valores de nome e telefone de cada linha a partir do indice
    nome = linha[0].value
    telefone= linha[1].value

    mensagem = f"Olá {nome} seu número é {telefone}"
    print(mensagem)
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={mensagem}'
    
    
    
#webbrowser.open(link_mensagem_whatsapp)
#input('')
