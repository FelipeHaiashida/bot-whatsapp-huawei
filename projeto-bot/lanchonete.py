import webbrowser
import openpyxl
from time import sleep
import pyautogui
from urllib.parse import quote


#webbrowser.open('https://web.whatsapp.com/')
#sleep(30)

    #Acessando a planilha
workbook = openpyxl.load_workbook('clientes 1.xlsx')
    #Acessando sheet1/contatos nos colchetes e salvando na variável pagina_clientes
pagina_clientes = workbook['Contatos']
    #Lendo informações de cada linha da planilha a partir da segunda linha
for linha in pagina_clientes.iter_rows(min_row=2):
        #Pegando os valores de nome e telefone de cada linha a partir do indice
    nome = linha[0].value
    telefone= linha[1].value

    mensagem = f"Olá {nome} seu número é {telefone}"
    print(mensagem)
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        click = pyautogui.locateCenterOnScreen('seta.png')
        sleep(10)
        pyautogui.click(click[0], click[1])
        sleep(10)
        pyautogui.hotkey('ctrl', 'w')
        sleep(10)
        fechar = pyautogui.locateCenterOnScreen('fechar.png')
        sleep(10)
    except:
        print(f"Não foi possivel contatar {nome}")
        with open("erros.csv", "a",newline='', encoding='utf-8') as arquivo:
            arquivo.write(f"{nome}, {telefone}")
        
