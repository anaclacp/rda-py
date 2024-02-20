import openpyxl
from urllib.parse import quote
import webbrowser
import keyboard
from time import sleep

# *** Requisitos de Funcionamento ***
# o seu whatsapp precisa estar logado no google
# precisa colocar o local de onde está localizada a tabela excel na linha 15
# para editar o texto exibido é só escrever dentro das aspas simples na linha 23

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('')
pagina_contatos = workbook['Sheet1']



for linha in pagina_contatos.iter_rows(min_row=1):
    #nome = linha[0].value
    telefone = linha[0].value
    mensagem = f''' '''

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    webbrowser.open(link_mensagem_whatsapp)
    sleep(30)
    try:
        keyboard.press_and_release('return')
        sleep(10)
        keyboard.press_and_release('command + w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {telefone}')
        with open('erros.csv','a', newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{telefone}')


