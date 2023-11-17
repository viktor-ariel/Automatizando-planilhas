import openpyxl
import pyperclip #para manter os acentos
import pyautogui
from time import sleep #para colocar uma pausa

#Entrar na planilha
workbook = openpyxl.load_workbook('nome da planilha.xlsx')

sheet123 = workbook['produtos'] #nome da sheets
#copiar informações de uum campo e colar no seu campo correspondente
for linha in sheet123.iter_rows(min_row=2):
    #campo 1
    valorDaCelula1 = linha[0].value
    pyperclip.copy(valorDaCelula1)
    #pip install mouseinfo para saber onde deve clicar, >python, >from mouseinfo import mouseinfo, 
    # >mouseInfo()
    pyautogui.click(1535,305,duration=1)
    pyautogui.hotkey('ctrl','v')

    #campo 2
    valorDaCelula2 = linha[1].value
    pyperclip.copy(valorDaCelula2)
    pyautogui.click(1235,505,duration=1)
    pyautogui.hotkey('ctrl','v')


    sleep(3)

    valorDaCelula3 = linha[2].value
    pyautogui.click(35,505,duration=1)
    #ler uma info da planilha
    #se for a clicar em a 
    #se for b clicar em b 
    #se for c clicar em c 
    if valorDaCelula3 == 'a':
        pyautogui.click(35,505,duration=1)
    elif valorDaCelula3 == 'b':
        pyautogui.click(35,505,duration=1)
    else:
        pyautogui.click(35,505,duration=1)

    valorDaCelula4 = linha[3].value
    valorDaCelula5 = linha[4].value
    valorDaCelula6 = linha[5].value
    valorDaCelula7 = linha[6].value
    valorDaCelula8 = linha[7].value


    #clicar no botão de concluir
    pyautogui.click(35,505,duration=1)

    #clicar no botão de concluir
    pyautogui.click(35,505,duration=1)

#repetir os passos até preencheer os campos dessa página
#clicar em proxima
#repetir os mesmos passos e ir  pra página 2
#repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em ok
#clicar em ok e finalizar os processos
#clicar em ok na mensagem para salvar no banco
#clicar em add mais um e repetir o processo até finalizar a planilha
