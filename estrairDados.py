from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

#acessar site
driver = webdriver.Google()
driver.get('https://www.google.com')
#Extrair todos os titulos
titulos = driver.find_elements(By.XPATH, "//a[@class='nome-produto']")
#Extrrair preços
precos = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

#criar planilha
workbook = openpyxl.Workbook()
#criar pagina produtos
workbook.create_sheet('produtos')
#selecione a pagina produtos
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'




#Inserir na Planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])


#salvar planilha

workbook.save('produtos.xlsx')