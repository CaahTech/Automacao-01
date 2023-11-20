from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# ACESSAR O SITE>
driver = webdriver.Chrome()
driver.get('https://www.novaliderinformatica.com.br/computadores-gamers')


# EXTRAIR TODOS OS TÍTULOS>

titulos = driver.find_elements(By.XPATH, "//a[@class='nome-produto']")

# EXTRAIR TODOS OS PREÇOS>

precos = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

# CRIANDO A PLANILHA>

workbook = openpyxl.Workbook()

# CRIANDO A PÁGINA 'PRODUTOS'>

workbook.create_sheet('produtos')

# SELECIONO A PÁGINA PRODUTOS>
sheet_produtos = workbook['produtos']
sheet_produtos['A1'].value = 'Produto'
sheet_produtos['B1'].value = 'Preço'


# INSERIR OS TITULOS E PREÇOS NA PLANILHA>

for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulos.text, preco.text])
    workbook.save('produtos.xlsx')

# COMO ENTREGAR PARA O CLIENTE>
