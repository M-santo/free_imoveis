from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime
import openpyxl

# Inicializar o driver do Chrome
driver = webdriver.Chrome()

# Acessar a página
driver.get('https://www.imoveismartinelli.com.br/pesquisa-de-imoveis/?locacao_venda=V&id_cidade%5B5D=21&finalidade=&dormitorio=&garagem=&vmi=&vma=&ordem=4%27')

# Encontrar os elementos com os preços e links
precos = driver.find_elements(By.XPATH, "//div[@class='card-valores']/div")
links = driver.find_elements(By.XPATH, "//a[@class='carousel-cell is-selected']")

# Carregar o arquivo Excel
workbook = openpyxl.load_workbook('imoveis.xlsx')
pagina_imoveis = workbook['imoveis']

# Iterar sobre os preços e links
for preco, link in zip(precos, links):
    # Formatar o preço
    preco_formatado = preco.text.split(' ')[1]
    
    # Obter o link pronto
    link_pronto = link.get_attribute('href')
    
    # Obter a data atual
    data_atual = datetime.now().strftime('%d/%m/%Y')
    
    # Adicionar as informações ao arquivo Excel
    pagina_imoveis.append([preco_formatado, link_pronto, data_atual])

# Salvar o arquivo Excel
workbook.save('imoveis.xlsx')

# Fechar o driver do Chrome
driver.quit()