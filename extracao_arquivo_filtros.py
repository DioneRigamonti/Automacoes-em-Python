from selenium import webdriver
from selenium.webdriver.chrome.service import Service
#from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
import shutil
from time import sleep
import datetime as dt
from datetime import timedelta
from zipfile import ZipFile
import xlwings as xw

# - - - Data Input - INCOMING - - -#

data_atual = dt.datetime.today()
data_retroativa = data_atual - timedelta(days=7) #periodo definido -7 dias
data_atual = dt.datetime.strftime(data_atual,r"%d/%m/%Y") # tratamento da data, de um tipo data para um tipo string.
data_retroativa = dt.datetime.strftime(data_retroativa,r"%d/%m/%Y")

# - - - Data Log - - - #

data = dt.datetime.today()
dia = data.day
mes = data.month
ano = data.year
hora = data.hour
minuto = data.minute

# - - - Caminho - - - #

servico = Service(r'CAMINHO ONDE O ARQUIVO DO WEBDRIVER ESTA')
caminho_automacao = r'CAMINHO DO ARQUIVO.PY'
caminho_incoming = r'CAMINHO PASTA TEMPORARIA PARA O LOG'
caminho_conv = r'CAMINHO DA MACRO EXCEL PARA CONVERSÃO DO ARQUIVO BAIXADO'
#servico = Service(ChromeDriverManager().install()) # auto instalação do webdriver

# - - - Argumentos para possiveis erros - - - #

option = Options()
option.add_argument("--dns-prefech-disable")
option.add_experimental_option('extensionLoadTimeout',30)
option.add_argument("--disable-gpu")
option.add_argument("--disable-extensions")
option.add_argument("--star-maximized")
option.add_argument("enable-automation")
option.add_argument("--no-sandbox")
option.add_argument("--disable-infobars")
option.add_argument("disable-dev-shm-usage")
option.add_argument("--disable-browser-side-navigation")
option.add_argument("--disable-features=VizDisplayCompositor")
option.add_experimental_option("prefs",{"download.default_directory":f'{caminho_incoming}'}) # importante para fazer o download para uma pasta especifica
option.add_argument("--headless=new") #segundo plano

# - - - Inicio do Processo - - - #

navegador = webdriver.Chrome(service=servico,options=option)
print('STC - INCOMING')
print(f'De: {data_retroativa} até: {data_atual}')

# - - - Exclusão da pasta - - - #

sleep(0.5)
print(f'Exluindo pasta . . .')
if f'incoming_temp' in os.listdir(caminho_automacao):
    shutil.rmtree(caminho_incoming)

    while "incoming_temp" in os.listdir(caminho_automacao):
        sleep(1)

else:
    pass

# - - - Criação da pasta - - - #

sleep(0.5)
print('Criando pasta . . .')
os.mkdir(caminho_incoming)

# - - - Outros - - - #

navegador.set_page_load_timeout(3600) # Primeira solução do erro timeout
#navegador.implicitly_wait(3600) # caso de erro de timeout descomentar isso aqui

# - - - Login - - - #

sleep(0.5)
print('Logando . . .')
navegador.get('SITE QUE VAI SER ACESSADO')
sleep(0.5)
navegador.find_element(By.ID,'user').send_keys('USUARIO')
navegador.find_element(By.ID,'password').send_keys('SENHA')
navegador.find_element(By.CLASS_NAME,'button').click()

# - - - Filtros checkbox- - - #

print('Aplicando filtros checkbox . . .')
sleep(1)
navegador.find_element(By.ID,'filtro1').click()
sleep(1)
navegador.find_element(By.ID,'filtro2').click()
sleep(1)
navegador.find_element(By.ID,'filtro3').click()
sleep(1)
navegador.find_element(By.ID,'filtro4').click()
sleep(1)
navegador.find_element(By.ID,'filtro5').click()
sleep(1)
navegador.find_element(By.CLASS_NAME,'button1').click()# play

# - - - Log - - - #

sleep(1)
with open(f'CAMINHO PASTA LOG{dia}.{mes}.{ano}','a') as arquivo:
    arquivo.write(f'\nFim log {dia}/{mes}/{ano}')

arquivo.close()

# - - - Auardando o arquivo ser baixado - - - #

contagem = 0

while contagem < 600:
    if f'ARQUIVO BAIXADO.zip' not in os.listdir(caminho_incoming):
        contagem += 1
        sleep(1)

    else:
        break

sleep(0.5)
print('Download concluido!')

# - - - Extração do arquivo zip - - - #

sleep(1)
print('Extraindo arquivo da pasta .zip')

with ZipFile(f'{caminho_incoming}\\NOME DA PASTA.ZIP','r') as zip:
    zip.extract('NOME DO ARQUIVO QUE ESTÁ DENTRO DA PASTA.ZIP',caminho_incoming)

sleep(1)
zip.close()

print('Extração concluida!')

# - - - Convertendo arquivo slk para xlsx - - - #

print('Convertendo arquivo . . .')

if 'arquivo.slk' in os.listdir(caminho_incoming):
    wb = xw.Book(caminho_conv)
    play = wb.macro('NOME DA MACRO')
    play()

sleep(0.5)
print('Fim do processo!')
