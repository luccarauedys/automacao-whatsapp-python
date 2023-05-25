import os
import time
import pandas as pd
import urllib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ler a base de dados
tabela = pd.read_excel('./Envios.xlsx')

chrome = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
chrome.get('https://web.whatsapp.com/')

# esperar o whatsapp carregar após leitura do qr code
while len(chrome.find_elements(By.ID, 'side')) == 0:
    time.sleep(1)
time.sleep(2)

# para cada linha da tabela, fazer o seguinte:
for linha in tabela.index:
    nome = tabela.loc[linha, 'nome']
    mensagem = tabela.loc[linha, 'mensagem']
    arquivo = tabela.loc[linha, 'arquivo']
    telefone = tabela.loc[linha, 'telefone']

    texto = mensagem.replace('fulano', nome)
    texto = urllib.parse.quote(texto)

    link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
    chrome.get(link)

    # esperar o whatsapp carregar
    while len(chrome.find_elements(By.ID, 'side')) == 0:
        time.sleep(1)
    time.sleep(2)

    # verificar se o número é válido (não aparece mensagem de erro)
    if len(chrome.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) == 0:
        chrome.find_element(
            By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        time.sleep(2)

        # verificar se existe aquivo para anexar
        if arquivo != "N":
            caminho_completo = os.path.abspath(f"arquivos/{arquivo}")

            botao_anexar = chrome.find_element(
                By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span')
            botao_anexar.click()

            input_anexar = chrome.find_element(
                By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input')
            input_anexar.send_keys(caminho_completo)
            time.sleep(5)

            botao_enviar = chrome.find_element(
                By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
            botao_enviar.click()

        time.sleep(5)
