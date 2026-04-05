from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os
import time

load_dotenv()

def criar_navegador ():
    """
    Cria e configura uma instância do navegador com as opções
    necessárias para a execução no Linux e para evitar a detecção de bots.

    :return: webdriver.Chrome: instância do navegador configurada e pronta para uso.
    """
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    service = Service(ChromeDriverManager().install())

    return webdriver.Chrome(service=service, options=options)

def busca_google_shopping(produto, termos_banidos, preco_min, preco_max):
    """
    Realiza a busca de um produto no Google Shopping e retorna as ofertas
    que atendem aos critérios de filtro informados.

    :param produto (str): Nome do produto a ser pesquisado.
    :param termos_banidos (str): Palavras separadas por espaço que, se presentes
        no nome do produto, fazem o resultado ser ignorado.
    :param preco_min (float): Preço mínimo aceito para a oferta.
    :param preco_max (float): Preço máximo aceito para a oferta.
    :return: list: lista de tuplas (nome, preco, link) com as ofertas encontradas.
    """

    navegador = criar_navegador()
    espera = WebDriverWait(navegador, 20)

    navegador.get("https://google.com")

    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")

    pesquisa = navegador.find_element(By.NAME, 'q')
    pesquisa.click()
    pesquisa.send_keys(produto)
    pesquisa.send_keys(Keys.ENTER)
    time.sleep(5)

    elementos = navegador.find_elements(By.CLASS_NAME, 'C6AK7c')
    for elemento in elementos:
        if elemento.text == "Shopping":
            elemento.click()
            break

    time.sleep(3)
    lista_ofertas = []
    preco_min = float(preco_min)
    preco_max = float(preco_max)
    try:
        lista_resultados = espera.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "mnr-c")))
        for resultado in lista_resultados:
            nome = resultado.find_element(By.CLASS_NAME, "bXPcId").text
            nome = nome.lower()

            tem_termos_banidos = any(palavra in nome for palavra in lista_termos_banidos)
            tem_termos_produto = all(palavra in nome for palavra in lista_termos_produto)

            if not tem_termos_banidos and tem_termos_produto:
                try:
                    preco = resultado.find_element(By.CLASS_NAME, "VbBaOe").text
                    preco = preco.replace("R$", "").replace(".", "").replace(",", ".")
                    preco = float(preco)

                    if preco_min <= preco <= preco_max:
                        link = resultado.find_element(By.CLASS_NAME, "plantl").get_attribute("href")
                        lista_ofertas.append((nome, preco, link))

                except:
                    continue

    except:
        espera_2 = WebDriverWait(navegador, 20)
        lista_resultados = espera_2.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "UC8ZCe")))

        for resultado in lista_resultados:
            try:
                nome = resultado.find_element(By.CLASS_NAME, "gkQHve").text
                nome = nome.lower()
            except:
                continue

            tem_termos_banidos = any(palavra in nome for palavra in lista_termos_banidos)
            tem_termos_produto = all(palavra in nome for palavra in lista_termos_produto)

            if not tem_termos_banidos and tem_termos_produto:
                try:
                    preco = resultado.find_element(By.CLASS_NAME, "lmQWe").text
                    preco = preco.replace("R$", "").replace(".", "").replace(",", ".")
                    preco = float(preco)

                    if preco_min <= preco <= preco_max:
                        resultado.click()
                        time.sleep(2)
                        link = navegador.current_url
                        navegador.back()
                        time.sleep(2)
                        lista_ofertas.append((nome, preco, link))
                except:
                    continue

    navegador.quit()
    return lista_ofertas

def busca_buscape(produto, termos_banidos, preco_min, preco_max):
    """
    Realiza a busca de um produto no Buscapé e retorna as ofertas
    que atendem aos critérios de filtro informados.

    :param produto (str): Nome do produto a ser pesquisado.
    :param termos_banidos (str): Palavras separadas por espaço que, se presentes
        no nome do produto, fazem o resultado ser ignorado.
    :param preco_min (float): Preço mínimo aceito para a oferta.
    :param preco_max (float): Preço máximo aceito para a oferta.
    :return: list: lista de tuplas (nome, preco, link) com as ofertas encontradas.
    """

    navegador = criar_navegador()
    espera = WebDriverWait(navegador, 20)

    navegador.get("https://buscape.com.br/")

    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")

    try:
        fecha = navegador.find_element(By.CLASS_NAME, "icon-container_icon-close__bcVjc")
        fecha.click()
    except:
        pass

    pesquisa = espera.until(EC.presence_of_element_located((By.ID, "searchInput")))
    pesquisa.click()
    pesquisa.send_keys(produto)
    pesquisa.send_keys(Keys.ENTER)
    time.sleep(5)

    lista_ofertas = []
    preco_min = float(preco_min)
    preco_max = float(preco_max)
    lista_resultados = espera.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "Hits_ProductCard__Bonl_")))
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, "Name_OrqProductCard_Name__KsaTM").text
        nome = nome.lower()

        tem_termos_banidos = any(palavra in nome for palavra in lista_termos_banidos)
        tem_termos_produto = all(palavra in nome for palavra in lista_termos_produto)

        if not tem_termos_banidos and tem_termos_produto:
            preco = resultado.find_element(By.CLASS_NAME, "Price_OrqProductCard_Price__TNBZB").text
            preco = preco.replace("R$", "").replace(".", "").replace(",", ".")
            preco = float(preco)

            if preco_min <= preco <= preco_max:
                link = resultado.find_element(By.CLASS_NAME, "ClickableArea_OrqProductCard_ClickableArea__jkrb3").get_attribute("href")
                lista_ofertas.append((nome, preco, link))

    navegador.quit()
    return lista_ofertas

def criar_tabela_ofertas(tabela_produtos):
    """
    Itera sobre a tabela produtos, chama as funções de busca para cada
    produto e consolida todas as ofertas encontradas num arquivo Excel.

    :param tabela_produtos (DataFrame): Tabela com colunas Nome, Termos Banidos,
        Preço Mńimo e Preço Máximo.
    :return (str): Caminho do arquivo Excel gerado com as ofertas encontradas.
    """
    tabela_ofertas = pd.DataFrame()

    for linha in tabela_produtos.index:
        produto = tabela_produtos.loc[linha, "Nome"]
        termos_banidos = tabela_produtos.loc[linha, "Termos banidos"]
        preco_min = tabela_produtos.loc[linha, "Preço mínimo"]
        preco_max = tabela_produtos.loc[linha, "Preço máximo"]

        lista_ofertas_google_shopping = busca_google_shopping(produto, termos_banidos, preco_min, preco_max)
        if lista_ofertas_google_shopping:
            tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=["produto", "preco", "link"])
            tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping], ignore_index=True)

        lista_oferta_buscape = busca_buscape(produto, termos_banidos, preco_min, preco_max)
        if lista_oferta_buscape:
            tabela_buscape = pd.DataFrame(lista_oferta_buscape, columns=["produto", "preco", "link"])
            tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape], ignore_index=True)

    tabela_ofertas.to_excel("dataframes/ofertas.xlsx", index=False)

    return "dataframes/ofertas.xlsx"

def mandar_email(arquivo):
    """
    Envia um email com o arquivo de ofertas em anexo para o destinatário
    configurado nas variáveis de embiente.

    :param arquivo (str): Caminho do arquivo a ser enviado como anexo.
    """
    remetente = os.getenv("gmail_remetente")
    destinatario = os.getenv("gmail_destinatario")
    senha = os.getenv("gmail_senha")

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = "Ofertas encontradas hoje"
    msg.attach(MIMEText("Segue em anexo a tabela de ofertas", "plain"))

    with open(arquivo, "rb") as f:
        anexo = MIMEBase("application", "octet-stream")
        anexo.set_payload(f.read())
        encoders.encode_base64(anexo)
        anexo.add_header("Content-Disposition", "attachment; filename=ofertas.xlsx")
        msg.attach(anexo)

    with smtplib.SMTP("smtp.gmail.com", 587) as servidor:
        servidor.starttls()
        servidor.login(remetente, senha)
        servidor.send_message(msg)
        print("Email enviado com sucesso!")

tabela_produtos = pd.read_excel("dataframes/buscas.xlsx")
mandar_email(criar_tabela_ofertas(tabela_produtos))

#TODO:
# 3- CRIAR O ARQUIVO README