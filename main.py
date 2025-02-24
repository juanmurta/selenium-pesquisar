from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import win32com.client as win32

# criar um navegador
nav = webdriver.Chrome()

# importar/visualizar a base de dados
tabela_produtos = pd.read_excel('buscas.xlsx')


def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos


def verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome):
    tem_todos_termos_produtos = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            tem_todos_termos_produtos = False
    return tem_todos_termos_produtos


def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # entrar no google
    nav.get('https://www.google.com/')
    nav.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/textarea').send_keys(produto, Keys.ENTER)

    # entrar na aba shopping
    nav.find_element('css selector', '.LatpMc.nPDzT.T3FoJb').click()

    # pegar os precos do produto
    lista_resultados = nav.find_elements('class name', 'i0X6df')

    for resultado in lista_resultados:
        nome = resultado.find_element('class name', 'tAxDx').text
        nome = nome.lower()

        # analisar se ele nao tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # analisar se ele tem todos os termos do nome do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)

        # selecionar so os elementos que tem_termos_banidos = False e ao mesmo tempo tem_todos_termos = True
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = resultado.find_element('css selector', '.a8Pemb.OFFNJ').text  # observação para mim caso a classe tenha espaco usar o css selector usando o ponto para cada espaco
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
            preco = float(preco)

        # verificar se o preco ta entre o preco_minimo e preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                # selecionando o elemento filho para voltar ao elemento pai e buscar o link
                elemento_referencia = resultado.find_element('class name', 'bONr3b')
                elemento_pai = elemento_referencia.find_element('xpath', '..')
                link = elemento_pai.get_attribute('href')
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas


def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # entrar no buscape
    nav.get('https://buscape.com.br/')
    nav.find_element('xpath', '/html/body/div[1]/main/header/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(
        produto, Keys.ENTER)

    # pegar o resultado
    time.sleep(2)
    lista_resultados = nav.find_elements('class name', 'ProductCard_ProductCard_Inner__gapsh')

    for resultado in lista_resultados:
        nome = resultado.find_element('class name', 'ProductCard_ProductCard_Name__U_mUQ').text
        nome = nome.lower()

        # analisar se ele nao tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # analisar se ele tem todos os termos do nome do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produtos(lista_termos_nome_produto, nome)

        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = resultado.find_element('class name', 'Text_MobileHeadingS__HEz7L').text
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
            preco = float(preco)

        # verificar se o preco ta entre o preco_minimo e preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                link = resultado.get_attribute('href')
                lista_ofertas.append((nome, preco, link))

    return lista_ofertas


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    # pesquisar pelo produto
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        lista_ofertas_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['Produto', 'Preco', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, lista_ofertas_google_shopping])
    else:
        lista_ofertas_google_shopping = None

    lista_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_buscape:
        lista_buscape = pd.DataFrame(lista_buscape, columns=['Produto', 'Preco', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, lista_buscape])
    else:
        lista_buscape = None

# exportando para excel
tabela_ofertas.to_excel('Ofertas.xlsx', index=False)

# enviando email
if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'juanmurta@gmail.com'
    mail.Subject = 'Ofertas'
    mail.HTMLBody = f'''
    <p>Prezados</p>
    <p>Encontramos alguns produtos em ofertas</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Att.,</p>
    '''
    mail.Send()