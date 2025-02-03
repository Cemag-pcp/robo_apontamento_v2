from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import time
import re
import pandas as pd
import requests
import datetime
from io import StringIO

def login(nav):

    try:
        # logando
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="username"]'))).send_keys("Luan araujo")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo7")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

        print("Acessou a página de login")

    except Exception as e:
        print(f"Ocorreu um erro durante o login: {e}")

def menu_innovaro_1(nav):
    
    """
    Função para abrir ou fechar menu no innovaro do tipo 1
    :nav: webdriver
    """
    
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    menu=WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'menuBar-button-label')))
    time.sleep(2.5)
    menu.click()
    time.sleep(2.5)

def menu_innovaro_2(nav):
    
    """
    Função para abrir ou fechar menu no innovaro do tipo 2
    :nav: webdriver
    """
    
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1898143037"]/table/tbody/tr/td[2]'))).click()

    time.sleep(2)

def menu_apontamento(nav):
    
    nav.switch_to.default_content()
    
    #menu
    try:
        menu_innovaro_1(nav)
        print('Menu aberto')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(2)
    
    #Clicando em "Produção"
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(2)
    click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(2)
        
    #Clicando em "Controle de fábrica (SFC)"
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(2)
    click_producao = test_list.loc[test_list[0] == 'Controle de fábrica (SFC)'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(2)

    #Clicando em "Apontamento da produção"
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(2)
    click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(2)

def clicando_producao_menu_aberto(nav):

    iframes(nav)
    nav.switch_to.default_content()
    
    #menu
    try:
        menu_innovaro_1(nav)
        print('Menu aberto')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(2)

    #Clicando em "Apontamento da produção"
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(2)
    click_producao = test_list.loc[test_list[0] == 'Apontamento da produção'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(2)

def iframes(nav):
    
    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

def listar(nav, classe):

    try:

        lista_menu = nav.find_elements(By.CLASS_NAME, classe)

        elementos_menu = []

        for x in range(len(lista_menu)):
            a = lista_menu[x].text
            elementos_menu.append(a)

        test_lista = pd.DataFrame(elementos_menu)
        test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

        print("listou as opções do menu")

    except Exception as e:
        print(f"Ocorreu um erro durante a listagem de opções: {e}")

    return (lista_menu, test_lista)

def data_hoje():
    data_hoje = datetime.datetime.now()
    ts = pd.Timestamp(data_hoje)
    data_hoje = data_hoje.strftime('%d/%m/%Y')
    
    return(data_hoje)

def hora_atual():
    hora_atual = datetime.datetime.now()
    ts = pd.Timestamp(hora_atual)
    hora_atual = hora_atual.strftime('%H:%M:%S')
    
    return(hora_atual)

def aguardando_requisicao_google_sheets(segundos):

    for i in range(segundos,-1,-1):
        time.sleep(1)
        print(f"Aguardando {i} segundos para fazer uma nova chamada a API do google")

def mudar_visao(nav):

    # Usar a função para clicar no botão "postButton"
    xpath_botao = '//*[@id="producoes"]//div[@id="changeViewButton"]'
    classe_esperada = "grid-titleBar-button grid-titleBar-changeToFormButton-hover"

    clicar_ate_classe(nav, xpath_botao, classe_esperada)

def voltar_visao(nav):

    # Usar a função para clicar no botão "postButton"
    xpath_botao = '//*[@id="producoes"]//div[@id="changeViewButton"]'
    classe_esperada = "grid-titleBar-button grid-titleBar-changeToFormButton-active"

    clicar_ate_classe(nav, xpath_botao, classe_esperada)

def carregamento(nav):

    nav.switch_to.default_content()

    print('procurando carregamento 1')
    try:
        # Espera inicial para verificar se a mensagem de carregamento existe
        carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="statusMessageBox"]')))
        
        # Enquanto o elemento existir, continue verificando
        while True:
            print("Carregando...")
            try:
                # Aguarde novamente pela presença do elemento
                carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="statusMessageBox"]')))
            except TimeoutException:
                # Se o elemento não for encontrado, interrompa o loop
                break
    except TimeoutException:
        # Não há mensagem de carregamento inicial
        pass    

    print('procurando carregamento 2')
    try:
        # Espera inicial para verificar se a mensagem de carregamento existe
        carregamento =  WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="waitMessageBox"]')))
        
        # Enquanto o elemento existir, continue verificando
        while True:
            print("Carregando...")
            try:
                # Aguarde novamente pela presença do elemento
                carregamento =  WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="waitMessageBox"]')))
            except TimeoutException:
                # Se o elemento não for encontrado, interrompa o loop
                break
    except TimeoutException:
        # Não há mensagem de carregamento inicial
        pass  

    print('procurando carregamento 3')
    try:
        # Espera inicial para verificar se a mensagem de carregamento existe
        carregamento =  WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="progressMessageBox"]')))
        
        # Enquanto o elemento existir, continue verificando
        while True:
            print("Carregando...")
            try:
                # Aguarde novamente pela presença do elemento
                carregamento =  WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="progressMessageBox"]')))
            except TimeoutException:
                # Se o elemento não for encontrado, interrompa o loop
                break
    except TimeoutException:
        # Não há mensagem de carregamento inicial
        pass    

    print('procurando carregamento 4')
    try:
        # Espera inicial para verificar se a mensagem de carregamento existe
        carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_waitMessageBox"]')))
        
        # Enquanto o elemento existir, continue verificando
        while True:
            print("Carregando...")
            try:
                # Aguarde novamente pela presença do elemento
                carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_waitMessageBox"]')))
            except TimeoutException:
                # Se o elemento não for encontrado, interrompa o loop
                break
    except TimeoutException:
        pass 
    
    print('procurando carregamento 5')
    try:
        # Espera inicial para verificar se a mensagem de carregamento existe
        carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]')))
        
        # Enquanto o elemento existir, continue verificando
        while True:
            print("Carregando...")
            try:
                # Aguarde novamente pela presença do elemento
                carregamento = WebDriverWait(nav, 3).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content_statusMessageBox"]')))
            except TimeoutException:
                # Se o elemento não for encontrado, interrompa o loop
                break
    except TimeoutException:
        # Não há mensagem de carregamento inicial
        iframes(nav)
        pass

def verificar_se_erro(nav):

    nav.switch_to.default_content()

    time.sleep(2)

    error=None
    try:
        error = WebDriverWait(nav, 4).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div/div/span[1]'))).text
        confirm_button = WebDriverWait(nav, 4).until(EC.presence_of_element_located((By.XPATH, "//button[@id='confirm']")))
        time.sleep(1)
        confirm_button.click()
        time.sleep(1)
    # WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
    #     By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
    except:
        pass

    if error == None:
        try:
            error = WebDriverWait(nav, 4).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]/div[2]/table/tbody/tr[1]/td[2]/div'))).text
            confirm_button = WebDriverWait(nav, 4).until(EC.presence_of_element_located((By.XPATH, "//button[@id='confirm']")))
            time.sleep(1)
            confirm_button.click()
            time.sleep(1)
        except:
            pass

    return error

def exibindo_erro_na_planilha(nav,wks,row,erro,coluna_erro,valor_chave_do_apontamento=""):
    
    erro_com_a_chave = f"{erro} - {valor_chave_do_apontamento}"
    wks.update_acell(coluna_erro + str(row['index'] + 1), erro_com_a_chave)
    print("Saindo da aba")
    WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
        By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
    time.sleep(1.5)
    clicando_producao_menu_aberto(nav)

def sair_da_aba_e_voltar_ao_menu(nav):
    nav.switch_to.default_content()

    WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
        By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
    time.sleep(1.5)
    clicando_producao_menu_aberto(nav)

def erro_na_data(nav,wks,row,erro):

    wks.update_acell('F'+ str(row['index'] + 1), erro)

    verificar_se_erro(nav)

    iframes(nav)

    data_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DATA"]')))
    data_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    data_input.send_keys(Keys.BACKSPACE)
    time.sleep(1.5)
    data_input.send_keys(Keys.TAB)
    time.sleep(1.5)
    nav.switch_to.default_content()
    WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
        By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
    time.sleep(1.5)
    clicando_producao_menu_aberto(nav)

def clicar_em_add(nav):
    # Usar a função para clicar no botão "postButton"
    xpath_botao = '//*[@id="producoes"]//div[@id="insertButton"]'
    classe_esperada = "grid-titleBar-button grid-titleBar-newRecordButton-inactive"
    print("add")
    clicar_ate_classe(nav, xpath_botao, classe_esperada)

def preencher_classe(nav):
    
    iframes(nav)
    
    classe_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="CLASSE"]')))
    classe_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    classe_input.send_keys('Produção por Máquina')
    time.sleep(1.5)
    classe_input.send_keys(Keys.TAB)

def preencher_data_apontamento(nav,data):
    
    iframes(nav)

    data_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DATA"]')))
    data_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(2)
    data_input.send_keys(data)
    time.sleep(2)
    data_input.send_keys(Keys.TAB)

def preencher_pessoa(nav, apontamento_atual,funcionario=""):

    iframes(nav)

    apontado_pelo_supervisor = ["serra", "usinagem", "pintura", "estamparia", "montagem","corte"]

    matricula_funcionario = "4054" if funcionario == "" else funcionario

    if apontamento_atual in apontado_pelo_supervisor:
        if apontamento_atual == "pintura":
            matricula_funcionario = "4359"
        elif apontamento_atual == "estamparia" or apontamento_atual == "montagem":
            matricula_funcionario = str(funcionario).split('-')[0].strip()
        elif apontamento_atual == "corte":
            matricula_funcionario = "4322"
        else:
            matricula_funcionario = "4057"

    pessoa_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PESSOA"]')))
    pessoa_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(.5)
    pessoa_input.send_keys(matricula_funcionario)
    time.sleep(1.5)
    pessoa_input.send_keys(Keys.TAB)

def preencher_recurso(nav, codigo):

    iframes(nav)
    print("CÓDIGO")
    print(codigo)
    recurso_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="RECURSO"]')))
    recurso_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(.5)
    recurso_input.send_keys(codigo)
    time.sleep(3)
    recurso_input.send_keys(Keys.TAB)

def preencher_processo(nav, processo):

    error = None

    iframes(nav)

    processo_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PROCESSO"]')))
    processo_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    processo_input.send_keys(processo)
    time.sleep(1.5)
    processo_input.send_keys(Keys.TAB)

    # # Bloco de carregamento e erro 
    # carregamento(nav)

    # error=verificar_se_erro(nav)

    # # ser der erro, refaz o input
    # if error:
    #     iframes(nav)

    #     time.sleep(1.5)

    #     processo_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
    #         By.XPATH, '//*[@id="producoes"]//input[@name="PROCESSO"]')))
    #     processo_input.send_keys(Keys.CONTROL + 'A')
    #     time.sleep(1.5)
    #     processo_input.send_keys(processo)
    #     time.sleep(1.5)  

    # return error         

def preecher_qtd_produzida(nav, qt_produzida):
    
    iframes(nav)
    
    qt_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PRODUZIDO"]')))
    qt_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    qt_input.send_keys(qt_produzida)
    time.sleep(1.5)
    qt_input.send_keys(Keys.TAB)

    # verifica se a quantidade ta vazia:
    qt_input = WebDriverWait(nav, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="producoes"]//input[@name="PRODUZIDO"]'))
    ).get_attribute("value")

    if qt_input == '':
        print('quantidade vazio')
        qt_input.send_keys(Keys.CONTROL + 'A')
        time.sleep(1.5)
        qt_input.send_keys(qt_produzida)
        time.sleep(1.5)
        qt_input.send_keys(Keys.TAB)

def prencher_qtd_desviada(nav,qtd_desviada):
    
    iframes(nav)
    if qtd_desviada == None:
        qtd_desviada = ""

    qt_input_desviado = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DESVIADO"]')))
    qt_input_desviado.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    qt_input_desviado.send_keys(qtd_desviada)
    time.sleep(1.5)
    qt_input_desviado.send_keys(Keys.TAB)

    # verifica se a quantidade ta vazia:
    qt_input_desviado = WebDriverWait(nav, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="producoes"]//input[@name="DESVIADO"]'))
    ).get_attribute("value")

    if qt_input_desviado == '' and (qtd_desviada != "" or qtd_desviada == None):
        print('quantidade vazia')
        qt_input_desviado.send_keys(Keys.CONTROL + 'A')
        time.sleep(1.5)
        qt_input_desviado.send_keys(qtd_desviada)
        time.sleep(1.5)
        qt_input_desviado.send_keys(Keys.TAB)
    else:
        print('qt não vazia')

def prencher_dep_destino(nav):
    
    iframes(nav)

    dep_destino_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DEPOSITODESV"]')))
    dep_destino_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    dep_destino_input.send_keys('Almox Qualidade')
    time.sleep(1.5)
    dep_destino_input.send_keys(Keys.TAB)

def preencher_processo_corte(nav, row, erro):

    dados_espessura_chapa = leitura_google_planilhas_apoio_chapas()

    chapa = row['Código Chapa']
    if chapa == None:
        chapa = ""

    iframes(nav)

    recurso_movimento_deposito = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="0"]/td[9]/div/div'))).text
    try:
        chapa_atual = str(recurso_movimento_deposito).split('-')[0].strip() if '-' in str(recurso_movimento_deposito) else str(recurso_movimento_deposito)
    except:
        if erro:
            erro_consumo = f"Não possui dados para consumo, verifique se foi apontado"
        else:
            erro_consumo = None
        nav.switch_to.default_content()
        return erro_consumo
    peso_antigo_webdrive = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="0"]/td[26]/div/div')))
    peso_antigo = float(peso_antigo_webdrive.text)
    try:
        espessura_antiga = float(dados_espessura_chapa[dados_espessura_chapa['CODIGO'] == chapa_atual].ESPESSURA.values[0].replace(" mm","").replace(",","."))
    except:
        erro = f"A Chapa: {chapa_atual}, não foi encontrada na aba 'Apoio Chapa'"
        nav.switch_to.default_content()
        return erro
    
    print("Peso Antigo")
    print(peso_antigo)
    print("Espessura Antiga")
    print(espessura_antiga)

    if chapa_atual == chapa or chapa == "": 
        print('Chapa igual, não precisa fazer nada')
        nav.switch_to.default_content()
        return erro
    else:
        try:
            espessura_nova = float(dados_espessura_chapa[dados_espessura_chapa['CODIGO'] == chapa].ESPESSURA.values[0].replace(" mm","").replace(",","."))
        except:
            erro = f"A Chapa: {chapa}, não foi encontrada na aba 'Apoio Chapa'"
            nav.switch_to.default_content()
            return erro

        print("Espessura Nova")
        print(espessura_nova)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="0"]/td[9]/div/div'))).click()
        time.sleep(0.5)
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="0"]/td[9]/div/input'))).send_keys(chapa)
        time.sleep(0.5)

        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="0"]/td[9]/div/input'))).send_keys(Keys.TAB)
        time.sleep(0.5)

        carregamento(nav)
    
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="0"]/td[11]/div/input'))).send_keys(chapa_atual)
        time.sleep(0.5)

        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="0"]/td[11]/div/input'))).send_keys(Keys.TAB)
        time.sleep(0.5)

        carregamento(nav)

        peso_antigo_webdrive.click()

        quantidade_movimentacao_deposito = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[26]/div/input')))
        
        quantidade_movimentacao_deposito.send_keys(Keys.CONTROL + 'A')
        
        quantidade_movimentacao_deposito.send_keys((peso_antigo/espessura_antiga) * espessura_nova)

        time.sleep(2)

        quantidade_movimentacao_deposito.send_keys(Keys.TAB)
        
        carregamento(nav)

        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
        time.sleep(2)

        WebDriverWait(nav, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input'))).click()
        time.sleep(1)

        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
        time.sleep(1)

        carregamento(nav)

        erro = verificar_se_erro(nav)

        if erro and "O recurso substituído" in erro:
            iframes(nav)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[11]/div'))).click()
            time.sleep(2)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[11]/div/input'))).send_keys(Keys.CONTROL + 'A')
            time.sleep(1)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[11]/div/input'))).send_keys(Keys.BACKSPACE)
            time.sleep(1)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[11]/div/input'))).send_keys(Keys.TAB)

            carregamento(nav)

            # insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]'))).click()

            carregamento(nav)

            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            
            carregamento(nav)

            nav.switch_to.default_content()

            erro = verificar_se_erro(nav)
            
    return erro

def preencher_processo_pintura(nav, row):
    
    iframes(nav)
    table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="movDeposConsumidos"]/tbody/tr[1]/td[1]/table')))

    table_html_prod = table_prod.get_attribute('outerHTML')
        
    time.sleep(2)
    
    tabelona = None

    tabelona = pd.read_html(StringIO(table_html_prod), header=None)
    tabelona = tabelona[0].iloc[1:]

    # headers = tabelona.iloc[:1]
    # tabelona = tabelona.set_axis(headers.values.tolist()[0],axis='columns')[1:]

    tabelona = tabelona.rename(columns={11:'Recurso',30:'Quantidade'})

    tabelona = tabelona.dropna(subset='Recurso')
    
    tabelona = tabelona.reset_index(drop=True)
    
    tabelona['localizacao_tabela'] = range(3, 3 + 2 * len(tabelona), 2)

    print(tabelona)

    quantidade_po = None
    localizacao_po = None
    quantidade_catalisador = None
    localizacao_catalisador = None
    codigo_antigo = None
    quantidade_pu = None
    localizacao_pu = None
    cor_antiga = None
    tipo_tinta = row['Tipo']
    cor = row['Cor']

    df_cores = pd.read_excel("tintas.xlsx")
    print(df_cores)
    codigo = df_cores[(df_cores['COR_SIGLA'] == cor) & (df_cores['TIPO'] == tipo_tinta)]['CODIGO'].values[0]

    try:
        quantidade_catalisador = pd.to_numeric(tabelona[tabelona['Recurso'].str.contains('CATA')]['Quantidade']).values[0]
        localizacao_catalisador = tabelona[tabelona['Recurso'].str.contains('CATA')]['localizacao_tabela'].values[0]
    except:
        pass

    try:
        quantidade_pu = pd.to_numeric(tabelona[tabelona['Recurso'].str.contains('ESM. PU')]['Quantidade']).values[0]
        localizacao_pu = tabelona[tabelona['Recurso'].str.contains('ESM. PU')]['localizacao_tabela'].values[0]
        cor_antiga = tabelona[tabelona['Recurso'].str.contains('ESM. PU')]['Recurso'].values[0].split(' ')[0]
    except:
        pass

    try:
        quantidade_po = pd.to_numeric(tabelona[tabelona['Recurso'].str.contains('TINTA PÓ')]['Quantidade']).values[0]
        localizacao_po = tabelona[tabelona['Recurso'].str.contains('TINTA PÓ')]['localizacao_tabela'].values[0]
        codigo_antigo = tabelona[tabelona['Recurso'].str.contains('TINTA PÓ')]['Recurso'].values[0].split(' - ')[0]
        cor_antiga = tabelona[tabelona['Recurso'].str.contains('TINTA PÓ')]['Recurso'].values[0].split(' ')[0]
    except:
        pass
    qtd_linhas = len(tabelona)
    linha_maxima = tabelona['localizacao_tabela'].max()
    print(codigo_antigo)
    if tipo_tinta == 'PU':

        # verificando se contem catalisador
        if len(tabelona[tabelona['Recurso'].str.contains('CATA')]) == 1:
            iframes(nav)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input'))).click()
            carregamento(nav)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            carregamento(nav)
            erro = verificar_se_erro(nav)

        else:
            calculo_pu = quantidade_po * 1.58
            calculo_catalisador = calculo_pu / 6
            
            time.sleep(1)
            # clicar em insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
            
            time.sleep(1)
            # clicando em deposito
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys("Pintura")
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys(Keys.TAB)

            carregamento(nav)
            # inputando recurso Cor
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys(str(codigo))
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys(Keys.TAB)

            carregamento(nav)

            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[11]/div/input'))).send_keys(str(codigo_antigo))
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[11]/div/input'))).send_keys(Keys.TAB)
            
            time.sleep(2)

            # inputando quantidade
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(calculo_pu)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(Keys.TAB)

            time.sleep(2)
            # insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]'))).click()
                
            ######### CATALISADOR #########            
            
            carregamento(nav)
            # clicar em insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()

            carregamento(nav)
            print(f"linha_maxima {linha_maxima + 2}")
            print(f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/div')

            if qtd_linhas <= 8:
                linha_maxima += 2
                localizacao_po += 2
            
            print(f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/div')

            # clicando em deposito
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys("Pintura")
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys(Keys.TAB)

            time.sleep(.5)
            # inputando recurso Catalisador 
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys('313210')
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys(Keys.TAB)

            time.sleep(.5)
            # inputando quantidade
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(calculo_catalisador)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(Keys.TAB)

            time.sleep(.5)
            # insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]'))).click()

            carregamento(nav)

            print(localizacao_po - 2)

            # selecionando po
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{localizacao_po - 2}]/td[1]/input'))).click()

            time.sleep(1)
            # apagando po
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[3]'))).click()

            time.sleep(1)

            nav.switch_to.default_content()
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/table/tbody/tr/td[2]/div/div[2]'))).click()

            carregamento(nav)

            #confirmando tabela de cima
            iframes(nav)
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()

            carregamento(nav)

            # verifica se deu erro
            nav.switch_to.default_content()

            erro = verificar_se_erro(nav)

    elif tipo_tinta == 'PÓ':
        
        if len(tabelona[tabelona['Recurso'].str.contains('TINTA PÓ')]) == 1:
            iframes(nav)
            WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input'))).click()
            carregamento(nav)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
            carregamento(nav)
            erro = verificar_se_erro(nav)

        else:
            #calculando quantidade de po
            calculo_po = quantidade_pu / 1.58

            time.sleep(1)
            # clicar em insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div'))).click()
            
            time.sleep(1)
            # clicando em deposito
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys("Pintura")
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[7]/div/input'))).send_keys(Keys.TAB)

            carregamento(nav)
            # inputando recurso
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys(str(codigo))
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[9]/div/input'))).send_keys(Keys.TAB)

            carregamento(nav)
            # inputando recurso substituido (codigo antigo)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[11]/div/input'))).send_keys(cor_antiga)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[11]/div/input'))).send_keys(Keys.TAB)

            time.sleep(2)
            # inputando quantidade
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div'))).click()
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(calculo_po)
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima+2}]/td[26]/div/input'))).send_keys(Keys.TAB)

            time.sleep(2)

            # insert
            WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]'))).click()

            carregamento(nav)
            # selecionando catalisador e pu
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{localizacao_pu}]/td[1]/input'))).click()
            time.sleep(2)
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{localizacao_catalisador}]/td[1]/input'))).click()
            
            time.sleep(2)
            # apagando
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[3]'))).click()
            
            carregamento(nav)
            # confirmando
            nav.switch_to.default_content()
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/table/tbody/tr/td[2]/div/div[2]'))).click()

            carregamento(nav)
            #confirmando tabela de cima
            iframes(nav)
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()

            carregamento(nav)

            erro = verificar_se_erro(nav)
            
            if erro and "O recurso substituído" in erro:
                iframes(nav)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima - 2}]/td[11]/div'))).click()
                time.sleep(2)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima - 2}]/td[11]/div/input'))).send_keys(Keys.CONTROL + 'A')
                time.sleep(1)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima - 2}]/td[11]/div/input'))).send_keys(Keys.BACKSPACE)
                time.sleep(1)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[{linha_maxima - 2}]/td[11]/div/input'))).send_keys(Keys.TAB)

                carregamento(nav)

                # insert
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="movDeposConsumidos"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]'))).click()

                carregamento(nav)

                WebDriverWait(nav, 1).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]/div'))).click()
                
                carregamento(nav)

                nav.switch_to.default_content()

                erro = verificar_se_erro(nav)

    return erro

def recomecar(nav):

    nav.switch_to.default_content()
    iframes(nav)

    # Clicar em adicionar novo item
    try:
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="explorer"]//div[@id="insertButton"]'))).click()
        time.sleep(1.5)
    except:
        status = 'Erro ao clicar na lupa'

def clicar_ate_classe(nav, xpath, classe_esperada, max_tentativas=5, intervalo=2):
    tentativas = 0
    actions = ActionChains(nav)
    while tentativas < max_tentativas:
        try:
            # Localizar o botão
            btn = WebDriverWait(nav, 1).until(EC.presence_of_element_located((By.XPATH, xpath)))
            
            # Capturar a classe do botão
            classe_atual = btn.get_attribute('class')
            print(f"Tentativa {tentativas + 1}: Classe atual - {classe_atual}")
            
            # Verificar se a classe esperada apareceu
            if classe_esperada in classe_atual:
                print("Classe esperada detectada!")
                return True  # Sai do loop
            
            # Tentar clicar no botão
            actions.move_to_element(btn).click().perform()
            print("Botão clicado.")
            
            # Aguarda antes de tentar novamente
            time.sleep(intervalo)
            tentativas += 1
        except TimeoutException:
            print("Elemento não encontrado. Tentando novamente...")
            time.sleep(intervalo)
            tentativas += 1
    
    print("Falha ao encontrar a classe esperada.")
    return False

def leitura_google_planilhas(worksheet, key_sheets):

    scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
    client = gspread.authorize(credentials)
    # Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

    sh = client.open_by_key(key_sheets)
    wks = sh.worksheet(worksheet)

    df = wks.get()

    dados = pd.DataFrame(df)

    return dados, wks

def leitura_google_planilhas_apoio_chapas():

    scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
    client = gspread.authorize(credentials)
    # Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

    sh = client.open_by_key('1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE')
    wks = sh.worksheet('APOIO CHAPA')

    df = wks.get()

    base = pd.DataFrame(df)

    headers = wks.row_values(4)[11:]

    dados = base.iloc[4:, 11:13].set_axis(headers, axis=1)

    return dados

def buscando_dados(dados_planilha, indice=None):
    # Parâmetros: Nome da Aba e Chave da planilha
    dados, wks = leitura_google_planilhas(dados_planilha['nome_aba'], dados_planilha['chave_planilha'])

    nomes_colunas = dados_planilha['nome_das_colunas']

    dados.columns = dados.iloc[4]
    dados = dados[5:]  # Removendo a primeira linha agora que é o cabeçalho
    print(dados.columns)
    print(dados)
    dados.reset_index(drop=False, inplace=True)
    dados[nomes_colunas['status_pcp']] = dados[nomes_colunas['status_pcp']].astype(str)

    dados[nomes_colunas['data']] = pd.to_datetime(dados[nomes_colunas['data']], format='%d/%m/%Y', errors='coerce')

    # Obtendo mês e ano atuais
    mes_atual = datetime.datetime.now().month
    ano_atual = datetime.datetime.now().year

    # Filtrando apenas os dados do mês e ano atuais
    dados_filtrados = dados[
        (dados[nomes_colunas['data']].dt.month == mes_atual) & 
        (dados[nomes_colunas['data']].dt.year == ano_atual)
    ]

    dados_filtrados[nomes_colunas['data']] = dados_filtrados[nomes_colunas['data']].dt.strftime('%d/%m/%Y')
    
    dados_filtrados = dados_filtrados[(dados_filtrados[nomes_colunas['status_pcp']] == 'None') | (dados_filtrados[nomes_colunas['status_pcp']] == "")]

    # Aplicar filtro de quantidade de linhas se num_linhas não for None
    if indice is not None:
        dados_filtrados = dados_filtrados[dados_filtrados['index'] == indice]


    dados_filtrados[nomes_colunas['codigo']] = dados_filtrados[nomes_colunas['codigo']].apply(lambda x: "0" + str(x) if len(str(x)) == 5 else str(x))
    dados_filtrados[nomes_colunas['status_pcp']] = dados_filtrados[nomes_colunas['status_pcp']].apply(lambda x: "" if str(x) == 'None' else "")
    dados_filtrados[nomes_colunas['codigo']] = dados_filtrados[nomes_colunas['codigo']].apply(
        lambda x: str(x).split('-')[0].strip() if '-' in str(x) else str(x)
    )
    print(dados_filtrados)

    return dados_filtrados, wks

def registrar_status(codigo,status):

    url = "http://127.0.0.1:8000/registrar-status/"
    
    # Dados a serem enviados
    payload = {
        "codigo": codigo,
        "status": status
    }

    # Envia a requisição POST
    response = requests.post(url, json=payload)

    # Verifica a resposta
    if response.status_code == 200:
        print("Resposta:", response.json())  # Resposta JSON
    else:
        print(f"Erro {response.status_code}: {response.text}")

def checkbox_apontamentos(filename):

    scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
    client = gspread.authorize(credentials)
    # Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

    sh = client.open_by_key('1mkOhwzLCmptx0Qg-aMlX9P-o6DZQpsKzUeG4UCZl7PI')
    wks = sh.worksheet('PAINEL')


    lista_checkbox_setores = wks.get()
    lista_checkbox_setores = pd.DataFrame(lista_checkbox_setores)
    lista_checkbox_setores = lista_checkbox_setores.iloc[:,16:]
    lista_checkbox_setores = lista_checkbox_setores.iloc[6:12,0:2]
    lista_checkbox_setores = lista_checkbox_setores.set_axis(['Setor','Ativador'], axis=1)
    lista_checkbox_setores = lista_checkbox_setores[lista_checkbox_setores['Ativador'] == 'TRUE']

    return lista_checkbox_setores

def enviar_chave_para_planilha(wks, row, nav, coluna_chave):

    chave_do_apontamento = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')))
    
    # Pega o valor do input
    valor_chave_do_apontamento = chave_do_apontamento.get_attribute('value')

    wks.update_acell(coluna_chave + str(row['index'] + 1), valor_chave_do_apontamento)

    return valor_chave_do_apontamento