from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import time
import pandas as pd
import requests

from verificar_chrome import *

def login(nav):

    try:
        # logando
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="username"]'))).send_keys("luan araujo")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys("luanaraujo8")
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
        # Não há mensagem de carregamento inicial
        iframes(nav)
        pass 

    iframes(nav)

def verificar_se_erro(nav):

    nav.switch_to.default_content()

    time.sleep(3)

    error=None
    try:
        error = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="errorMessageBox"]')))
        confirm_button = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, "//button[@id='confirm']")))
        time.sleep(2)
        confirm_button.click()
        time.sleep(2)
    # WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
    #     By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
    except:
        return error

    return error

def clicar_em_add(nav):
    # Usar a função para clicar no botão "postButton"
    xpath_botao = '//*[@id="producoes"]//div[@id="insertButton"]'
    classe_esperada = "grid-titleBar-button grid-titleBar-newRecordButton-inactive"

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

def preencher_data_apontamento(nav):
    
    iframes(nav)

    data_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DATA"]')))
    data_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    data_input.send_keys('03/12/2024')
    time.sleep(1.5)
    data_input.send_keys(Keys.TAB)

def preencher_pessoa(nav):

    iframes(nav)

    pessoa_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PESSOA"]')))
    pessoa_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    pessoa_input.send_keys('4054')
    time.sleep(1.5)
    pessoa_input.send_keys(Keys.TAB)

def preencher_recurso(nav):

    iframes(nav)

    recurso_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="RECURSO"]')))
    recurso_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    recurso_input.send_keys('013213')
    time.sleep(1.5)
    recurso_input.send_keys(Keys.TAB)

def preencher_processo(nav):

    error = None

    iframes(nav)

    processo_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PROCESSO"]')))
    processo_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    processo_input.send_keys('S Estamparia')
    time.sleep(1.5)
    processo_input.send_keys(Keys.TAB)

    # Bloco de carregamento e erro 
    carregamento(nav)

    error=verificar_se_erro(nav)

    # ser der erro, refaz o input
    if error:
        print('erro')
        
        time.sleep(3)

        processo_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
            By.XPATH, '//*[@id="producoes"]//input[@name="PROCESSO"]')))
        processo_input.send_keys(Keys.CONTROL + 'A')
        time.sleep(1.5)
        processo_input.send_keys('S Estamparia')
        time.sleep(1.5)  

    return error         

def preecher_qtd_produzida(nav):
    
    iframes(nav)
    
    qt_input = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="PRODUZIDO"]')))
    qt_input.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    qt_input.send_keys('1')
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
        qt_input.send_keys('1')
        time.sleep(1.5)
        qt_input.send_keys(Keys.TAB)

def prencher_qtd_desviada(nav):
    
    iframes(nav)
    
    qt_input_desviado = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((
        By.XPATH, '//*[@id="producoes"]//input[@name="DESVIADO"]')))
    qt_input_desviado.send_keys(Keys.CONTROL + 'A')
    time.sleep(1.5)
    qt_input_desviado.send_keys('1')
    time.sleep(1.5)
    qt_input_desviado.send_keys(Keys.TAB)

    # verifica se a quantidade ta vazia:
    qt_input_desviado = WebDriverWait(nav, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="producoes"]//input[@name="DESVIADO"]'))
    ).get_attribute("value")

    if qt_input_desviado == '':
        print('quantidade vazia')
        qt_input_desviado.send_keys(Keys.CONTROL + 'A')
        time.sleep(1.5)
        qt_input_desviado.send_keys('1')
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
    dep_destino_input.send_keys('Almox Sucata')
    time.sleep(1.5)
    dep_destino_input.send_keys(Keys.TAB)

def preencher_cadastro(nav, dados):
    
    # buscar registro no google sheets
    try:
        actions = ActionChains(nav)
            
        iframes(nav)

        mudar_visao(nav)

        #Clica em "Inserir"
        iframes(nav)
        
        clicar_em_add(nav)

        carregamento(nav)

        # classe
        try:
            preencher_classe(nav)
            error=verificar_se_erro(nav)

            if error:
                preencher_classe(nav)
        except:
            return f'Erro ao inputar classe'

        carregamento(nav)

        # data
        try:
            preencher_data_apontamento(nav)
            error=verificar_se_erro(nav)

            if error:
                print('erro')
                preencher_data_apontamento(nav)
        except:
            return f'Erro ao inputar data'
        
        carregamento(nav)

        # pessoa
        try:
            preencher_pessoa(nav)
            error=verificar_se_erro(nav)

            if error:
                print('erro')
                preencher_pessoa(nav)
        except:
            return f'Erro ao inputar pessoa'

        carregamento(nav)

        #recurso
        try:
            error=preencher_recurso(nav)

            if error:
                print('erro')
                preencher_recurso(nav)
        except:
            return f'Erro ao inputar recurso'

        carregamento(nav)

        #processo
        try:
            preencher_processo(nav)
            error=verificar_se_erro(nav)

            if error:
                print('erro')
        except:
            return f'Erro ao inputar processo'

        carregamento(nav)

        #qt produzido
        try:
            preecher_qtd_produzida(nav)
        except:
            return f'Erro ao inputar qt produzida'

        carregamento(nav)

        # se tiver desvio
            # try:
            #     prencher_qtd_desviada(nav)
                # error=verificar_se_erro(nav)

                # if error:
                #     print('erro')
            # except:
            #     return f'Erro ao preencher desvio'
            
            # carregamento(nav)

            # try:
            #     prencher_dep_destino(nav)
                # error=verificar_se_erro(nav)

                # if error:
                #     print('erro')
            # except:
            #     return f'Erro ao preencher dep destino'

        #insert
        iframes(nav)
        # Usar a função para clicar no botão "postButton" (botão de confirmar)
        btn = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="producoes"]//div[@id="postButton"]')))
        actions.move_to_element(btn).click().perform()
        btn = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="producoes"]//div[@id="postButton"]')))
        actions.move_to_element(btn).click().perform()

        carregamento(nav)

        # verifica erro se der erro fecha a página e recomeça se não apenas aperta em add:
        erro = verificar_se_erro(nav)
        if erro:
            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
                By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
            time.sleep(1.5)
            clicando_producao_menu_aberto(nav)
        else:
            voltar_visao(nav)
    except:
        return 'Erro ao iniciar'

    return 'ok'

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

chrome_driver_path = verificar_chrome_driver()
nav = webdriver.Chrome(chrome_driver_path)
nav.maximize_window()
# nav.get("http://192.168.3.141/sistema")
nav.get("https://hcemag.innovaro.com.br/sistema/")

login(nav)

menu_apontamento(nav)

def buscando_dados():

    dados = pd.read_csv("dados.csv", sep=',')
    dados['codigo'] = dados['codigo'].apply(lambda x: "0" + str(x) if len(str(x)) == 5 else str(x) )

    return dados

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

dados  = buscando_dados()

status = preencher_cadastro(nav, dados)

