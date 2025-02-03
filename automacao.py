from dict import dados_planilha
from verificar_chrome import verificar_chrome_driver
from funcoes.utils import *

def preencher_cadastro(nav, dados, wks, apontamento_atual,dados_planilha,df_ultimo_apontamento_que_gerou_um_erro):
    
    # buscar registro no google sheets
    try:

        actions = ActionChains(nav)
            
        iframes(nav)

        # Esse status serve para informar o status do ultimo apontamento, pois se tiver dado OK,
        # não tem necessidade de mudar de visão
        ultimo_status_apontamento = ""

        nomes_colunas = dados_planilha['nome_das_colunas']
        processo = dados_planilha['processo']

        coluna_ok = dados_planilha['coluna_de_ok']
        coluna_erro = dados_planilha['coluna_de_erro']
        coluna_chave = dados_planilha['coluna_da_chave']
    
        if not dados.empty:
            for _, row in dados.iterrows():
                
                iframes(nav)

                print("INDEX")
                print(row['index'])
                planilha, wks = buscando_dados(dados_planilha, row['index'])

                # Identificar se preencheram o OK manualmente nessa linha
                if planilha.empty:
                    segundos = 60
                    aguardando_requisicao_google_sheets(segundos)
                    continue
                
                if ultimo_status_apontamento != "OK":
                    mudar_visao(nav)

                ultimo_status_apontamento = ""

                try:
                    clicar_em_add(nav)
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    continue

                # classe
                print("chave")
                try:
                    valor_chave_do_apontamento = enviar_chave_para_planilha(wks, row, nav, coluna_chave)
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = "Erro ao receber a chave do apontamento"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro)
                    continue


                # classe
                print("classe")
                try:
                    preencher_classe(nav)
                    error=verificar_se_erro(nav)

                    if error:
                        preencher_classe(nav)
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher a classe"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue

                # data
                print("data")
                try:
                    preencher_data_apontamento(nav,row[nomes_colunas['data']])
                    error=verificar_se_erro(nav)

                    if error:
                        print('erro',error)
                        erro_na_data(nav,wks,row,error)
                        continue
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher a data do apontamento"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue
                

                # pessoa
                print("pessoa")
                try:
                    if nomes_colunas['funcionário'] in row:
                        preencher_pessoa(nav, apontamento_atual,row[nomes_colunas['funcionário']])
                    else:
                        preencher_pessoa(nav, apontamento_atual)

                    error=verificar_se_erro(nav)

                    if error:
                        print('erro')
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher o campo de pessoa"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue

                carregamento(nav)
                #recurso
                print("recurso")
                try:
                    preencher_recurso(nav,row[nomes_colunas['codigo']])
                    error=verificar_se_erro(nav)

                    if error:
                        exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                        continue
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher o recurso"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue


                #processo
                print("processo")
                try:
                    preencher_processo(nav,processo)
                    
                    error=verificar_se_erro(nav)

                    print('error')
                    print(error)

                    if error:
                        exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                        continue
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher o processo"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue


                #qt produzido
                print("qt produzido")
                try:
                    preecher_qtd_produzida(nav, row[nomes_colunas['quantidade']])
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher a quantidade produzida"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue

                carregamento(nav)

                # se tiver desvio
                print('se tiver desvio')
                try:
                    if nomes_colunas['mortas'] in row:
                        prencher_qtd_desviada(nav,row[nomes_colunas['mortas']])
                        error=verificar_se_erro(nav)

                        if error:
                            print('erro')
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher a quantidade desviada"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue
                
                carregamento(nav)

                # dep_destino
                print('dep_desv')
                try:
                    prencher_dep_destino(nav)
                    error=verificar_se_erro(nav)

                    if error:
                        print('erro')
                except:
                    sair_da_aba_e_voltar_ao_menu(nav)
                    error = f"Erro ao preencher o dep destino"
                    exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                    continue

                #insert
                iframes(nav)
                # Usar a função para clicar no botão "postButton" (botão de confirmar)
                btn = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="producoes"]//div[@id="postButton"]')))
                actions.move_to_element(btn).click().perform()
                btn = WebDriverWait(nav, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="producoes"]//div[@id="postButton"]')))
                actions.move_to_element(btn).click().perform()

                print("postButton")

                carregamento(nav)

                erro = verificar_se_erro(nav)

                if apontamento_atual == 'corte':
                    try:
                        erro = preencher_processo_corte(nav, row, erro)
                    except:
                        sair_da_aba_e_voltar_ao_menu(nav)
                        error = f"Erro ao preencher os dados de corte (Mov. de depósito)"
                        exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                        continue

                elif apontamento_atual == 'pintura':
                    # Ao terminar o apontamento de pintura, resetar o campo do ultimo erro

                    try:
                        erro = preencher_processo_pintura(nav, row)
                    except:
                        sair_da_aba_e_voltar_ao_menu(nav)
                        error = f"Erro ao preencher os dados de pintura (Mov. de depósito)"
                        exibindo_erro_na_planilha(nav,wks,row,error,coluna_erro,valor_chave_do_apontamento)
                        continue
                
                if erro:
                    print(erro)
                    exibindo_erro_na_planilha(nav,wks,row,erro,coluna_erro, valor_chave_do_apontamento)
                    carregamento(nav)
                else:
                    iframes(nav)
                    clicar_em_add(nav)
                    carregamento(nav)
                    print('OK')
                    ultimo_status_apontamento = 'OK'
                    wks.update_acell(coluna_ok + str(row['index'] + 1), f'OK {data_hoje()} {hora_atual()}')
        else:
            return "Sem dados para apontar"
    except:
        return f"Erro inesperado ao apontar o {apontamento_atual}, reiniciando processo de apontamento"
    
    sair_da_aba_e_voltar_ao_menu(nav)
    return "Sem dados para apontar"

status_execucao = ""

while True:
    if status_execucao != "Sem dados para apontar":
        try:
            nav = webdriver.Chrome()
        except:
            chrome_driver_path = verificar_chrome_driver()
            nav = webdriver.Chrome(chrome_driver_path)

        nav.maximize_window()
        nav.get("http://192.168.3.141/sistema")
        # nav.get("https://hcemag.innovaro.com.br/sistema/")

        login(nav)
        
        menu_apontamento(nav)
    
    try:

        for apontamento_atual, value in dados_planilha.items():
            
            # apontamento_atual == nome do setor
            # value == os dados do setor
            # visualizar esses dados no dict.py

            df_ultimo_apontamento_que_gerou_um_erro = pd.read_csv('Ultimo erro no apontamento - Página1.csv', sep=";")

            ultimo_erro = df_ultimo_apontamento_que_gerou_um_erro.loc[0].values[0]

            # Evitar se manter no mesmo erro, porém quando chegar em pintura, esse campo será resetado 
            if ultimo_erro == apontamento_atual:
                continue

            dados, wks = buscando_dados(value)

            status_execucao = preencher_cadastro(nav, dados, wks, apontamento_atual,value,df_ultimo_apontamento_que_gerou_um_erro)

            if status_execucao == "Sem dados para apontar":
                print("Sem dados para apontar")
                segundos = 60
                aguardando_requisicao_google_sheets(segundos)
                continue
    except:
        continue