from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as condicaoEsperada
from time import sleep
from colorama import Fore, Style, init
import pyautogui as pg
import pandas as pd
import datetime
import os
import openpyxl
import PyPDF2
import requests
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename


class tj_bot:

    def __init__(self) -> None:
        init()
        print(Fore.GREEN + '==================================')
        print('*****  Desenvolvido por:     *****')
        print('*****    Geovanne Murata     *****')
        print('==================================\n')
        print(Fore.CYAN + '====> Whatsapp: +55 11 95284-2140\n\n')

        #print(Fore.RED + 'Por favor, coloque exatamente onde está o seu arquivo XLSX aqui + o nome do seu arquivo e aperte enter duas (2) vezes.')

        #diretorio = input('DIRETÓRIO + NOME DO ARQUIVO + EXTENSÃO, COMO POR EXEMPLO: D:\what-test\whatsapp.xlsx:\n' + Style.RESET_ALL)
        self.tela_incial = "https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fesaj%2Fj_spring_cas_security_check"
        sleep(2)

        self.usuario = '03332793870'
        self.senha = 'cigano24780109'

        Tk().withdraw()

        dirOrigem = askopenfilename(filetypes = (('xlsx files','*.xlsx'),))

        if not dirOrigem:
            print(Fore.RED + 'Nenhum arquivo selecionado' + Style.RESET_ALL)
            exit()

        self.df = pd.read_excel(dirOrigem)

        try:
            os.mkdir('./Resultado')
        except OSError:
            pass

        self.destino_excel = os.getcwd() + "\\Resultado\\"

    def start(self):
        self.carrega_pagina_web()
        self.login()
        self.mudar_para_cnpj()
        self.foro_e_cnpj()

    def carrega_pagina_web(self) -> None:
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument("--start-maximized")
        print( Fore.GREEN + 'Iniciando Browser\n' + Style.RESET_ALL)
        try:
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.wait = WebDriverWait(self.driver, 5)
            self.wait2 = WebDriverWait(self.driver, 120)
            self.driver.get(self.tela_incial)
        except:
            print(Fore.RED + 'Não foi possivel abrir a pagina web.' + Style.RESET_ALL)
            sleep(4)

    def login(self) -> None:
        lLogin: str = '//*[@id="usernameForm"]'
        lSenha: str = '//*[@id="passwordForm"]'
        lEntrar: str = '//*[@id="pbEntrar"]'
        lEsperar: str = '/html/body/header/nav/h1'
        lPesquisa: str = 'https://esaj.tjsp.jus.br/cpopg/open.do'

        try:
            seleciona_login = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lLogin)))
            seleciona_login.send_keys(self.usuario)
        except:
            print(Fore.RED + 'Erro no login' + Style.RESET_ALL)

        try:
            seleciona_senha = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lSenha)))
            seleciona_senha.send_keys(self.senha)
        except:
            print(Fore.RED + 'Erro no login' + Style.RESET_ALL)

        try:
            bt_entrar = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lEntrar)))
            bt_entrar.click()
        except:
            print(Fore.RED + 'Botão não encontrado' + Style.RESET_ALL)

        self.driver.get(lPesquisa)

        try:
            esperar_titulo = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lEsperar)))
        except:
            print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

    def mudar_para_cnpj(self):
        ltipo_cnpj: str = '//*[@id="cbPesquisa"]'

        try:
            selecao_doc = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, ltipo_cnpj)))
            selecao_object = Select(selecao_doc)
            selecao_object.select_by_value('DOCPARTE')
        except:
            pass
            print(Fore.RED + 'Não encontrado' + Style.RESET_ALL)

    def foro_e_cnpj(self):
        lCnpj: str = '//*[@id="campo_DOCPARTE"]'
        lForo: str = '//*[@id="select2-chosen-1"]'
        lPesquisaForo: str = '//*[@id="s2id_autogen1_search"]'
        lSelecionaForo: str = '/html/body/div[4]/ul/li[1]'
        lBotaoConsulta: str = '//*[@id="botaoConsultarProcessos"]'
        lVerMais: str = '/html/body/div[1]/div[3]/div/div[1]/a/span[1]'
        lVerAutos: str = '//*[@id="linkPasta"]'
        MudarJanela = self.driver.window_handles
        l_data_site = []
        l_data_pdf = []
        l_data_banco = []
        l_data = []

        for index, row in self.df.iterrows():
            foro_pesquisa = row[0]  # Repete o valor da primeira coluna para cada linha
            for col in row[1:]:
                proxima_lista = 1
                proximo_processo = 1
                cnpj:str = str (col)
                try:
                    selecao_cnpj = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lCnpj)))
                    selecao_cnpj.send_keys(cnpj)
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                try:
                    selecao_foro = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lForo)))
                    selecao_foro.click()
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                try:
                    pesquisa_foro = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lPesquisaForo)))
                    pesquisa_foro.send_keys(foro_pesquisa)
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                try:
                    selecao_lista = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lSelecionaForo)))
                    selecao_lista.click()
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                try:
                    selecao_botao = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lBotaoConsulta)))
                    selecao_botao.click()
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                sleep(2)

                while True:

                    resultado = self.verificar_processo(proxima_lista)

                    if resultado == False:
                        break

                    l_data_banco = self.pegar_nome_banco(proxima_lista)
                    lAcessarProcesso = f'/html/body/div[2]/div[2]/ul/li[{proximo_processo}]/div/div/div[1]/a'

                    # Acessa o processo
                    try:
                        acessar_processo = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, lAcessarProcesso)))
                        # acessar_processo.click()
                        processo = acessar_processo.get_attribute('href')
                        # sleep(1)
                    except:
                        print(Fore.RED + 'Processo ñ encontrado' + Style.RESET_ALL)
                    self.driver.get(processo)
                    # acessar_processo.click()

                    try:
                        ver_mais = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, lVerMais)))
                        ver_mais.click()
                    except:
                        print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                    l_data_site = self.pegar_info() # Pega info do processo

                    try:
                        ver_auto = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, lVerAutos)))
                        ver_auto.click()
                    except:
                        print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                    l_data_pdf =self.frame_pdf() # Pega info do PDF
                
                    data = {**l_data_banco, **l_data_site, **l_data_pdf} # Agrupa os dados

                    l_data.append(data) # Gera lista

                    self.gerar_arquivo(l_data)

                    proxima_lista += 1
                    proximo_processo += 1

                    self.driver.back()

                    sleep(1)

                    if proximo_processo >25:
                        break

                self.driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")

        self.gerar_arquivo(l_data)
        print(Fore.GREEN + 'Excel gerado na pasta resultado')

        sleep(1)

    def pegar_nome_banco(self, proxima_lista):
        lBanco: str = f'/html/body/div[2]/div[2]/ul/li[{proxima_lista}]/div/div/div[2]/div'

        try:
            nome_banco = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lBanco))).text
            print(nome_banco)
        except:
            print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

        data1 = {
            "Banco": nome_banco
        }

        return data1

    def verificar_processo(self, proxima_lista) -> bool:  # pelas 25 datas da lista de 25
        lData = f'/html/body/div[2]/div[2]/ul/li[{proxima_lista}]/div/div/div[4]/div'
        lEsperar: str = '//*[@id="listagemDeProcessos"]/h2'
        hoje = datetime.date.today()

        try:
            lista_datas = self.driver.find_elements(By.XPATH, lData)
        except:
            pass

        for l in lista_datas:

            try:
                data_processo = self.wait2.until(
                    condicaoEsperada.presence_of_element_located((By.XPATH, lData)))
                data_text = data_processo.text
                print(data_text[0:10])
            except:
                print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

        # data_text = '21/05/2023'
        data_site = datetime.datetime.strptime(data_text[0:10], '%d/%m/%Y').date() #mudar data_test por data_text
        if data_site == hoje:
            return True
        else:
            return False

    def pegar_info(self):
        lNumeroProcesso: str = '/html/body/div[1]/div[2]/div/div[1]/div/span[1]'
        lClasse: str = '//*[@id="classeProcesso"]'
        lAssunto: str = '//*[@id="assuntoProcesso"]'
        lForoInfo: str = '//*[@id="foroProcesso"]'
        lVaraCivel: str = '//*[@id="varaProcesso"]'
        lJuiz: str = '//*[@id="juizProcesso"]'
        lDistribuicao: str = '//*[@id="dataHoraDistribuicaoProcesso"]'
        lControle: str = '//*[@id="numeroControleProcesso"]'
        lArea: str = '//*[@id="areaProcesso"]/span'
        lValorAcao: str = '/html/body/div[1]/div[3]/div/div[2]/div/div[4]/div'

        pegar_classe: str = "Classe"
        pegar_assunto: str = "Assunto"
        pegar_foro: str = "Foro"
        pegar_vara: str = "Vara Cível"
        pegar_juiz: str = "Juiz"
        pegar_distrib: str = "Distribuição"
        pegar_controle: str = "Controle"
        pegar_area: str = "Área"
        pegar_valor: str = "Valor da Ação"

        try:
            pegar_processo = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lNumeroProcesso))).text
        except:
            print(Fore.RED + 'Classe não encontrada' + Style.RESET_ALL)

        try:
            pegar_classe = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lClasse))).text
        except:
            print(Fore.RED + 'Classe não encontrada' + Style.RESET_ALL)

        try:
            pegar_assunto = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lAssunto))).text
        except:
            print(Fore.RED + 'Assunto não encontrado' + Style.RESET_ALL)

        try:
            pegar_foro = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lForoInfo))).text
        except:
            print(Fore.RED + 'Foro não encontrado' + Style.RESET_ALL)

        try:
            pegar_vara = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lVaraCivel))).text
        except:
            print(Fore.RED + 'Vara não encontrada' + Style.RESET_ALL)

        try:
            pegar_juiz = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lJuiz))).text
        except:
            print(Fore.RED + 'Juiz não encontrada' + Style.RESET_ALL)

        try:
            pegar_distrib = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lDistribuicao))).text
        except:
            print(Fore.RED + 'Distribuição não encontrada' + Style.RESET_ALL)

        try:
            pegar_controle = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lControle))).text
        except:
            print(Fore.RED + 'Controle não encontrada' + Style.RESET_ALL)

        try:
            pegar_area = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lArea))).text
        except:
            print(Fore.RED + 'Área não encontrada' + Style.RESET_ALL)

        try:
            pegar_valor = self.wait.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lValorAcao))).text
        except:
            print(Fore.RED + 'Valor da ação não encontrada' + Style.RESET_ALL)

        data = {
            "Processo": pegar_processo,
            "Classe": pegar_classe,
            "Assunto": pegar_assunto,
            "Foro": pegar_foro,
            "Vara Cível": pegar_vara,
            "Juiz": pegar_juiz,
            "Distribuição": pegar_distrib,
            "Controle": pegar_controle,
            "Área": pegar_area,
            "Valor da Ação":pegar_valor
        }

        return data

    def frame_pdf(self):
        lPagUm: str = '//*[@id="pagina_1_cont_0_anchor"]'
        lPaginas: str = '//*[@id="viewer"]/div'
        lFrame: str = '//*[@id="documento"]'
        cpf_pattern = r'\d{3}\.\d{3}\.\d{3}-?\d{1,2}'
        cnpj_pattern = r'\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}'
        janela_inicial = self.driver.current_window_handle  # salvar janela atual
        janelas = self.driver.window_handles                # janela que abrir

        for janela in janelas:
            if janela not in janela_inicial:
                self.driver.switch_to.window(janela)

                try:
                    pag_um = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lPagUm)))
                    pag_um.click()
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL) # clica na pág 1

                try:
                    selecaoFrame = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lFrame)))
                    self.driver.switch_to.frame(selecaoFrame)
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                # elemento_texto = self.driver.find_element(By.TAG_NAME, 'body')
                # texto = elemento_texto.text

                try:
                    selecaoPaginas = self.wait2.until(
                        condicaoEsperada.presence_of_element_located((By.XPATH, lPaginas)))
                except:
                    print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                paginas = self.driver.find_elements(By.XPATH, lPaginas)
                pag_total = len(paginas) - 1

                dataCpf = {"CPF1": "", "CPF2": "", "CPF3": "", "CPF4": "", "CPF5": "", "CPF6": "", "CPF7": "" }
                dataCnpj = {"CNJP1": "", "CNJP2": "", "CNJP3": "", "CNJP4": "", "CNJP5": "", "CNJP6": "" }
                for x in range(pag_total):
                    contCpf: int = 1
                    contCnpj: int = 1
                    try:
                        selecaoPdf = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, '//*[@id="viewer"]/div[1]/div[2]/span')))
                    except:
                        print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)

                    spans = self.driver.find_elements(By.XPATH, f'//*[@id="viewer"]/div[{x + 1}]/div[2]/span')
                    # elemento_texto = self.driver.find_element(By.XPATH, lTextLayer)
                    # texto = elemento_texto.text

                    for span in spans:
                        text = span.text
                        if re.search(cpf_pattern, text):
                            cpf = re.findall(cpf_pattern, text)[0]
                            dataCpf[f"CPF{contCpf}"] = cpf

                            # dataCpf = {
                            #     f"CPF{contCpf}": cpf
                            # }
                            contCpf += 1

                        elif re.search(cnpj_pattern, text):
                            cnpj = re.findall(cnpj_pattern, text)[0]
                            dataCnpj[f"CNPJ{contCnpj}"] = cnpj

                            # dataCnpj = {
                            #     f"CNPJ{contCnpj}": cnpj
                            # }
                            contCnpj += 1
                
                l_data = {**dataCpf , **dataCnpj}

                self.driver.close()
                self.driver.switch_to.window(janela_inicial)

                return l_data

    def gerar_arquivo(self, data):
        df = pd.DataFrame(data, columns=["Banco", "Processo", "Classe", "Assunto", "Foro", "Vara Cível", "Juiz", "Distribuição", "Controle", "Área", "Valor da Ação", "CPF1", "CPF2", "CPF3", "CPF4", "CPF5", "CPF6", "CPF7", "CPF8", "CNPJ1", "CNPJ2", "CNPJ3", "CNPJ4", "CNPJ5", "CNPJ6"])
        
        df.to_excel(self.destino_excel + "Excel_Resultado.xlsx", index=False)
    
if __name__ == '__main__':
    executa = tj_bot()
    executa.start()

input('Escreva algo para fechar')