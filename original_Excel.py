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
        self.senha = 'cigano2478'

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
            self.wait = WebDriverWait(self.driver, 10)
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
        lAcessarProcesso: str = '/html/body/div[2]/div[2]/ul/li[1]/div/div/div[1]/a'
        lVerMais: str = '/html/body/div[1]/div[3]/div/div[1]/a/span[1]'
        l_data = []
        for index, row in self.df.iterrows():
            foro_pesquisa = row[0]  # Repete o valor da primeira coluna para cada linha
            for col in row[1:]:
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

                resultado = self.verificar_processo()

                if resultado == True:
                    try:
                        acessar_processo = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, lAcessarProcesso)))
                        acessar_processo.click()
                        sleep(1)
                    except:
                        print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)   

                    try:
                        ver_mais = self.wait2.until(
                            condicaoEsperada.presence_of_element_located((By.XPATH, lVerMais)))
                        ver_mais.click()
                    except:
                        print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)
                    
                    l_data = self.pegar_info(l_data)

                else:
                    sleep(1)                   
                    pass

                self.driver.get("https://esaj.tjsp.jus.br/cpopg/open.do")

                if len(l_data) == 3: #Coloquei para teste, se não vai rodar todas linhas do excel até gerar o arquivo
                    self.gerar_arquivo(l_data)
                    print(Fore.GREEN + 'Excel gerado na pasta resultado')
                    exit()

                sleep(1)

            # self.gerar_arquivo(l_data) Usar esta linha de código para salvar todos os dados coletados.

    def verificar_processo(self) -> bool:
        lData: str = '/html/body/div[2]/div[2]/ul/li[1]/div/div/div[4]/div'
        lEsperar: str = '//*[@id="listagemDeProcessos"]/h2'
        hoje = datetime.date.today()

        try:
            data_processo = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lData)))
            data_text = data_processo.text
            print(data_text[0:10])
        except:
            print(Fore.RED + 'Página não encontrada' + Style.RESET_ALL)
        
        data_test = '16/05/2023'
        data_site = datetime.datetime.strptime(data_test[0:10], '%d/%m/%Y').date() #mudar data_test por data_text
        if data_site == hoje:
            return True
        else:
            return False

    def pegar_info(self, l_data):
        lClasse: str = '//*[@id="classeProcesso"]'
        lAssunto: str = '//*[@id="assuntoProcesso"]'
        lForoInfo: str = '//*[@id="foroProcesso"]'
        lVaraCivel: str = '//*[@id="varaProcesso"]'
        lJuiz: str = '//*[@id="juizProcesso"]'
        lDistribuicao: str = '//*[@id="dataHoraDistribuicaoProcesso"]'
        lControle: str = '//*[@id="numeroControleProcesso"]'
        lArea: str = '//*[@id="areaProcesso"]/span'
        lValorAcao: str = '//*[@id="valorAcaoProcesso"]'   
#        lCpfPalavra: str = '//*[contains(text(),"CPF")]' # acha o texto cpf
#        lCpfNumero: str = '' # achar o numero do cpf

        try:
            pegar_classe = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lClasse))).text
        except:
            print(Fore.RED + 'Classe não encontrada' + Style.RESET_ALL)

        try:
            pegar_assunto = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lAssunto))).text
        except:
            print(Fore.RED + 'Assunto não encontrado' + Style.RESET_ALL)

        try:
            pegar_foro = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lForoInfo))).text
        except:
            print(Fore.RED + 'Foro não encontrado' + Style.RESET_ALL)

        try:
            pegar_vara = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lVaraCivel))).text
        except:
            print(Fore.RED + 'Vara não encontrada' + Style.RESET_ALL)

        try:
            pegar_juiz = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lJuiz))).text
        except:
            print(Fore.RED + 'Juiz não encontrada' + Style.RESET_ALL)

        try:
            pegar_distrib = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lDistribuicao))).text
        except:
            print(Fore.RED + 'Distribuição não encontrada' + Style.RESET_ALL)

        try:
            pegar_controle = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lControle))).text
        except:
            print(Fore.RED + 'Controle não encontrada' + Style.RESET_ALL)

        try:
            pegar_area = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lArea))).text
        except:
            print(Fore.RED + 'Área não encontrada' + Style.RESET_ALL)

        try:
            pegar_valor = self.wait2.until(
                condicaoEsperada.presence_of_element_located((By.XPATH, lValorAcao))).text
        except:
            print(Fore.RED + 'Valor da ação não encontrada' + Style.RESET_ALL)

        data = {
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

        l_data.append(data)

        return l_data
       
    def gerar_arquivo(self, data):
        df = pd.DataFrame(data, columns=["Classe", "Assunto", "Foro", "Vara Cível", "Juiz", "Distribuição", "Controle", "Área", "Valor da Ação"])
        df.to_excel(self.destino_excel + "Excel_Resultado.xlsx", index=False)


# debug: marcar ponto de interrupção e f10 pra passar a linha por linha

# Trocar de navegador
#                 selecaoFrame = self.wait2.until(
#                    condicaoEsperada.presence_of_element_located((By.XPATH, lFrame)))
#                self.driver.switch_to.frame(selecaoFrame)  
#iniciar_app()

if __name__ == '__main__':
    executa = tj_bot()
    executa.start()

input('Escreva algo para fechar')