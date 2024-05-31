from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from time import sleep
import pyautogui
import openpyxl


options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

class Automation:

    def __init__(self):
        self.servico = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(options=options, service = self.servico)

    def acessar_pagina_navegador(self):
        self.driver.get('https://contabilidade-devaprender.netlify.app/')
        sleep(5)
        pyautogui.hotkey("win", "left")
        sleep(2)

    def cadastrar_credenciais(self):
        #email = self.driver.find_element('xpath', '//*[@id="email"]') 
        # email = self.driver.find_element(By.XPATH, '//*[@id="email"]')
        email = self.driver.find_element(By.ID, 'email')
        sleep(2)
        email.send_keys('nicolas@icloud.com')

        senha = self.driver.find_element(By.ID, 'senha')
        sleep(2)
        senha.send_keys("123456789")
        
        botao = self.driver.find_element(By.ID, 'Entrar')
        sleep(2)
        botao.click()
        sleep(5)
        

    def ler_planilha(self):
        self.workbook = openpyxl.load_workbook('empresas.xlsx')
        self.sheet_empresas = self.workbook['empresas']
        
    def cadastrar_planilha_no_site(self):
        for linha in self.sheet_empresas.iter_rows(min_row=2, values_only=True):
            self.nome_da_empresa, self.email, self.telefone, self.endereco, self.cnpj, self.area_atuacao, self.qntd_funcionarios, self.data_fundacao = linha
            self.driver.find_element(By.ID, 'nomeEmpresa').send_keys(self.nome_da_empresa)
            sleep(1)
            self.driver.find_element(By.ID, 'emailEmpresa').send_keys(self.email)
            sleep(1)
            self.driver.find_element(By.ID, 'telefoneEmpresa').send_keys(self.telefone)
            sleep(1)
            self.driver.find_element(By.ID, 'enderecoEmpresa').send_keys(self.endereco)
            sleep(1)
            self.driver.find_element(By.ID, 'cnpj').send_keys(self.cnpj)
            sleep(1)
            self.driver.find_element(By.ID, 'areaAtuacao').send_keys(self.area_atuacao)
            sleep(1)
            self.driver.find_element(By.ID, 'numeroFuncionarios').send_keys(self.qntd_funcionarios)
            sleep(1)
            self.driver.find_element(By.ID, 'dataFundacao').send_keys(self.data_fundacao)
            sleep(1)
            self.driver.find_element(By.ID, 'Cadastrar').click()
            sleep(5)



    def iniciar(self):
        self.acessar_pagina_navegador()
        self.cadastrar_credenciais()
        self.ler_planilha()
        self.cadastrar_planilha_no_site()
        
automacao = Automation()
automacao.iniciar()