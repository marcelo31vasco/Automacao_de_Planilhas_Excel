# Cria um chatboot que possibilite extrair informações de um arquivo do excel e
# Preencher automaticamente em qualquer sistema/site que esteja sendo utilizado.


from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl


# Caminho até o site
groa = webdriver.Chrome()
groa.get('https://contabilidade-devaprender.netlify.app/')
sleep(3)

# Validar Email
email_input = groa.find_element(By.XPATH, "//input[@id='email']")
sleep(2)
email_input.send_keys('groa@contabilidadegroa.com')

# Validar Senha
senha_input = groa.find_element(By.XPATH, "//input[@id='senha']")
sleep(2)
senha_input.send_keys('groa123')

# Clicar em Entrar
botao_entrar = groa.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(2)
botao_entrar.click()

sleep(5)

# Extraindo as informações da planilha

empresas = openpyxl.load_workbook('./empresas.xlsx')
pagina_empresas = empresas['dados empresas']

for linha in pagina_empresas.iter_rows(min_row=2, values_only=True):
    nome_empresa, email_empresa, telefone, endereco, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha

# Preenchendo as informações no sistema/site
    groa.find_element(By.ID,'nomeEmpresa').send_keys(nome_empresa)
    sleep(1)
    groa.find_element(By.ID,'emailEmpresa').send_keys(email_empresa)
    sleep(1)
    groa.find_element(By.ID,'telefoneEmpresa').send_keys(telefone)
    sleep(1)
    groa.find_element(By.ID,'enderecoEmpresa').send_keys(endereco)
    sleep(1)
    groa.find_element(By.ID,'cnpj').send_keys(cnpj)
    sleep(1)
    groa.find_element(By.ID,'areaAtuacao').send_keys(area_atuacao)
    sleep(1)
    groa.find_element(By.ID,'numeroFuncionarios').send_keys(quantidade_de_funcionarios)
    sleep(1)
    groa.find_element(By.ID,'dataFundacao').send_keys(data_fundacao)
    sleep(1)

    # Clicar em Cadastrar
    groa.find_element(By.ID,'Cadastrar').click()
    sleep(4)
