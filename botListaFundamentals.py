import time
from selenium import webdriver 
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

def inserirDados(data):
    book = load_workbook('learningPathFundamentalsRPA.xlsx')
    learningpathRPA = book['Sheet1']
    learningpathRPA.append(data)
    book.save('learningPathFundamentalsRPA.xlsx')

def pegarURLModulos(exame, link):
    listaLinks = []
    browser = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    browser.get(link)
    time.sleep(3)
    cardTemplate = browser.find_elements(By.CLASS_NAME, "card-template")
    for link in cardTemplate:
        url =  link.find_element(By.TAG_NAME, "a")
        url = url.get_attribute('href')
        print(url)
        listaLinks.append(url)
        
    # ALTERAR FUNCAO PARA CADA SUBTOPICO
    browser.quit()
    bot(listaLinks, exame)

def bot(links, prova):
    browser = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    for link in links:
        modulos = []
        browser.get(link)
        time.sleep(3)
        todasAsDivs = browser.find_elements(By.TAG_NAME, "div")
        for m in todasAsDivs:
            if(m.get_attribute("class") == "column is-auto padding-none padding-top-sm padding-sm-tablet"):
                modulos.append(m)

        for modulo in modulos:
            nomeModulo = modulo.find_element(By.TAG_NAME, "h3").text
            tempoNecessarioModulo = modulo.find_element(By.CLASS_NAME,'module-time-remaining').get_attribute('innerHTML')
            unidades = modulo.find_elements(By.CSS_SELECTOR, ".has-content-margin-right-xxs")
            inserirDados(['', '',nomeModulo, '', tempoNecessarioModulo, '','',prova])
            for unidade in unidades:
                nomeUnidade = unidade.find_element(By.TAG_NAME,"a").get_attribute('innerHTML')
                tempoNecessarioUnidade = unidade.find_element(By.TAG_NAME,"span").get_attribute('innerHTML')
                linkUnidade = unidade.find_element(By.TAG_NAME,"a").get_attribute('href')
                inserirDados(['','','',nomeUnidade, tempoNecessarioUnidade, '',linkUnidade,prova])


    browser.quit()  

pegarURLModulos('MB-910','https://learn.microsoft.com/pt-br/certifications/exams/mb-910')
pegarURLModulos('MB-920','https://learn.microsoft.com/pt-br/certifications/exams/mb-920')