import time
from selenium import webdriver 
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

def inserirDados(data):
    book = load_workbook('learningPathRPA.xlsx')
    learningpathRPA = book['learningPathRPA']
    learningpathRPA.append(data)
    book.save('learningPathRPA.xlsx')

def pegarURLModulos(exame, link):
    listaLinks = []
    browser = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    browser.get(link)
    time.sleep(2.8)
    cardTemplate = browser.find_elements(By.CLASS_NAME, "card-template")
    for link in cardTemplate:
        url =  link.find_element(By.TAG_NAME, "a")
        url = url.get_attribute('href')
        print(url)
        listaLinks.append(url)
        
    browser.quit()
    bot(listaLinks, exame)


def bot(links, exame):
    browser = webdriver.Chrome(executable_path=r'./chromedriver.exe')
    for link in links:
        browser.get(link)
        time.sleep(2)
        nomeRoteiro = browser.find_element(By.TAG_NAME, "h1")
        try:
            tempoLearnPath = browser.find_element(By.CSS_SELECTOR, "#time-remaining")
            data = [exame, nomeRoteiro.text, link,tempoLearnPath.get_attribute('innerHTML')]
        except:
            try:   
                tempoLearnPath = browser.find_element(By.ID, "module-duration-minutes")
                data = [exame, nomeRoteiro.text, link,tempoLearnPath.get_attribute('innerHTML')]
            except:
                data = [exame, nomeRoteiro.text, link,'']    

        inserirDados(data)
        print(data)

    browser.quit()  

# pegarURLModulos('MB-800', 'https://learn.microsoft.com/pt-br/certifications/exams/mb-800')
# pegarURLModulos('MS-101', 'https://learn.microsoft.com/pt-br/certifications/exams/ms-101')
# pegarURLModulos('MS-220', 'https://learn.microsoft.com/pt-br/certifications/exams/ms-220')
# pegarURLModulos('MS-500', 'https://learn.microsoft.com/pt-br/certifications/exams/ms-500')
# # pegarURLModulos('MS-600', 'https://learn.microsoft.com/pt-br/certifications/exams/ms-600')
pegarURLModulos('MB-200', 'https://learn.microsoft.com/pt-br/certifications/exams/mb-400')
pegarURLModulos('MB-400', 'https://learn.microsoft.com/pt-br/certifications/exams/mb-400')
# pegarURLModulos('MD-101', 'https://learn.microsoft.com/pt-br/certifications/exams/md-101')








