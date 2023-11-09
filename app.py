from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import chromedriver_autoinstaller

import pandas as pd
import time
import gspread
import numpy as np

tabela_pecas = pd.read_csv("pecas.csv")
lista_pesos = pd.read_csv("lista_pesos.csv")

lista_pesos = lista_pesos[['Código','Nome','Procedência','Qtd.', 'Un.', 'Custo', 'nome_peca']]
tabela_pecas = tabela_pecas[['DEPOIS','status']]

#chromedriver_autoinstaller.install()

def iframes(nav):

    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass


try:

    #filename = r"C:\Users\pcp2\service_account.json"
    filename = "service_account.json"

    ############CHROME##################### --------------

    link1 = "http://192.168.3.141/"
    #link1 = "http://cemag.innovaro.com.br/sistema"
    nav = webdriver.Chrome('chromedriver.exe')
    nav.get(link1)

    #logando 
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="username"]'))).send_keys("ti.dev")
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys("cem@1616")

    time.sleep(2)

    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

    time.sleep(2)

    #abrindo menu
    #WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1898143037"]'))).click()    
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, 'menuBar-button-label'))).click()
    
    time.sleep(2)

    #clicando em producao
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="/Menu/Produção"]'))).click()

    time.sleep(1.5)

    #clicando em cons gerenciais
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="/Menu/Produção/Consultas gerenciais"]'))).click()

    time.sleep(1.5)

    #clicando em BOM
    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="-1897148439, Processos e composição de recursos com custos (BOM).il"]'))).click()

    iframe = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/iframe')))
    nav.switch_to.frame(iframe)

    #colocando data
    # data_input = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[4]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')))
    # data_input.click()
    # time.sleep(1)
    # data_input.clear()
    # time.sleep(1)
    # data_input.send_keys(data)    
    # data_input.send_keys(Keys.TAB)
    
    nav.switch_to.default_content()

    #inputando carreta
    #14:34
    #13 minutos = 32 carretas

    time.sleep(2)

    # carretas = carretas.reset_index()
    # ult_linha = carretas['index'].max()

    try:

        for item in range(len(tabela_pecas)):
            
            if tabela_pecas['status'][item] != 'ok' and tabela_pecas['status'][item] != 'erro':
                    
                nav.switch_to.default_content()
                
                #abrindo menu
                #WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1898143037"]'))).click()    
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, 'menuBar-button-label'))).click()

                #clicando em BOM
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="-1897148439, Processos e composição de recursos com custos (BOM).il"]'))).click()

                nav.switch_to.default_content()
                
                time.sleep(3)

                iframes(nav)
                
                carreta = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.NAME, 'recursos')))   
                
                #carreta = WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))   
                
                time.sleep(2)
                carreta.click()
                time.sleep(2)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.CONTROL + 'a')
                time.sleep(2)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.DELETE)
                time.sleep(2)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(tabela_pecas['DEPOIS'][item])
                time.sleep(3)
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input'))).send_keys(Keys.TAB)
                time.sleep(2)

                try:
                    time.sleep(2)
                    nav.switch_to.default_content()
                    WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div[2]/table/tbody/tr[2]/td/div/button'))).click()
                    tabela_pecas['status'][item] = 'erro'
                    continue
                except:
                    pass
                
                time.sleep(1.5)
                
                nav.switch_to.default_content()
                
                executar = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[p='Executar']")))

                executar.click()

                try:
                    while WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.ID, 'statusMessageBox'))):
                        print("carregando")
            
                except:
                    pass
                
                try:
                    while WebDriverWait(nav, 2).until(EC.element_to_be_clickable((By.ID, 'progressMessageBox'))):
                        print("carregando2")
            
                except:
                    pass
                
                time.sleep(1.5)
                
                nav.switch_to.default_content()
                
                iframes(nav)
                
                time.sleep(2)

                table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[1]')))
                table_html_prod = table_prod.get_attribute('outerHTML')
                codigo_peca = pd.read_html(str(table_html_prod), header=None)
                codigo_peca = codigo_peca[0]
                # codigo_peca = codigo_peca.dropna()
                # codigo_peca = codigo_peca.reset_index(drop=True)
                codigo_peca = codigo_peca.droplevel(level=0,axis=1)
                peca = codigo_peca['Recurso'][6]

                table_prod = WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/table[2]')))
                table_html_prod = table_prod.get_attribute('outerHTML')
                
                time.sleep(2)
                
                tabelona = pd.read_html(str(table_html_prod), header=None)
                tabelona = tabelona[0]
                tabelona = tabelona.dropna()
                
                tabelona = tabelona.reset_index(drop=True)
                
                tabelona = tabelona.droplevel(level=0,axis=1)
                tabelona['Qtd.'][0] = int(tabelona['Qtd.'][0])/100
                tabelona['Custo'][0] = int(tabelona['Custo'][0])/100
                tabelona['nome_peca'] = peca
                tabelona = tabelona.iloc[:1,:]
                
                lista_pesos = pd.concat([lista_pesos, tabelona], ignore_index=True)

                time.sleep(2)
                
                nav.switch_to.default_content()
                                    
                time.sleep(2)

                #Fechar aba
                WebDriverWait(nav, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]/span/div'))).click()

                time.sleep(3)
                
                tabela_pecas['status'][item] = 'ok'
            
            else:
                pass
    
        lista_pesos.to_csv("lista_pesos.csv")
        tabela_pecas.to_csv("pecas.csv")

    except:
        nav.quit()

except:
    pass