import datetime
from click import option
import pandas as pd
import os
import requests 
import pyexcel as p
import time
import win32com.client as client
import tabula
import io

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
from re import I
from datetime import timedelta, date

def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + timedelta(n)

def find_window(content: str):
            wids = driver.window_handles
            for window in wids:
                driver.switch_to.window(window)
                if content in driver.page_source.lower():
                    break

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

#Login
driver.get ("http://franquias.plenokw.com.br:6083/usuario/login")
driver.find_element(By.XPATH,('//*[@id="usuario"]')).send_keys("joao")
driver.find_element(By.XPATH,('//*[@id="senha"]')).send_keys("1234")
driver.find_element(By.XPATH,('//*[@id="tela-login"]/div/form/input[4]')).click()


#Gerar Relatório
driver.get ("http://franquias.plenokw.com.br:6083/relatorio-item-vendido-por-dia")

#### Inicialização
arquivo = open("setInit.txt", "a")
frases = list()
frases.append("2022\n")
frases.append("06\n")
frases.append("01\n")
arquivo.writelines(frases)

resultDf = pd.DataFrame(columns  = ["Cod", "Item", "Valor", "QTD", "Data"])
resultDf = resultDf.to_csv(r"C:\Users\Usuario\Desktop\Loja\relatorioAnual.csv")
###########################


arquivo = open("setInit.txt", "r")
y = int(arquivo.readline())
m = int(arquivo.readline())
d = int(arquivo.readline())
arquivo.close()

start_date = date(y, m, d)#data ultimo relatorio
end_date = datetime.date.today()

for single_date in daterange(start_date, end_date):
    #print(single_date.strftime("%d/%m/%Y"))
    driver.find_element(By.XPATH,('//*[@id="dataI"]')).clear()
    driver.find_element(By.XPATH,('//*[@id="dataI"]')).send_keys(single_date.strftime("%d/%m/%Y"))#automatizado já
    driver.find_element(By.XPATH,('//*[@id="gerar"]')).click() 

    time.sleep(0.5)
    find_window('CSV')

    nullDay = "registros para este dia."

    if not (nullDay in driver.page_source.lower()):

        
        driver.find_element(By.XPATH,('/html/body/div[2]/form/button[1]')).click()

        driver.close()
        find_window('Relatório Item Vendido Por Dia Consolidado')

        time.sleep(0.5)
        #move e renomeia o arquivo
        pdfNameFile = single_date.strftime("%d-%m-%Y") + ".pdf"
        #move e renomeia o arquivo
        file_oldname = os.path.join(r"C:\Users\Usuario\Downloads","relatorio.pdf")
        file_newname_newfile = os.path.join(r"C:\Users\Usuario\Desktop\Loja", pdfNameFile)

        os.rename(file_oldname, file_newname_newfile)

        #trasforma pdf e tranforma em csv
        csvNameFile = single_date.strftime("%d-%m-%Y") + ".csv"
        df = tabula.read_pdf(r"C:\Users\Usuario\Desktop\Loja\\" + pdfNameFile, pages='all')[0]
        tabula.convert_into(r"C:\Users\Usuario\Desktop\Loja\\" + pdfNameFile,r"C:\Users\Usuario\Desktop\Loja\\" + csvNameFile, output_format="csv", pages='all')

        time.sleep(2)
        #remover arquivo pdf
        os.remove(r"C:\Users\Usuario\Desktop\Loja\\" + pdfNameFile)

        #colocar coluna data adiciona datas e remove coluna Cod.F e substitui arquivo
        df = io.open(r"C:\Users\Usuario\Desktop\Loja\\" + csvNameFile, encoding='latin-1')
        df = pd.read_csv(df) #on_bad_lines='skip')#add skip 
        df.columns  = ["Cod", "Item", "Cod.F", "Valor", "QTD",]
        df.insert(5, "Data", single_date.strftime("%d/%m/%Y") , allow_duplicates=False)#inclui coluna data
        df = df.dropna()
        df = df.drop(columns=['Cod.F'])
        #print(df)
        os.remove(r"C:\Users\Usuario\Desktop\Loja\\" + csvNameFile)
        df.to_csv(r"C:\Users\Usuario\Desktop\Loja\\" + csvNameFile)

        
        #mesclando os csv
        resultDf = pd.read_csv(r"C:\Users\Usuario\Desktop\Loja\relatorioAnual.csv")
        resultDf = pd.concat([resultDf, df])
        resultDf = resultDf.drop(columns=['Unnamed: 0'])
        #print(resultDf)
        resultDf = resultDf.to_csv(r"C:\Users\Usuario\Desktop\Loja\relatorioAnual.csv")
        os.remove(r"C:\Users\Usuario\Desktop\Loja\\" + csvNameFile)
    
    else:
        driver.close()
        find_window('Relatório Item Vendido Por Dia Consolidado')

os.remove("setInit.txt")
'''arquivo = open("setInit.txt", "a")
frases = list()
frases.append(end_date.strftime("%Y\n"))
frases.append(end_date.strftime("%m\n"))
frases.append(end_date.strftime("%d\n"))
arquivo.writelines(frases)
arquivo.close()'''
