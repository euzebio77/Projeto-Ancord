from selenium import webdriver 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd 
import time 

servico = Service(ChromeDriverManager().install())
path = pd.read_excel("certificados.xlsx", dtype=str)

certificado_result = []

for i, cpf in enumerate(path["CPF"]):
    nome = path.loc[i, "Nome"]
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://credenciamento.ancord.org.br/login.html")
    time.sleep(2)
    navegador.find_element_by_xpath('//*[@id="btnVoltarInformativo"]').click()
    time.sleep(2) 
    navegador.find_element_by_xpath('//*[@id="lnkConsultaPublica"]').click()
    time.sleep(2) 
    navegador.find_element_by_xpath('//*[@id="menu"]/ul/li[2]/a').click()
    time.sleep(2) 
    navegador.find_element_by_xpath('//*[@id="menu"]/ul/li[2]/ul/li/a/span').click()
    time.sleep(2) 
    navegador.find_element_by_xpath('//*[@id="nomeCompleto"]').send_keys(cpf)
    time.sleep(2)
    navegador.find_element_by_xpath('//*[@id="cpf"]').send_keys(cpf)
    time.sleep(2) 
    navegador.find_element_by_xpath('//*[@id="btnConsultar"]').click()
    time.sleep(5) 
    navegador.find_element_by_xpath('/html/body/section/div[1]/div/section/div[2]/div/section/div/div[1]/div[2]/table/tbody/tr/td[2]').click()
    time.sleep(3)
    navegador.find_element_by_xpath('//*[@id="abaVinculo"]/a').click()
    time.sleep(3)
    contato = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]')
    for contatos in contato:
        if (contatos.text == "Vínculo não encontrado."):
            break

        else: 
            ide = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[1]')
            ins = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[2]')
            esvnc = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[3]')
            dtinc = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[4]')
            dtfim = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[5]')
            dtcria = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[6]')
            criador = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[7]')
            status = navegador.find_elements_by_xpath('//*[@id="datatable-vinculo"]/tbody/tr[1]/td[8]')
            nome = navegador.find_elements_by_xpath('//*[@id="datatable-consultar-solicitacao"]/tbody/tr/td[2]')
        
            for i in range(len(ide)):
                certificados_data = {'ID': ide[i].text,
                                     'Inst Financeira': ins[i].text,
                                     'Especie de Vinculo': esvnc[i].text,
                                     'Data Inicio de Vigencia': dtinc[i].text,
                                     'Data Fim de Vigencia': dtfim[i].text,
                                     'Data Criação': dtcria[i].text,
                                     'Criador': criador[i].text,
                                     'Status': status[i].text
                                    }       
                
        certificado_result.append(certificados_data)  
        df = pd.DataFrame(certificado_result)
        df = df.drop_duplicates(subset=['ID'])
        df.to_excel('vinculosv52.xlsx', index = False)
                
    time.sleep(1)
    navegador.quit()