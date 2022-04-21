import shutil
import time
import xml.etree.ElementTree as ET
import xlrd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import clipboard
import sys,os


#urls
smartFactorURL = "https://www.smartfactor.com.br/smart/"
NFSURL = "https://registro.ginfes.com.br/"

#credentials
loginSm = ""
pwdSm = ""
loginNFS = ""
pwdNFS = ""

globalVar = {
    'resultados': '',
    'tries': 0
}


class titulosResult:
    def __init__(self, values, data=None, NFS=None, sacado=None, titulo=None, valor=None):
        values = values.split("%")
        self._cedente = values[0]
        self._op = values[1]
        self._data = data
        self._NFS = NFS
        self._sacado = sacado
        self._titulo = titulo
        self._valor = valor

    @property
    def cedente(self):
        return self._cedente
    
    @property
    def op(self):
        return self._op

    @property
    def data(self):
        return self._data

    @data.setter
    def data(self, data):
        self._data = data

    @property
    def NFS(self):
        return self._NFS

    @NFS.setter
    def NFS(self, NFS):
        self._NFS = NFS

    @property
    def sacado(self):
        return self._sacado

    @sacado.setter
    def sacado(self, sacado):
        self._sacado = sacado

    @property
    def titulo(self):
        return self._titulo

    @titulo.setter
    def titulo(self, titulo):
        self._titulo = titulo

    @property
    def valor(self):
        return self._valor

    @valor.setter
    def valor(self, valor):
        self._valor = valor

    def getValues(self):
        return {
            'cedente': self._cedente,
            'op': self._op,
            'data': self._data,
            'NFS': self._NFS,
            'sacado': self._sacado,
            'titulo': self._titulo,
            'valor': self._valor
        }

def download_wait(path):
    while len(os.listdir(path)) == 0:
        print("aguardando download...")
        time.sleep(1)
    files = os.listdir(path)
    finished = False
    while not finished:
        files = os.listdir(path)
        print(files)
        if(".xml" in files[0]):
            finished = True
        time.sleep(1)
    
    return files[0]

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()
#nfes download dir
def create_dir():
    filename = sys.argv[0]
    path = os.path.abspath(filename)
    path = path.rstrip("scrapper.py")
    path = path+"cache\\"
    try:
        if os.path.exists(path):
            shutil.rmtree(path, ignore_errors=False)
        os.makedirs(path)
    except:
        create_dir()
    else:
        return path
    return "Algo deu errado na criação do diretório"
#get lista de titulos
def lista_titulos(nomearquivotabelafdc):
    arquivo = xlrd.open_workbook(nomearquivotabelafdc)

    tabela = arquivo.sheet_by_index(0)


    listTitulos = []
    for x in range(11, tabela.nrows-1): 
        listTitulos.append(tabela.cell_value(rowx=x, colx=1)+"%"+tabela.cell_value(rowx=x, colx=2)+"%"+tabela.cell_value(rowx=x,colx=3)+"%"+tabela.cell_value(rowx=x, colx=6))

    return listTitulos

def encontrarSacado(driver, data):
    driver.switch_to.parent_frame()
    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/body/div[1]/iframe[2]"))

    rows = 1+len(driver.find_elements(By.XPATH, "/html/body/form/table[1]/tr"))

    #encontrar valores
    tValue = []
    for r in range(1, rows):
        titulo = driver.find_element(By.XPATH, "/html/body/form/table[1]/tr["+str(r)+"]/td[6]/input").get_attribute('value')
        sacado = driver.find_element(By.XPATH, "/html/body/form/table[1]/tr["+str(r)+"]/td[9]").text
        tValue.append(titulo+"%"+sacado)

    for x in tValue:
        if(similar(x, data) >= 0.62):
            return True
    return False

def setOpts():
    opt = webdriver.ChromeOptions()
    opt.add_argument("--disable-infobars")
    opt.add_argument("start-maximized")
    opt.add_argument("--disable-extensions")
    opt.add_argument("--disable-popup-blocking")
    prefs = {
        'download.default_directory': path,
        'download.prompt_for_download': 'false',
        'download.extensions_to_open': 'xml',
        'safebrowsing.enabled': 'false'
    }
    opt.add_experimental_option('prefs', prefs)
    # opt.add_argument("--headless")
    return opt

def main(listaTitulos):
    print(globalVar['tries'])
    if(globalVar['tries'] != 3):
        with webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=setOpts()) as driver:
        
            try:
                wait = WebDriverWait(driver, 15)

                #log in to ginfes
                print("Conectando ao site GINFES")

                driver.get(NFSURL)
                # time.sleep(2)
                btnBoneco = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[alt="Acesso Exclusivo Prestador"]')))
                btnBoneco.click()

                radioBtn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[value="on"]'))) #ec works this way
                radioBtn.click()
                print("Fazendo log-in no sistema")
                inputInscMun = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/div/form/fieldset[2]/div/div/div[3]/div[1]/input')))
                inputInscMun.send_keys(loginNFS)

                driver.find_element(By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/div/form/fieldset[2]/div/div/div[4]/div[1]/input').send_keys(pwdNFS)
                driver.find_element(By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/div/form/table/tbody/tr/td[1]/table/tbody/tr/td[2]/em/button').click()

                btnConsultar = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/div/img')))
                btnConsultar.click()

                opcConsulta = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/fieldset/div/div/div/div/div/div/div[2]/div/div/span/input')))
                opcConsulta.click()

                selectDate = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/fieldset/div/div/div/div/div/form/div[2]/div/div/div/div[1]/div/div/div/div[1]/div[1]/input')))
                selectDate.clear()
                selectDate.send_keys("01/12/2021")
                #botao para consultar
                print("Aguardando resultado da consulta.")
                driver.find_element(By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[9]/td/table/tbody/tr/td/table/tbody/tr/td[2]/em/button').click()
                #baixar resposta
                btnBaixar = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[6]/td/fieldset/div/div/div/div/div/div/div[1]/div/div/img')))
                btnBaixar.click()

                downloadedFilePath = path+download_wait(path)
                print("Arquivo baixado: "+downloadedFilePath)

                arquivo = downloadedFilePath
                
                #organiza os dados do arquivo
                tree = ET.parse(arquivo)
                root = tree.getroot()

                nfsList = {}
                tags = ['Numero', 'DataEmissao', 'ValorServicos', 'Discriminacao', 'RazaoSocial']
                for x in range(0, len(root.findall("{http://www.w3.org/2000/09/xmldsig#}Nfse"))):
                    temp = {}
                    op = None
                    for y in root[x].iter():
                        if(y.text and "L & P" not in y.text):
                            tag = y.tag
                            index = tag.index("}")
                            tag = tag[index+1:]
                            if(tag in tags and not tag in temp):
                                texto = y.text
                                if(tag == "Discriminacao"):
                                    texto = texto.split("R$")[0]
                                    texto = "".join([i for i in texto if i in "0123456789"])
                                    op = texto
                                temp[tag] = texto

                    nfsList[op] = temp #nfsList contém todos os dados ncessários do arquivo baixado para consulta

                
                #log in to smartFactor
                driver.get(smartFactorURL)
                #pesquisa de titulo
                titleSearchURL = "https://www.smartfactor.com.br/smart/financeiro/frmfinanceiro.php?page=financeiro/pesquisa.php"

                driver.find_element(By.ID, 'fEmail').send_keys(loginSm)
                driver.find_element(By.ID, 'fPassword').send_keys(pwdSm)
                driver.find_element(By.ID, 'OK').click()
                #ENTRA NA PAGINA DE PESQUISA POR TITULO TODO: dynamic wait
                time.sleep(6)
                listaFinalSmart = []
                for tituloSacado in listaTitulos:
                    tituloSacadoSplit = tituloSacado.split("%") 
                    titulo = tituloSacadoSplit[0]
                    sacado = tituloSacadoSplit[1]
                    vcto = tituloSacadoSplit[2]
                    valor = tituloSacadoSplit[3]

                    driver.get(titleSearchURL) #works
                    driver.switch_to.frame(driver.find_element(By.ID, 'code'))
                    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/frameset/frame"))
                    inNumTitle = driver.find_element(By.XPATH, '/html/body/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input')
                    #DIGITAR NUMERO DO TITULO
                    inNumTitle.send_keys(titulo)
                    #PESQUISA
                    driver.find_element(By.NAME, 'Submit').click()
                    #pegar tabela
                    driver.switch_to.parent_frame()
                    driver.switch_to.frame(driver.find_element(By.XPATH, "/html/frameset/frame"))
                    #localização da tabela
                    rows = len(driver.find_elements(By.XPATH, "/html/body/form[1]/table/tbody/tr[2]/td[2]/table[1]/tbody/tr"))

                    #encontrar valores
                    tValue = ""
                    for r in range(2, rows):
                        cedentePesquisa = driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr[2]/td[2]/table[1]/tbody//tr["+str(r)+"]/td[20]").text
                        opPesquisa = driver.find_element(By.XPATH, "/html/body/form[1]/table/tbody/tr[2]/td[2]/table[1]/tbody//tr["+str(r)+"]/td[23]").text
                        if (cedentePesquisa and opPesquisa):
                            tValue = tValue + "[" + cedentePesquisa+"%"+opPesquisa
                    
                    if len(tValue) <= 1:
                        raise Exception("Nenhuma operação encontrada para: ", titulo)

                    #separa resultados encontrados
                    tValueSplit = tValue.split('[')
                    tValueSplit.pop(0)
                    tValueResult = []
                    #lista de resultados da pesquisa
                    for i in range(len(tValueSplit)):
                        tValueResult.append(titulosResult(tValueSplit[i]))

                    #busca operação e salva valores
                    tentativa = 0
                    success = False
                    while tentativa != len(tValueResult):
                        driver.get("https://www.smartfactor.com.br/smart/operacaoajax/web/frmoperacao.php?action=edit&op="+tValueResult[tentativa].op+"&tipo=O")
                        driver.switch_to.frame(driver.find_element(By.ID, 'code'))
                        driver.switch_to.frame(driver.find_element(By.ID, 'framesetTitulo'))
                        cedente = driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[1]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/input[1]").get_attribute('value')

                        if(tValueResult[tentativa].cedente == cedente and encontrarSacado(driver, titulo+"%"+sacado)):
                            tValueResult[tentativa].data = vcto
                            tValueResult[tentativa].titulo = titulo
                            tValueResult[tentativa].sacado = sacado
                            tValueResult[tentativa].NFS = nfsList[tValueResult[tentativa].op]['Numero']
                            tValueResult[tentativa].valor = valor
                            success = True
                            listaFinalSmart.append(tValueResult[tentativa])
                            break
                        tentativa += 1
                        pass
                    
                    if not success: 
                        raise Exception('Não foi encontrado nenhuma operação correspondente a: ', titulo)
                    

                result = "" 
                for x in range(0, len(listaFinalSmart)):
                    values = listaFinalSmart[x].getValues()
                    result += "\n REF - NF"+values['NFS']+", SM"+values['op']+", VALOR R$"+values['valor']+", VCTO "+values['data']+", Cedente: "+values['cedente']
                
                return result
            except:
                driver.close()
                print("Falha na conexão, tentando novamente...")
                globalVar['tries'] = globalVar['tries']+1
                main(listaTitulos)
    else:
        string = "Erros sucessivos resultaram no término da aplicação"
        return string
   
def open_text_file():
    # file type
    filetypes = (
        ('Excel files', '.xlsx .xls'),
    )   
    # show the open file dialog
    f = fd.askopenfile(filetypes=filetypes)
    # read the text file and show its content on the Text
    listaTitulos = lista_titulos(f.name)
    globalVar['resultados'] = main(listaTitulos)
    text.configure(state='normal')
    texto = globalVar['resultados']
    text.insert('1.0', texto)
    text.configure(state='disabled')
    copy_clipboard.configure(state='normal')


def copy_to_clipboard():
    root.update()
    clipboard.copy(globalVar['resultados'])
    copy_clipboard.configure(text='Copiado!')
    root.update()




path = create_dir()

# Root window
root = tk.Tk()
root.title('AutomatronInator 3.000')
root.resizable(False, False)
root.geometry('600x300')

# Text editor
text = tk.Text(root, height=12)
text.grid(column=0, row=0, sticky='nsew')
text.configure(state='disabled')
text.bind("<1>", lambda event: text.focus_set())



# open file button
open_button = ttk.Button(
    root,
    text='Selecionar Arquivo',
    command=open_text_file
)

#copiar texto
copy_clipboard = ttk.Button(
    root,
    text='Copiar resultado',
    command=copy_to_clipboard
)

open_button.grid(column=0, row=1, sticky='w', padx=10, pady=10)
copy_clipboard.grid(column=0, row=2, sticky='w', padx=10, pady=10)
copy_clipboard.configure(state='disabled')

root.after(1000, copy_clipboard.configure(text='Copiar resultado'))
root.mainloop()
