import io
import threading
import time
from tkinter import messagebox
from tkinter.ttk import Progressbar

import wget
import zipfile
import requests
import os

from multiprocessing import Manager
from multiprocessing.context import Process
from multiprocessing.spawn import freeze_support
from selenium import webdriver
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
from xlrd import open_workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

link_global = 'https://compranet.hacienda.gob.mx/esop/guest/go/opportunity/detail?opportunityId='

from tkinter import *

def makeform(root):
    row = Frame(root)

    root.title("AutoCompraNet")
    root.geometry('1000x350+500+200')


    label_titulo = Label(row, width=30, text="Lista de links de expedientes: ")
    label_titulo.grid(row = 1,column = 0, padx=5, pady=5)

    multitext = Text(row, height=15)
    multitext.grid(row = 1,column = 1, padx=0, pady=5)

    scrollbar = Scrollbar(root)
    scrollbar.pack(side=RIGHT, fill=Y)
    scrollbar.config(command=multitext.yview)
    multitext.config(yscrollcommand=scrollbar.set)

    row.pack(side=TOP, fill=X, padx=5, pady=5)

    b1 = Button(root, text='Descargar Expedientes',
                command=(lambda e=None: automatizar_descargas(multitext.get('1.0', END).splitlines(), root)))
    b1.pack(side=BOTTOM, padx=5, pady=5)


path_del_proyecto = os.path.realpath(__file__)
directorio_del_proyecto = os.path.dirname(path_del_proyecto)

def automatizar_descargas(columna_links = None, root = None):

    if len(columna_links)==0 or columna_links is None:
        # Rescatamos la columna y lista de links del excel
        book = open_workbook("seed.xlsx")
        sheet = book.sheet_by_index(0)  # If your data is on sheet 1
        columna_links = []
        for row in range(3, sheet.nrows):  # start from 1, to leave out row 0
            columna_links.append(" ".join(str(sheet.row_values(row)[45]).strip().split()))  # extract from zero col
    else:
        columna_links = columna_links[:-1]

    # Descargamos el ultimo binario de ChromeDriver
    get_driver()

    descargar_archivos_persona(columna_links, root)

def descargar_archivos_persona(links, root):
    driver = get_navigator()

    if os.name == 'nt':
        download_dir = directorio_del_proyecto + "\\expedientes"
    else:
        download_dir = directorio_del_proyecto + "/expedientes"

    for link in links:
        while True:
            try:
                # Verificamos carga de la pagina, doy un sleep para evitar mensaje de muchas peticiones en la pagina
                driver.get(link)

                # Agarramos el
                nombre_expediente = driver.find_element(By.XPATH,
                                                        '/html/body/div/div[2]/div[4]/div[1]/div[3]/div/div/div[2]/div[1]')

                if os.name == 'nt':
                    # Existe carpeta?
                    path = directorio_del_proyecto + "\\expedientes\\" + nombre_expediente.text

                    if os.path.isdir(path):
                        break

                    # Cambio direccion de descarga del chromedrive con el nombre de expediente
                    download_dir_temp = download_dir + "\\" + nombre_expediente.text
                else:
                    # Existe carpeta?
                    path = directorio_del_proyecto + "/expedientes/" + nombre_expediente.text

                    if os.path.isdir(path):
                        break

                    # Cambio direccion de descarga del chromedrive con el nombre de expediente
                    download_dir_temp = download_dir + "/" + nombre_expediente.text




                driver.command_executor._commands["send_command"] = (
                    "POST", '/session/$sessionId/chromium/send_command')
                params = {'cmd': 'Page.setDownloadBehavior',
                          'params': {'behavior': 'allow', 'downloadPath': download_dir_temp}}
                driver.execute("send_command", params)

                # Mientras haya un boton de siguiente en la ultima tabla...
                while True:

                    lista_archivos = []
                    try:
                        # Espero a que aparezca la lista de archivos
                        WebDriverWait(driver, 5).until(expected_conditions.visibility_of_element_located(
                            (By.XPATH, "/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[3]/tbody")))

                        # Consigo la lista de archivos
                        lista_archivos = driver.find_elements(By.XPATH,
                                                              '/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[3]/tbody/tr')
                    except:
                        # Consigo la lista de archivos
                        lista_archivos = driver.find_elements(By.XPATH,
                                                              '/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[2]/tbody/tr')



                    # Recorro la lista de archivos
                    for archivo in lista_archivos[1:]:
                        # Encuentro el link y descargo
                        link_archivo = archivo.find_element(By.TAG_NAME, 'a')
                        link_archivo.click()

                    # Verifico el boton de siguiente en la tabla
                    try:
                        boton_siguiente = driver.find_element(By.XPATH,
                                             '/html/body/div/div[2]/div[4]/div[1]/div[7]/form/div/div[4]/div/div[2]/span/span[2]/a')
                        boton_siguiente.click()
                    except:
                        break
            except Exception as e:

                time.sleep(5)

                driver.quit()
                driver = get_navigator()

                print(e)
                continue
            break

    time.sleep(5)

    driver.quit()
    messagebox.showinfo("Proceso Terminado", "Expedientes descargados")



def chunkIt(seq, num):
    avg = len(seq) / float(num)
    out = []
    last = 0.0

    while last < len(seq):
        out.append(seq[int(last):int(last + avg)])
        last += avg

    return out


def every_downloads_chrome(driver):
    if not driver.current_url.startswith("chrome://downloads"):
        driver.get("chrome://downloads/")
    return driver.execute_script("""
        var items = document.querySelector('downloads-manager')
            .shadowRoot.getElementById('downloadsList').items;
        if (items.every(e => e.state === "COMPLETE"))
            return items.map(e => e.fileUrl || e.file_url);
        """)


def get_driver():
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
    response = requests.get(url)
    version_number = response.text



    if os.name == 'nt':
        # build the donwload url
        download_url = "https://chromedriver.storage.googleapis.com/" + version_number + "/chromedriver_win32.zip"

        chromedriver = directorio_del_proyecto + "\\chromedriver.exe"
        download_dir = directorio_del_proyecto + "\\expedientes"
    else:
        # build the donwload url
        download_url = "https://chromedriver.storage.googleapis.com/" + version_number + "/chromedriver_linux64.zip"

        chromedriver = directorio_del_proyecto + "/chromedriver"
        download_dir = directorio_del_proyecto + "/expedientes"

    # download the zip file using the url built above
    latest_driver_zip = wget.download(download_url, 'chromedriver.zip')

    # extract the zip file
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall()  # you can specify the destination folder path here
    # delete the zip file downloaded above
    os.remove(latest_driver_zip)

def get_navigator():
    if os.name == 'nt':
        chromedriver = directorio_del_proyecto + "\\chromedriver.exe"
        download_dir = directorio_del_proyecto + "\\expedientes"
    else:
        chromedriver = directorio_del_proyecto + "/chromedriver"
        download_dir = directorio_del_proyecto + "/expedientes"

    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False,
        'profile.default_content_setting_values.automatic_downloads': 1
    })
    chrome_options.add_argument('--no-sandbox')
    # chrome_options.add_argument("--headless")

    if os.name != 'nt':
        os.chmod(chromedriver, 0o755)

    ser = Service(chromedriver)
    driver = webdriver.Chrome(service=ser, options=chrome_options)
    return driver


def fetch(e):
    pass


if __name__ == '__main__':
    root = Tk()
    makeform(root)
    root.mainloop()