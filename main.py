import threading
import time
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

# link_global = 'https://compranet.hacienda.gob.mx/esop/guest/go/opportunity/detail?opportunityId='
#
# from tkinter import *
# fields = ('Código del expediente', 'Lista de codigos', 'Loan Principle', 'Monthly Payment', 'Remaining Loan')
#
# def makeform(root, fields):
#    entries = {}
#    for field in fields:
#       row = Frame(root)
#       lab = Label(row, width=22, text=field+": ", anchor='w')
#       ent = Entry(row)
#       ent.insert(0,"0")
#       row.pack(side = TOP, fill = X, padx = 5 , pady = 5)
#       lab.pack(side = LEFT)
#       ent.pack(side = RIGHT, expand = YES, fill = X)
#       entries[field] = ent
#    return entries



path_del_proyecto = os.path.realpath(__file__)
directorio_del_proyecto = os.path.dirname(path_del_proyecto)

def automatizar_descargas():

    # Rescatamos la columna y lista de links del excel
    book = open_workbook("seed.xlsx")
    sheet = book.sheet_by_index(0)  # If your data is on sheet 1
    columna_links = []
    for row in range(3, sheet.nrows):  # start from 1, to leave out row 0
        columna_links.append(" ".join(str(sheet.row_values(row)[45]).strip().split()))  # extract from zero col

    lista_personas_no_encontradas = manager.list()  # <-- can be shared between processes.
    lista_personas_encontradas = manager.list()  # <-- can be shared between processes.

    links = chunkIt(columna_links, 1)
    processes = []

    # Descargamos el ultimo binario de ChromeDriver
    get_driver()

    # Recorremos links
    for link in links:
        p = Process(target=descargar_archivos_persona,
                    args=(link, lista_personas_no_encontradas, lista_personas_encontradas))  # Passing the list
        p.start()
        processes.append(p)
    for p in processes:
        p.join()

    print("+++++++++++++++++ Descargas terminadas!!!!!")


def descargar_archivos_persona(links, lista_personas_no_encontradas, lista_personas_encontradas):
    driver = get_navigator()

    if os.name == 'nt':
        download_dir = directorio_del_proyecto + "\\expedientes"
    else:
        download_dir = directorio_del_proyecto + "/expedientes"

    print(links)
    # try:
    for link in links:
        while True:
            try:
                # Verificamos carga de la pagina, doy un sleep para evitar mensaje de muchas peticiones en la pagina
                driver.get(link)

                # Agarramos el
                nombre_expediente = driver.find_element(By.XPATH,
                                                        '/html/body/div/div[2]/div[4]/div[1]/div[3]/div/div/div[2]/div[1]')

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

                driver.quit()
                driver = get_navigator()

                print(e)
                continue
            break



    print("Expedientes descargados")
    # except:
    #     print("ups")
    #     time.sleep(30)


def chunkIt(seq, num):
    avg = len(seq) / float(num)
    out = []
    last = 0.0

    while last < len(seq):
        out.append(seq[int(last):int(last + avg)])
        last += avg

    return out


threadLocal = threading.local()

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

    driver = getattr(threadLocal, 'driver', None)
    if driver is None:
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
    freeze_support()
    manager = Manager()
    automatizar_descargas()

    # root = Tk()
    # ents = makeform(root, fields)
    # root.bind('<Return>', (lambda event, e = ents: fetch(e)))
    # b1 = Button(root, text = 'Final Balance',
    #   command=(lambda e = ents: final_balance(e)))
    # b1.pack(side = LEFT, padx = 5, pady = 5)
    # b2 = Button(root, text='Monthly Payment',
    # command=(lambda e = ents: monthly_payment(e)))
    # b2.pack(side = LEFT, padx = 5, pady = 5)
    # b3 = Button(root, text = 'Quit', command = root.quit)
    # b3.pack(side = LEFT, padx = 5, pady = 5)
    # root.mainloop()