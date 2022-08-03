import time
import wget
import zipfile
import requests
import os
import pandas as pd
import openpyxl

from tkinter import messagebox
from tkinter import *
from selenium import webdriver
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

ACCESS_LINK = 'https://compranet.hacienda.gob.mx/esop/guest/go/opportunity/detail?opportunityId='
APP_PATH = os.path.dirname(os.path.realpath(__file__))


def make_main_view(root):
    row = Frame(root)

    root.title("AutoCompraNet")
    root.geometry('1000x350+500+200')

    lb_title = Label(row, width=30, text="Lista de links de expedientes: ")
    lb_title.grid(row=1, column=0, padx=5, pady=5)

    mt_ids = Text(row, height=15)
    mt_ids.grid(row=1, column=1, padx=0, pady=5)

    scrollbar = Scrollbar(root)
    scrollbar.pack(side=RIGHT, fill=Y)
    scrollbar.config(command=mt_ids.yview)
    mt_ids.config(yscrollcommand=scrollbar.set)

    row.pack(side=TOP, fill=X, padx=5, pady=5)

    btn_start_workflow = Button(root, text='Descargar Expedientes',
                                command=(lambda e=None: main_workflow(mt_ids.get('1.0', END).splitlines())))
    btn_start_workflow.pack(side=BOTTOM, padx=5, pady=5)


def main_workflow(ids_list):
    if len(ids_list) == 1:
        # get seed.xlsx with pandas and get the 45th column, skip first two rows and clean each strings
        df = pd.read_excel(APP_PATH + "/seed.xlsx", skiprows=2)
        ids_list = df.iloc[:, 45].str.strip().tolist()
    else:
        ids_list = ids_list[:-1]

    get_latest_driver()
    download_workflow(ids_list)


def download_workflow(links):
    driver = use_driver()
    os_string = ""
    os_separator = ""

    if os.name == 'nt':
        download_dir = APP_PATH + "\\expedientes"
        os_string = "\\expedientes\\"
        os_separator = "\\"
    else:
        download_dir = APP_PATH + "/expedientes"
        os_string = "/expedientes/"
        os_separator = "/"

    for link in links:
        while True:
            try:
                # Verificamos carga de la pagina, doy un sleep para evitar mensaje de muchas peticiones en la pagina
                driver.get(link)

                # Agarramos el
                name_folder = driver.find_element(By.XPATH,
                                                  '/html/body/div/div[2]/div[4]/div[1]/div[3]/div/div/div[2]/div[1]')

                # Existe carpeta?
                path = APP_PATH + os_string + name_folder.text

                if os.path.isdir(path):
                    break

                # Cambio direccion de descarga del chromedrive con el nombre de expediente
                download_dir_temp = download_dir + os_separator + name_folder.text

                driver.command_executor._commands["send_command"] = (
                    "POST", '/session/$sessionId/chromium/send_command')
                params = {'cmd': 'Page.setDownloadBehavior',
                          'params': {'behavior': 'allow', 'downloadPath': download_dir_temp}}
                driver.execute("send_command", params)

                # Mientras haya un boton de siguiente en la ultima tabla...
                while True:
                    files_list = []
                    try:
                        # Espero a que aparezca la lista de archivos
                        WebDriverWait(driver, 5).until(expected_conditions.visibility_of_element_located(
                            (By.XPATH, "/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[3]/tbody")))
                        # Consigo la lista de archivos
                        files_list = driver.find_elements(By.XPATH,
                                                          '/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[3]/tbody/tr')
                    except:
                        files_list = driver.find_elements(By.XPATH,
                                                          '/html/body/div[1]/div[2]/div[4]/div[1]/div[7]/form/div/table[2]/tbody/tr')

                    # Recorro la lista de archivos
                    for archivo in files_list[1:]:
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

                time.sleep(4)

                driver.quit()
                driver = use_driver()

                print(e)
                continue
            break

    time.sleep(5)

    driver.quit()
    messagebox.showinfo("Proceso Terminado", "Expedientes descargados")


def chunkList(list, num):
    avg = len(list) / float(num)
    out = []
    last = 0.0

    while last < len(list):
        out.append(list[int(last):int(last + avg)])
        last += avg

    return out


def get_latest_driver():
    url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
    response = requests.get(url)
    latest_version = response.text

    if os.name == 'nt':
        download_url = "https://chromedriver.storage.googleapis.com/%s/chromedriver_win32.zip" % latest_version
    else:
        download_url = "https://chromedriver.storage.googleapis.com/%s/chromedriver_linux64.zip" % latest_version

    # download the zip file using the url built above
    latest_driver_zip = wget.download(download_url, 'chromedriver.zip')

    # extract the zip file
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall()  # you can specify the destination folder path here
    # delete the zip file downloaded above
    os.remove(latest_driver_zip)


def use_driver():
    if os.name == 'nt':
        chromedriver = APP_PATH + "\\chromedriver.exe"
        download_dir = APP_PATH + "\\expedientes"
    else:
        chromedriver = APP_PATH + "/chromedriver"
        download_dir = APP_PATH + "/expedientes"

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


if __name__ == '__main__':
    root = Tk()
    make_main_view(root)
    root.mainloop()
