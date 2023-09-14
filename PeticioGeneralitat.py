import csv
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time
from commonUtils import configParams

config = configParams()
params = config.loadConfig()

def read_csv_to_list(file_path):
    data_list = []
    
    with open(file_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            data_list.append(row)
    
    return data_list

def creaPeticio(peticio, dades, driver):

    codi = dades['Codi']
    document = peticio["Document"]
    psf=peticio["PSF"]
    pdfName = os.path.splitext(f'{codi}-{document}')[0] + '.pdf'
    pathPDF = os.path.join(params["carpetaPDFs"], pdfName)

    pathPDFRebutName = f"{codi}-{psf}-Acusament_rebuda.pdf"

    # Open the website
    driver.get(params["webTramits"])

    wait = WebDriverWait(driver, timeout=5)

    try:
        # Accepta pop-ups cookies
        cookies_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.cookieConsent__Button')))
        cookies_button.click()

        cookies_accept_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.cookieConsent__Button.cookieConsent__Button--Close')))
        cookies_accept_button.click()

    except:
        pass

    wait = WebDriverWait(driver, timeout=20)

    #driver.get(params["webTramits"])

    # Step 1
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.step-link.openstep"))).click()

    # Per Internet
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "p.title.internet"))).click()

    # Inici Tramit
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="acordio-t-1-1"]/div[1]/div[2]/a'))).click()

    # Tramit sense ID electrònic
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="article"]/div/div[1]/div[1]/div[3]/input'))).click()

    # Fill in Form
    peticioBox = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="codiPersonal-input"]')))
    peticioBox.send_keys(peticio["Peticio"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Nom"]')))
    nomBox.send_keys(dades["Nom"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Primer cognom"]')))
    nomBox.send_keys(dades["Cognom1"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Segon cognom"]')))
    nomBox.send_keys(dades["Cognom2"])

    select_element = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Tipus de document d\'identificació"]')))
    select_element.send_keys("DNI")

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Número d\'identificació"]')))
    nomBox.send_keys(dades["DNI"])

    dateBirth = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="text" and contains(@aria-label, "Data de naixement ")]')))
    dateBirth.send_keys(dades["DataNaixement"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Adreça electrònica"]')))
    nomBox.send_keys(dades["CorreuElectronic"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Telèfon mòbil"]')))
    nomBox.send_keys(dades["Mobil"])

    checkbox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Correu electrònic"]')))
    checkbox.click()

    select_element = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Tipus de via"]')))
    select_element.send_keys(dades["TipusDeVia"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Nom de la via"]')))
    nomBox.send_keys(dades["NomDeLaVia"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Número"]')))
    nomBox.send_keys(dades["Numero"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Pis"]')))
    nomBox.send_keys(dades["Pis"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Porta"]')))
    nomBox.send_keys(dades["Porta"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Codi postal"]')))
    nomBox.send_keys(dades["CodiPostal"])

    select_element_prov = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Província"]')))
    dropdown_prov = Select(select_element_prov)
    dropdown_prov.select_by_visible_text(dades["Provincia"])
    time.sleep(1)

    select_element_comarca = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Comarca"]')))
    dropdown_comarca = Select(select_element_comarca)
    dropdown_comarca.select_by_visible_text(dades["Comarca"])
    time.sleep(1)

    select_element_municipi = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Municipi"]')))
    select_element_municipi.send_keys(dades["Municipi"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//textarea[@aria-label="Assumpte"]')))
    nomBox.send_keys(peticio["Assumpte"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//textarea[@aria-label="Exposo"]')))
    nomBox.send_keys(peticio["Exposo"])

    nomBox = wait.until(EC.presence_of_element_located((By.XPATH, '//textarea[@aria-label="Sol·licito"]')))
    nomBox.send_keys(peticio["Sollicito"])

    select_element_destinatari = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@aria-label="Ens destinatari"]')))
    dropdown_destinatari = Select(select_element_destinatari)
    dropdown_destinatari.select_by_visible_text(peticio["Destinatari"])
    
    element = wait.until(EC.presence_of_element_located((By.XPATH, '//span[@class="iconButton-label" and @data-guide-button-label="true" and text()="Copiar adreça"]')))

    label_element = driver.find_element(By.XPATH, '//label[text()="Persona física"]')
    label_element.click()

    # Per identificar el segon Correu Electronic
    allEmailElements = driver.find_elements(By.XPATH, '//div[@class="guideFieldWidget textField" and @data-original-title=""]//input[@aria-label="Adreça electrònica"]')
    allEmailElements[2].send_keys(dades["CorreuElectronic"])

    # Upload del document PDF
    file_input = driver.find_elements(By.XPATH, '//div[@class="guideFieldWidget afFileUpload fileUpload"]/input[@type="file"]')
    file_input[1].send_keys(pathPDF)

    # Dades del sol.licitant
    select_element_prov = driver.find_elements(By.XPATH, '//select[@aria-label="Província"]')
    select_element_prov[2].send_keys(dades["Provincia"])
    time.sleep(1)
	
    select_element = driver.find_elements(By.XPATH, '//select[@aria-label="Tipus de via"]')
    select_element[2].send_keys(dades["TipusDeVia"])

    select_element_comarca = driver.find_elements(By.XPATH, '//select[@aria-label="Comarca"]')
    select_element_comarca[2].send_keys(dades["Comarca"])
    time.sleep(1)

    nomBox = driver.find_elements(By.XPATH, '//input[@aria-label="Nom de la via"]')
    nomBox[2].send_keys(dades["NomDeLaVia"])

    nomBox = driver.find_elements(By.XPATH, '//input[@aria-label="Número"]')
    nomBox[2].send_keys(dades["Numero"])

    nomBox = driver.find_elements(By.XPATH, '//input[@aria-label="Pis"]')
    nomBox[2].send_keys(dades["Pis"])

    nomBox = driver.find_elements(By.XPATH, '//input[@aria-label="Porta"]')
    nomBox[2].send_keys(dades["Porta"])

    nomBox = driver.find_elements(By.XPATH, '//input[@aria-label="Codi postal"]')
    nomBox[4].send_keys(dades["CodiPostal"])

    select_element_municipi = driver.find_elements(By.XPATH, '//select[@aria-label="Municipi"]')
    select_element_municipi[2].send_keys(dades["Municipi"])

    checkbox_element = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="checkbox" and @aria-label="He llegit i accepto la informació bàsica sobre protecció de dades"]')))
    checkbox_element.click()

    input("Please take your manual action and then press Enter to continue...")

    button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="dadesJsonFormButton"]')))
    button.click()

    # Baixa el rebut del registre
    link = wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Descarregueu el rebut de registre")))
    link.click()

    # Cambia el nom de l'arxiu descarregat
    downloaded_file_path = os.path.join(params["carpetaPDFRebuts"], "Acusament_rebuda.pdf")
    time.sleep(7) #Dona temps de baixar l'arxiu
    if os.path.exists(downloaded_file_path):
        new_file_path = os.path.join(download_dir, pathPDFRebutName)
        os.rename(downloaded_file_path, new_file_path)
    else:
        print(f"No he trobat l'arxiu {downloaded_file_path}!")
        print(peticio)
        print(dades)
        exit()

peticions = read_csv_to_list(params["csvPeticions"])
signatures = read_csv_to_list(params["csvSignatures"])

resPeticions = any(peticio.get("Activar") == "Y" for peticio in peticions)
resSignatures = any(signatura.get("Peticio") != "Y" for signatura in signatures)

if resPeticions and resSignatures:
    # carpeta per descarregar els rebuts
    download_dir = params["carpetaPDFRebuts"]
    os.makedirs({params["carpetaPDFRebuts"]}, exist_ok=True)

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,   # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()

    for signatura in signatures:
        for peticio in peticions:
            if peticio['Activar'] == "Y":
                if signatura['Peticio'] != "Y":
                    creaPeticio(peticio, signatura, driver)
                    input("Vols continuar? Prem Enter per continuar...")

    driver.quit()
else:
    if not resPeticions: print("No hi ha cap petició activa!")
    if not resSignatures: print("No hi ha signatura per crear una petició!")
        
