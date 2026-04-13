# Code not to be use as its version gets expired, but to be used as a base for the next versions, in case of any problem with the new code.


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from tkinter import messagebox
import os

resposta = messagebox.askquestion("Confirmação", "Deseja prosseguir com a extração?")
if resposta == 'yes': 
    try:
        # Substitua o caminho abaixo pelo caminho onde você baixou o ChromeDriver

        script_dir = os.path.dirname(os.path.abspath(__file__))

        chrome_driver_path =os.path.join(script_dir, "Driver", "chromedriver.exe")


        # Configurando o serviço do ChromeDriver
        chrome_service = Service(chrome_driver_path)

        # Configurando o driver do Selenium com o ChromeDriver
        driver = webdriver.Chrome(service=chrome_service)

        url = 'https://grouppurchasing.fiat.com/irj/portal/gssm?standAlone=true&sapDocumentRenderingMode=Edge&HistoryMode=1&TarTitle=Source%20Package%20Management&windowId=WID1703160349761&NavMode=0'

        driver.get(url)

        driver.maximize_window()

        # Definindo uma espera explícita de 10 segundos
        wait = WebDriverWait(driver, 10)
        wait_2 = WebDriverWait(driver, 300)

        # Aguardando até que o elemento esteja visível e clica nele
        element_user = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="logonuidfield"]')))
        element_user = driver.find_element(By.XPATH, '//*[@id="logonuidfield"]')
        ############ LOGIN AQUI ############
        element_user.send_keys('SD62318')

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_password = driver.find_element(By.XPATH, '//*[@id="logonpassfield"]')
        ############ SENHA AQUI ############
        element_password.send_keys('Stellantis1997')

        time.sleep(1)

        element_login = driver.find_element(By.XPATH, '/html/body/span/table/tbody/tr/td/table/tbody/tr/td/form/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[2]/input')
        element_login.click()

        time.sleep(3)

        # Aguardando até que o elemento esteja visível e clica nele
        element_application = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr[1]/td/table/tbody/tr[1]/td/div[2]/ul[2]/div/li[2]')))
        element_application = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/table/tbody/tr[1]/td/div[2]/ul[2]/div/li[2]')
        element_application.click()

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_globalSourcingTool = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]')))
        element_globalSourcingTool = driver.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]')
        element_globalSourcingTool.click()

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_sourcePackageManagement = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/img')))
        element_sourcePackageManagement = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/img')
        element_sourcePackageManagement.click()

        time.sleep(1)

        wait.until(EC.frame_to_be_available_and_switch_to_it("ivuFrm_page0ivu1"))

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_reporting_numeroSourcePackage = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[3]/div[1]')))
        element_sourcePackageDashboard_reporting = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[3]/div[1]')
        element_sourcePackageDashboard_reporting.click()

        time.sleep(1)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_reporting_display = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[1]/div/div/table/tbody/tr/td[1]/div/span/span/div')))
        element_reporting_display = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[1]/div/div/table/tbody/tr/td[1]/div/span/span/div')
        element_reporting_display.click()

        time.sleep(5)

        element_all_modification = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[16]/td[3]/span/input')
        element_all_modification.send_keys("Semana Anterior")

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ENTER).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ARROW_UP).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ARROW_UP).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ENTER).perform()

        time.sleep(1)   

        element_reporting_status = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[13]/td[3]/span/input')
        element_reporting_status.send_keys('Technical Data Completed')

        time.sleep(1)        

        element_reporting_view = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[1]/span[2]/input')
        element_reporting_view.send_keys('RPA')

        time.sleep(2)     

        element_reporting_searchbutton = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        element_reporting_searchbutton.click()

        time.sleep(5)   

        element_reporting_downloadbutton = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[3]/div/table/tbody/tr/td/div/div/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[3]/div')
        element_reporting_downloadbutton.click()

        time.sleep(5)

        driver.switch_to.default_content()

        time.sleep(1)

        element_topheader_gst = driver.find_element(By.XPATH, '//*[@id="gssm_breadcrumb"]/div/a[2]')
        element_topheader_gst.click()

        time.sleep(1)

        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[6]/td[1]/table/tbody/tr[1]/td[1]/img')))
        element_sourcingManagement = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/div/table/tbody/tr[6]/td[1]/table/tbody/tr[1]/td[1]/img')
        element_sourcingManagement.click()

        time.sleep(1)

        wait.until(EC.frame_to_be_available_and_switch_to_it("ivuFrm_page0ivu1"))

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        # Aguardando até que o elemento esteja visível e clica nele
        element_sourceProcessDashboard_all = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[2]/div[1]')))
        element_sourceProcessDashboard_all = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td/table/tbody/tr/td[2]/div/div[2]/div[1]')
        element_sourceProcessDashboard_all.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        element_all_modification = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[11]/td[3]/span/input')
        element_all_modification.send_keys("Semana Anterior")

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ENTER).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ARROW_UP).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ARROW_UP).perform()

        time.sleep(1)

        ActionChains(driver).send_keys(Keys.ENTER).perform()

        time.sleep(1)

        element_all_regiao = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input')
        element_all_regiao.send_keys('Global')

        time.sleep(1)

        element_all_view = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[1]/span[2]/input')
        element_all_view.send_keys('RPA')

        time.sleep(3)

        element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        element_all_search.click()

        # element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        # element_all_search.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        element_all_download.click()

        # element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        # element_all_download.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        element_all_regiao = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input')
        element_all_regiao.send_keys('LATAM')

        # element_all_regiao = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input')
        # element_all_regiao.send_keys('LATAM')

        time.sleep(1)

        element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        element_all_search.click()

        # element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        # element_all_search.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        element_all_download.click()

        # element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        # element_all_download.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        element_all_regiao = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input')
        element_all_regiao.send_keys('Neutral')

        # element_all_regiao = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[1]/td/div/div/div[2]/span/span/table/tbody/tr/td/div/table/tbody/tr[8]/td[3]/span/input')
        # element_all_regiao.send_keys('Neutral')

        time.sleep(1)

        element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        element_all_search.click()

        # element_all_search = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[2]/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/span[1]/div')
        # element_all_search.click()

        time.sleep(3)

        try:
            # Lidando com o carregamento
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            try:
                wait_2.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="ur-loading-itm2"]')))
            except:
                pass
        except:
            pass

        time.sleep(1)

        # element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        # element_all_download.click()

        element_all_download = driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr/td/div/table/tbody/tr[2]/td/div/table/tbody/tr[3]/td/div[2]/div/table/tbody/tr/td/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/span/span/table/tbody/tr/td/span/span[4]/table/tbody/tr/td/span/span[1]/div/div/div/span/span/table/tbody/tr[1]/td/div/div/div/div[1]/span[9]/div')
        element_all_download.click()

        time.sleep(10)

        driver.quit()

        messagebox.showinfo("Sucesso", "Extração realizada com sucesso !")

    except:
        messagebox.showerror("Erro", "Erro ao executar a extração !")
        messagebox.showinfo("Aviso", "Tente novamente! Caso o erro persista, entre em contato com o suporte.")

else:
    # Coloque aqui o código que deseja executar se o usuário cancelar
    messagebox.showwarning("Cancelado", "Operação cancelada !")