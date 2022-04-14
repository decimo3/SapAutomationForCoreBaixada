# coding: utf-8
import time
import subprocess
# import login
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

winium = subprocess.Popen("C:\\Users\\ruan.camello\\Documents\\Development\\winium\\Winium.Desktop.Driver.exe")
time.sleep(10)

#Abrindo o SAP Gui
driver = webdriver.Remote(
    command_executor='http://localhost:9999',
    options=webdriver.ChromeOptions(
        {
        'app': r'C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe'
    })
    # desired_capabilities={
    #     'app': r'C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe'
    # }
    )

#Clicando no ambiente que deseja acessar no SAP Gui.
time.sleep(5)
elemento = driver.find_element_by_name("QAS").click()
actionchains = ActionChains(driver)
actionchains.double_click(elemento).perform()

#Passando usuário e senha para logar.
usuario = "2258038" #login.usuario()
senha = "Light#44" #login.senha()

time.sleep(3)
elemento = driver.find_element_by_id("1004").click()
elemento = driver.find_element_by_id("1004")
elemento.send_keys(usuario)

elemento = driver.find_element_by_id("1005").click()
elemento = driver.find_element_by_id("1005")
elemento.send_keys(senha)

elemento = driver.find_element_by_name("Logon").click()

time.sleep(10)


elemento = driver.find_element_by_id("1001").click()
elemento = driver.find_element_by_id("1001")
elemento.send_keys('ME2N')
elemento.submit()

time.sleep(2)

elemento = driver.find_element_by_id("4004").click() #chamar variante
time.sleep(42)

#Fechando o SAP Gui e winium
driver.close()
winium.terminate()