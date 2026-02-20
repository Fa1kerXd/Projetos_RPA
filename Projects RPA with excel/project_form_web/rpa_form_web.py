"""
Modulo responsavel por preencher formulario de forma automatica capturando as informações
de uma planilha Excel.

Fluxo:
    1.
    2.
    3.


"""

# Importando biblioteca selenium para a automatização do fluxo.
# Pyautogui para dar um intervalo entre processos.
from selenium import webdriver as driver
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook, load_workbook
import pyautogui as pause
from pathlib import Path


COLUMNS = {
    'Nome': 'nome',
    'Email': 'email',
    'Telefone': 'telefone',
    'Sobre': 'sobre',
    'Gênero': 'genero'
}











# Capturando a URL passado no get para abrir o recurso com o driver selenium.
browser = driver.Edge()
browser.get("https://pt.surveymonkey.com/r/WLXYDX2")

name = browser.find_element(By.NAME, "166517069")

pause.sleep(3)

name.send_keys("Augusto Cesar")

pause.sleep(3)

email = browser.find_element(By.NAME, "166517072")

pause.sleep(3)

email.send_keys("fa1ker@icloud.com")

pause.sleep(3)

telefone = browser.find_element(By.NAME, "166517070")

pause.sleep(3)

telefone.send_keys("(31) 97185-0807")

pause.sleep(3)


click = browser.find_element(By.NAME, "question-field-166517071")
click.click()

pause.sleep(3)

about = browser.find_element(By.NAME, "166517073")

pause.sleep(3)

about.send_keys("Desenvolvedor RPA/Full stack/Django")

pause.sleep(3)

send = browser.find_element(By.XPATH, '//*[@id="view-pageNavigation"]/div/button')
send.click()
pause.sleep(4)

