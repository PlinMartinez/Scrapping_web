from selenium import webdriver
from selenium.webdriver.ie.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import win32gui
import win32com
import win32com.client as comclt
import time
driver = webdriver.Ie('IEDriverServer.exe')
driver.get('https://www.movistar.es/l0/priv/servlet/Inicio')
time.sleep(1)
def accept_cookies():
    try:
        driver.find_element_by_xpath('//button[text()="Accept All Cookies"]').click()
    except:
        return False
    else:
        return True
def login():
    driver.switch_to.default_content()
    time.sleep(1)
    accept_cookies()
    accept_cookies()
    try:
        driver.switch_to.frame(driver.find_element_by_xpath('//Iframe[@id="iframe_principal"]'))
    except NoSuchElementException:
        return 'Not on Login page'
    nic = driver.find_element_by_xpath('//input[@id="concontrasena_usuario"]')
    pas = driver.find_element_by_xpath('//input[@id="concontrasena_clave"]')
    remember = driver.find_element_by_xpath('//input[@id="concontrasema_mantenerme"]')
    submit = driver.find_element_by_xpath('//button[@id="enterButtonUC"]')
    if remember.get_attribute('checked')==None:
        driver.execute_script("arguments[0].click();",remember)
    else:
        None
    driver.execute_script("arguments[0].value='50842426';", nic)
    nic.send_keys('B')
    driver.execute_script("arguments[0].value='PAZ2019';", pas)
    pas.send_keys('M')
    submit.click()
def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_list(handles=[]):
    mlst=[]
    win32gui.EnumWindows(window_enum_handler, handles)
    for handle in handles:
        mlst.append(handle)
    return mlst
def hit_save(filename):
    shell = win32com.client.Dispatch("WScript.Shell")
    appwindows = get_app_list()
    for i in appwindows:
        if "Internet Explorer" in i[1]:
            win_name = i[1]
    hwnd = win32gui.FindWindowEx(0,0,0, win_name)
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(2)
    shell.SendKeys("%n",0)
    shell.SendKeys("{TAB}",0)
    shell.SendKeys("{ENTER} ",0)
    print('Saving '+str(filename.text))
login()
time.sleep(2)
driver.switch_to.default_content()
driver.switch_to.frame(driver.find_element_by_xpath('//frame[@name="izquierdo"]'))
driver.find_element_by_xpath('//font[text()="ficheros de llamadas"]').click()
driver.switch_to.default_content()
driver.switch_to.frame(driver.find_element_by_xpath('//frame[@name="central"]'))
while True:
    try:
        if driver.find_element_by_xpath('//table').text=='CARGANDO PAGINA\nEspere por favor...':
            print('.',end='.')
        else:
            break
    except:
        break
for filename in driver.find_elements_by_xpath('//tbody/tr//a'):
    driver.execute_script("arguments[0].click();",filename)
    hit_save(filename)
print('Closing Internet Explorer!')
driver.close()
