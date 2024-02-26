import os
import subprocess
import pyperclip
import pyautogui
from dotenv import load_dotenv
from genericfunctions import (pressingKey,make_window_visible)
from time import sleep

# Open VPN Client and Connect to NAE VPN
def openAndConnectVPNForticlient():    
    forticlient_connected = None
    subprocess.Popen('C:\\Program Files\\Fortinet\\FortiClient\\FortiClient.exe')
    sleep(3)
    pressingKey('tab',2)
    sleep(2)
    pyautogui.write('dgard')    
    sleep(1)
    pressingKey('tab',1)
    sleep(1)
    pyautogui.write('Colombia2023*')
    pressingKey('enter')
    while forticlient_connected is None:
        forticlient_connected = pyautogui.locateOnScreen('C:/ActualizacionComentarios/assets/forti_client_connected.png', grayscale = True,confidence=0.9)    
    print("FortiClient Connected!")
    sleep(1)
    pyautogui.getWindowsWithTitle("FortiClient")[0].minimize()

# Open CRM Onyx App
def openCRM():    
    sleep(1)
    pressingKey('win')
    sleep(1)
    pyautogui.write('Nuevo CRM')    
    sleep(1)
    pressingKey('enter')
    make_window_visible('Control de Acceso Unificado CRM')
    
# Connect CRM Onyx App    
def connectToCRM():    
    crm_select_app = None
    crm_dashboard = None       
    
    pressingKey('tab',2)
    sleep(1)
    pyautogui.write(os.getenv('USER_CRM'))
    sleep(1)
    pressingKey('tab') 
    pyperclip.copy(os.getenv('PASS_CRM'))
    sleep(0.5)
    pyautogui.hotkey('ctrl','v')
    sleep(1)
    pressingKey('tab')
    sleep(2)
    pressingKey('enter')
    # se debe corer con la version pip install pyScreeze==0.1.28
    print(os.path.join("C:/ActualizacionComentarios/assets/select_app.png"))
    while crm_select_app is None:
        print("buscando ventana emergente!")
        crm_select_app = pyautogui.locateOnScreen('C:/ActualizacionComentarios/assets/select_app.png', grayscale = True,confidence=0.9)
    print("CRM select box is present!")

    pressingKey('enter')
    pressingKey('tab',4)
    pressingKey('enter')
    
    while crm_dashboard is None:
        crm_dashboard = pyautogui.locateOnScreen('C:/ActualizacionComentarios/assets/crm_dashboard.png', grayscale = True,confidence=0.85)    
    print("CRM Dashboard GUI is present!")    
    