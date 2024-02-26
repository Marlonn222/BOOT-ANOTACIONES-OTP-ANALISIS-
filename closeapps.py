import subprocess
import pyautogui
from genericfunctions import (pressingKey)
from time import sleep

def disconnectVPNForticlient():
    pyautogui.getWindowsWithTitle("FortiClient")[0].maximize()
    sleep(1)
    pressingKey('tab')
    pressingKey('enter')
  
def closeFortiClient():
    pyautogui.getWindowsWithTitle("FortiClient")[0].close()
    
def closeCRM():
    if(pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")):
        pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].maximize()
        sleep(1)
        pyautogui.getWindowsWithTitle("Sistema Avanzado de Administración de Clientes [Versión 4.2.2.3]")[0].close()
        pressingKey('s')
        print("closeapps module has ended")
    else:
        print("CRM is not active")
        return
    
 