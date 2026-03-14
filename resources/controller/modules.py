import os
import sys
import time
import shlex
import socket
import winreg
import subprocess

from resources.files.load_config import return_data


def clear():
    os.system("cls")

def delay(seg: int):
    time.sleep(seg)

def detect_controller(mode):
    if (mode == "system"):
        try:
            clave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
            product_name, _ = winreg.QueryValueEx(clave, "ProductName")
            current_build, _ = winreg.QueryValueEx(clave, "CurrentBuild")

            winreg.CloseKey(clave)

            if int(current_build) >= 22000:
                product_name = product_name.replace("Windows 10", "Windows 11")

            return product_name
        
        except Exception:
            return False
    
    elif (mode == "arquitecture"):
        return "64 bits" if sys.maxsize > 2**32 else "32 bits"
    
def connect_internet() -> bool:
    hosts = return_data("return_hosts")

    for host in hosts:
        try:
            socket.setdefaulttimeout(3)
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, 53))
            return True
        except socket.error:
            continue
    return False




class Launch_KMS:
    def __init__(self, code_kms):
        self.code_kms = code_kms

    def execute_P1(self):
        return subprocess.Popen(
            shlex.split(f"slmgr /ipk {self.code_kms}"), 
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, 
            text=True
        )    

    def execute_P2(self):
        return subprocess.Popen(
            shlex.split("slmgr /skms kms.digiboy.ir"), 
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, 
            text=True
        )

    def execute_P3(self):
        return subprocess.Popen(
            shlex.split("slmgr /ato"), 
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT, 
            text=True
        )
    
    def finalizar(self):
        subprocess.run(["shutdown", "/r", "/t", "0"])