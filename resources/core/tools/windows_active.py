import os
import time
import sys
import winreg
import json
import subprocess
import shlex
from colorama import Fore, Style

def load_json_blurkms(dat):
    ruta_base = os.path.dirname(os.path.dirname(__file__))
    if (dat == 0):
        name = "core.json"
    else:
        name = "config_blurkms.json"

    ruta_archivo = os.path.join(ruta_base, name)
    with open(ruta_archivo, "r", encoding="utf-8") as archivo:
        return json.load(archivo) 
def load_coredata():
    return load_json_blurkms(0)   
def load_configkms():
    return load_json_blurkms(1)


def clear():
    os.system("cls")
def duration(seg):
    time.sleep(seg)


def Bits_Windows():
    return "64 bits" if sys.maxsize > 2**32 else "32 bits"

def SO_Windows():
    try:
        clave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        product_name, _ = winreg.QueryValueEx(clave, "ProductName")
        current_build, _ = winreg.QueryValueEx(clave, "CurrentBuild")

        winreg.CloseKey(clave)

        if int(current_build) >= 22000:
            product_name = product_name.replace("Windows 10", "Windows 11")

        return product_name
    except Exception:
        return "Error: No se pudo obtener información del sistema."
    


def buscador_clave_kms(clave):
        kms_data = load_coredata()
        claves = kms_data.get("Claves_KMS_Windows", {})
        return claves.get(clave, None)

def execute_cmd(CMD):
    subprocess.run(shlex.split(CMD), shell=True)




def activador_kms_windows_funcion(version_win, code_kms):
    data = load_coredata()
    linea_cmd = data["Activador_Windows"]

    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}╔{"═"*24}{Fore.GREEN + Style.NORMAL} [📲] {Style.BRIGHT + Fore.WHITE}Proceso de instalacion{Style.NORMAL + Fore.GREEN} [📲] {Style.BRIGHT + Fore.YELLOW}{"═"*25}╗\n{" "*8}{Fore.YELLOW}║{" "*83}║{Style.RESET_ALL}")
    
    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.GREEN}💻  Instalando clave KMS {" "*54}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*83}║{Style.RESET_ALL}")
    execute_cmd(f"slmgr /ipk {code_kms}")
    duration(2)


    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.BLUE}⚙️  Configurando administrador clave KMS {" "*38}{Fore.YELLOW + Style.BRIGHT}║")
    for line in linea_cmd:
       execute_cmd(line)
    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*8}{Style.RESET_ALL}{Style.DIM + Fore.BLUE}[ Administrador de Claves: kms.digiboy.ir ]{Style.RESET_ALL}{" "*32}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*83}║{Style.RESET_ALL}")
    duration(2)


    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.MAGENTA}🆗  Activando Clave {Style.RESET_ALL}{" "*59}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*83}║{Style.RESET_ALL}")
    duration(1)

    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.CYAN}👍  Listo!, {version_win} activado con exito! {Style.RESET_ALL}{" "*32}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*83}║{Style.RESET_ALL}")

    print(f"{" "*8}{Fore.YELLOW + Style.BRIGHT}╚{"═"*83}╝{Style.RESET_ALL}\n")

    print(f"{" "*8}{Fore.GREEN + Style.BRIGHT} NOTA: {Style.NORMAL}Es nesesario reiniciar, por favor cierre todas las aplicaciones abiertas\n{" "*10}antes de continuar.")
    input(f"{Fore.YELLOW + Style.DIM}\n{" "*12} <Presione Enter Para Reiniciar>")

    os.system("shutdown /r /t 0")