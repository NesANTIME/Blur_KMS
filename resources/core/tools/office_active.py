import os
import json
import winreg
import shlex
import subprocess
import time
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



def functios_for_cmdoffice(mode, command):
    if mode == "cmd_ospp":
        Ospp_rut = Rut_OSPP()
        if Rut_OSPP() is None:
            return False
        
        subprocess.run(f'cscript "{Ospp_rut}" {command}', shell=True, check=True)

    elif mode == "Shell":

        subprocess.run(shlex.split(command), shell=True, capture_output=True, text=True)
    elif mode == "cmd_ospp-Debug":
            
            try:
                resul = subprocess.run(f'cscript "{Rut_OSPP()}" {command}', shell=True, capture_output=True, text=True)
                return resul.stdout + resul.stderr
            except Exception as e:
                return str(e)
    else:
        subprocess.run(command, shell=True, check=True, capture_output=True, text=True)



def functions_for_office(dat):
    load_core = load_coredata()
    if (dat == 0):
        rutas_office = load_core["Rutas_Verificar_Office"]
        for exix in rutas_office:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, exix):
                    return True
            except FileNotFoundError:
                continue
        
        return False
    
    elif (dat == 1):
        rutas_ospp = load_core["Rutas_Verificar_Ospp"]
        for ruta in rutas_ospp:
            if os.path.exists(ruta):
                return ruta
        return None
    
    elif (dat == 2):
        rutas_offic = load_core["Rutas_Office"]
        for path in rutas_offic:
            new_path = os.path.expandvars(path)
            if os.path.exists(new_path):
                functios_for_cmdoffice("Shell", f'cd /d "{new_path}"')
                break
            else:
                return False
def detect_office():
    return functions_for_office(0)
def Rut_OSPP():
    return functions_for_office(1)
def Office_Ruta():
    return functions_for_office(2)

    

def Funcion_ActivationOffice(Office):
    Ospp_rut = Rut_OSPP()
    data = load_coredata()

    Office_Ruta()

    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}╔{"═"*20}{Fore.GREEN + Style.NORMAL} [📲] {Style.BRIGHT + Fore.WHITE}Proceso de conversion a licencia volumen{Style.NORMAL + Fore.GREEN} [📲] {Style.BRIGHT + Fore.YELLOW}{"═"*20}╗\n{" "*8}{Fore.YELLOW}║{" "*92}║{Style.RESET_ALL}")
    
    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.GREEN}💻  Convirtiendo Office a Licencia Volumen {" "*45}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*92}║{Style.RESET_ALL}")
    codeLisense = data.get("Converter_Lisencevolumen_Office", {})
    code = codeLisense.get(Office, "Clave no encontrada")

    duration(2)

    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.BLUE}⚙️  Instalando Clave KMS {" "*63}{Fore.YELLOW + Style.BRIGHT}║")
    functios_for_cmdoffice("Shell", code)
    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*8}{Style.RESET_ALL}{Style.DIM + Fore.BLUE}[ Accion Ejecutada Exitosamente ]{Style.RESET_ALL}{" "*51}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*92}║{Style.RESET_ALL}")

    print(f"{Fore.YELLOW + Style.BRIGHT}{" "*8}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.MAGENTA}🆗  Habilitando Registros {Style.RESET_ALL}{" "*62}{Fore.YELLOW + Style.BRIGHT}║\n{" "*8}{Fore.YELLOW}║{" "*92}║{Style.RESET_ALL}")
    duration(1)

    print(f"{" "*8}{Fore.YELLOW + Style.BRIGHT}╚{"═"*92}╝{Style.RESET_ALL}\n")

    if Ospp_rut == None:
        print(f"{Fore.RED}{" "*11}[ERROR] No se encontró ospp.vbs en las rutas esperadas." + Fore.WHITE)
    else:
        lineactivate = data[f"Activador_{Office}"]
        for line in lineactivate:
            functios_for_cmdoffice("cmd_ospp", line)

    duration(2)

    line_cmd = functios_for_cmdoffice("cmd_ospp-Debug", "/act")
    if "ERROR" in line_cmd or "No Office KMS licenses were found" in line_cmd:
        print(f"\n{Fore.RED}{" "*11}[!] Activación fallida. Revisa los errores:\n {Fore.YELLOW} {line_cmd} {Style.RESET_ALL}")
        print(f"{Fore.GREEN + Style.DIM}{" "*11}[!] Es posible que ya office este activado.{Style.RESET_ALL}")
    else:
        print(f"{Fore.GREEN + Style.DIM}{" "*11}[!] Activacion Completada. Activado Con Exito.{Style.RESET_ALL}")


    duration(1)
    input(f"{Fore.YELLOW + Style.DIM}\n{" "*12} <Presione Enter Para Finalizar>")
