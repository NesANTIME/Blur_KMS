import os
import sys
import json
import time
import shlex
import winreg
import requests
import subprocess
import tkinter as tk
import platform as plataform
from PIL import Image, ImageTk
from colorama import Fore, init, Style
init()

def Icon():
    clear()
    print(" \n")
    print(Fore.CYAN + "    ██████╗ ██╗     ██╗   ██╗██████╗               ██╗  ██╗███╗   ███╗███████╗")
    print("    ██╔══██╗██║     ██║   ██║██╔══██╗              ██║ ██╔╝████╗ ████║██╔════╝")
    print("    ██████╔╝██║     ██║   ██║██████╔╝    █████╗    █████╔╝ ██╔████╔██║███████╗")
    print("    ██╔══██╗██║     ██║   ██║██╔══██╗    ╚════╝    ██╔═██╗ ██║╚██╔╝██║╚════██║")
    print("    ██████╔╝███████╗╚██████╔╝██║  ██║              ██║  ██╗██║ ╚═╝ ██║███████║")
    print("    ╚═════╝ ╚══════╝ ╚═════╝ ╚═╝  ╚═╝              ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝")
    print(Fore.BLUE + f"      {Fore.CYAN + Style.DIM}Blur-KMS v4.0{Fore.BLUE + Style.NORMAL} By NesAnTime" + Fore.WHITE)

def clear():
    if plataform.system() == "Windows":
        os.system("cls")

def Load_BD():
    with open("Scritps/BD_BlurKMS.json", "r", encoding="utf-8") as archivo:
        return json.load(archivo)
    
def refresh(dat):
    time.sleep(dat)

def Search(dato):
        kms_data = Load_BD()
        claves = kms_data.get("Claves_KMS_Windows", {})
        return claves.get(dato, "Clave no encontrada")

def Internet_Conexion():
    try:
        subprocess.check_call(["ping", "-n", "1", "8.8.8.8"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False
    
def Command(CMD, mode):
    if mode == "Prompt":
        subprocess.run(shlex.split(CMD), shell=True)
    elif mode == "cmd_ospp":
        Ospp_rut = Rut_OSPP()
        subprocess.run(f'cscript "{Ospp_rut}" {CMD}', shell=True, check=True)
    elif mode == "Shell":
        subprocess.run(shlex.split(CMD), shell=True, capture_output=True, text=True)
    elif mode == "cmd_ospp-Debug":
            try:
                resul = subprocess.run(f'cscript "{Rut_OSPP()}" {CMD}', shell=True, capture_output=True, text=True)
                return resul.stdout + resul.stderr
            except Exception as e:
                return str(e)
    else:
        subprocess.run(shlex.split(CMD), shell=True, check=True, capture_output=True, text=True)

    
# VERIFICADORES ------------------------------------------

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
    
def Office_Version():
    datos = Load_BD()
    Codes = datos["Rutas_Verificar_Office"]
    
    print("Buscando Microsoft Office en el registro de Windows...")
    
    for clave in Codes:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, clave):
                return True
        except FileNotFoundError:
            continue
    
    return None

def Rut_OSPP():
    datos = Load_BD()
    rutas = datos["Rutas_Verificar_Ospp"]
    for ruta in rutas:
        if os.path.exists(ruta):
            return ruta
    return None

def Office_Ruta():
    datos = Load_BD()
    rut = datos["Rutas_Office"]

    for path in rut:
        new_path = os.path.expandvars(path)
        if os.path.exists(new_path):
            Command(CMD=f'cd /d "{new_path}"', mode="Shell")
            break
    else:
        print(Fore.RED + "[!] No se encontró la ruta de instalación de Microsoft Office." + Fore.WHITE)
        return

    
# --------------------------------------------------------

#UPDATER --------
def Update():  
    datos = Load_BD()      
    VerLocal = datos["Versiones_BlurKMS"][0]
    VerNew = datos["Versiones_BlurKMS"][1]
    
    
    def V_NewVer(VerNew):
        try:
            response = requests.get(VerNew)
            if response.status_code == 200:
                remote_data = response.json()
                return remote_data["Versiones_BlurKMS"][0]
            else:
                print(f"[!] Error al obtener el archivo remoto, código de estado: {response.status_code}")
                return
        except requests.exceptions.RequestException as e:
            print(f"[!] Error al conectarse: {e}")
            return
        
    def V_LocalVer(VerLocal, V_NewUp):
        try:
            if VerLocal == V_NewUp:
                return f"{Fore.RESET + Style.RESET_ALL}"
            else:
                return f"{Fore.CYAN + Style.BRIGHT}[!] {Fore.WHITE}Hay una nueva version disponible {Style.DIM + V_NewUp + Fore.GREEN} {Style.BRIGHT}¡Descargala YA!{Fore.RESET + Style.NORMAL}\n"
        except Exception as e:
            print(f"{Fore.RED}[!] Error al leer el archivo local {VerLocal}: {e}")
            
    V_NewUp = V_NewVer(VerNew)
    if V_NewUp is not None:
        return V_LocalVer(VerLocal, V_NewUp)
    else:
        return Fore.RED + "[!] No se pudo obtener el contenido remoto para comparar."
#---------------

# Complementos de Funciones Principales  --------------------------

def SO_Edit():
    clear()
    Icon()
    print(Fore.CYAN + "\n[!] Elije La Version De Tu Sistema Operativo: " + Fore.WHITE)

    datos = Load_BD()
    Versiones = datos["SO_EditCache"]
    print("\n".join([f"{Style.DIM}[{i+1}] {Style.NORMAL}{Versiones[i]}" for i in range(len(Versiones))]))

    while True:
        try:
            opc = input("\nSeleccione una opcion: ")

            if opc.isdigit():
                opcion = int(opc)
                if 1 <= opcion <= len(Versiones):
                    print(Fore.GREEN + f"\n[!] Has seleccionado: {Versiones[opcion - 1]}" + Fore.WHITE)
                    refresh(1)
                    return Versiones[opcion - 1]
                else:
                    print(Fore.RED + "\n[!] Opción inválida. Debe estar dentro del rango." + Fore.WHITE)
            else:
                print(Fore.RED + "\n[!] Entrada no válida. Debe ser un número." + Fore.WHITE)

        except Exception as e:
            print(Fore.RED + f"[!] Error inesperado: {e}" + Fore.WHITE)
            refresh(2)

#--------------------------------------------------

# FUNCIONES DE ACTIVACION ------------------------

def Funcion_ActivationSO(SO_Version, KMS_Code):
    data = Load_BD()
    lineactivate = data["Activador_Windows"]
    print("\n[!] Iniciando Instalacion de Claves... ")
    Command(CMD = f"slmgr /ipk {KMS_Code}", mode = "Prompt")

    refresh(2)
    print("\n[!] Configurando Administrador de Claves...")
    for line in lineactivate:
        Command(CMD=line, mode = "Prompt")
    print(Style.DIM + "     [ Administrador de Claves: kms.digiboy.ir ]" + Style.NORMAL)

    refresh(2)
    print("\n[!] Activando Clave... ")

    refresh(1)
    print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)
    print(Fore.GREEN + Style.DIM + f"       [!] Activacion Completada. {SO_Version} Activado Con Exito.")
    print(Fore.GREEN + f"           Es Nesesario Reiniciar, Para Completar. (Puede Reiniciar Mas Tarde)" + Style.NORMAL)
    refresh(1)
    input(Fore.RED + "\nPresione Enter Para Finalizar...")



def Funcion_ActivationOffice(Office):
    print(Fore.GREEN + "\n[!] Iniciando Instalacion de Claves KMS En Office... " + Fore.WHITE)
    data = Load_BD()
    code = data.get("Converter_Lisencevolumen_Office", {})
    Office_Ruta()
    Ospp_rut = Rut_OSPP()
    refresh(2)

    print(Fore.GREEN + "\n[!] Convirtiendo Office a Licencia Volumen... ")

    comando = code.get(Office, "Clave no encontrada")
    Command(comando, mode = "Shell")
    
    print(Style.DIM + "  | Accion Ejecutada Exitosamente..." + Style.NORMAL)
    refresh(2)
    print(Fore.GREEN + "\n[!] Activando Clave... " + Fore.WHITE)

    print("\n [!!] Habilitando Registros \n " + Style.DIM)

    if Ospp_rut == None:
        print(Fore.RED + "[ERROR] No se encontró ospp.vbs en las rutas esperadas." + Fore.WHITE)
    else:
        app = f"Activador_{Office}"
        lineactivate = data[app]
        for line in lineactivate:
            Command(CMD=line, mode = "cmd_ospp")

    refresh(2)
    exit = Command(CMD = "/act", mode = "cmd_ospp-Debug")
    print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)

    if "ERROR" in exit or "No Office KMS licenses were found" in exit:
        print(Fore.RED + "\n[!] Activación fallida. Revisa los errores:\n" + Fore.YELLOW + exit + Fore.WHITE)
    else:
        print(Fore.GREEN + Style.DIM + f"[!] Activacion Completada. Activado Con Exito.")

    refresh(1)
    input(Fore.RED + "\nPresione Enter Para Finalizar...")

#--------------------------------------------------

# FUNCIONES PRINCIPALES & PREPARACION------------------------
    
def Main_KMSWindows(SO_Version, SO_Bits):
    while True:
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
        print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + SO_Version}")
        print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + SO_Bits} \n")
        print(Fore.YELLOW + Style.DIM + f"     [!] Si Considera un Error La Version Obtenida De Windows, O la Arquitectura. Elija una opcion a Continuacion." + Style.NORMAL + Fore.WHITE)
        print(f"{Style.DIM}          [1]{Style.NORMAL} Cambiar Version de Windows")
        print(f"{Style.DIM}          [2]{Style.NORMAL} Cambiar Arquitectura de Windows")
        print(f"{Style.DIM}          [3]{Style.NORMAL} Continuar Con Activacion de Windows")

        opc = input(f"\n{Fore.GREEN}     [¡] Esperando Opcion: " + Style.NORMAL + Fore.WHITE)

        if (opc == "1") or (opc == "Windows") or (opc == "WINDOWS"):
            SO_Version = SO_Edit()
            clear()

        elif (opc == "2") or (opc == "Bits") or (opc == "BITS"):
            print(Fore.GREEN +"\n[!] Su Sistema Operativo: "+ Style.DIM + SO_Version + Style.NORMAL + " Es de " + Style.DIM + Fore.WHITE + "[1] " + Style.NORMAL + "32 bits  -  " + Style.DIM + "[2] " + Style.NORMAL + "64 bits")
            try:
                Mip = int(input(f"  [!] Ingrese Su Opcion: {Style.DIM}"))
                print(Style.RESET_ALL)
                while (Mip < 1) or (Mip > 2):
                    print("Opcion Ingresada Invalida.")
                    Mip = int(input(f"  [!] Ingrese Nuevamente Su Opcion: {Style.DIM}"))
                if Mip == 1:
                    SO_Bits = "32 bits"
                elif Mip == 2:
                    SO_Bits = "64 bits"
            except ValueError:
                print(Fore.RED + "[!] Error. Entrada no válida.")
                refresh(2)
            clear()

        elif (opc == "3") or (opc == "continuar") or (opc == "Continuar"):
            clear()
            Icon()
            print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
            print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + SO_Version}")
            print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + SO_Bits} \n")

            print(Fore.GREEN + "[!] Iniciando Programa de Activacion...")

            KMS_Code = Search(SO_Version)

            if KMS_Code == "Clave no encontrada":
                print(Fore.RED + "[!] Sistema Operativo Desconocido, Sin Clave KMS Valida.")
            else:
                refresh(2)
                Funcion_ActivationSO(SO_Version, KMS_Code)
                break

        else:
            print(Fore.RED + "[!] Opcion Invalida.")
            refresh(2)


def Main_KMSOffice(Of_Exists):
    def Execute(Of_Version):
        clear()
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
        print(f"      {Style.BRIGHT + Fore.GREEN}[!]  {Of_Version + Style.NORMAL} \n")

        input(Fore.YELLOW + f"{Style.BRIGHT}[!] NOTA: {Style.DIM}La herramienta de claves KMS para Office tiene un 50% de probabilidades de \nresultar exitosa. Si obtiene un resultado desfaborable se pueden le proporcionar \nEl listado de claves genericas para Office." + Fore.WHITE + Style.NORMAL + " \n \n | Presione Enter para Continuar...")
        
        Icon()
        print(Fore.GREEN + "\n [!] Iniciando Programa de Activacion KMS...")
        while True:
            if Internet_Conexion() == True:
                print(Fore.GREEN + Style.DIM + "\n   ----- [!] Conectado a Internet." + Fore.WHITE + Style.NORMAL)
                break
            else:
                print(Fore.RED + Style.DIM + "\n    ----- [!] No se ha detectado conexión a Internet." + Fore.WHITE)
                print("   - Intentelo Nuevamente Conectado a Internet, Se Intentara Conectar Automaticamente.")
                refresh(2)

        Funcion_ActivationOffice(Of_Version)


    
    Icon()
    print(Fore.YELLOW + Style.DIM + "\n [¡] Buscando Microsoft Office en el registro de Windows...\n" + Style.NORMAL)
    refresh(2)

    if Of_Exists == None:
        print(f"      {Style.BRIGHT + Fore.RED}- [!] Microsoft Office no está instalado en este sistema. {Style.NORMAL}")
    else: 
        print(f"{Style.BRIGHT + Fore.GREEN}  [!] Microsoft Office Detectado \n")
        print(Fore.CYAN + "***    Menu de Versiones de Office   ***" + Fore.WHITE)
        print(f"{Style.DIM}     [1]{Style.NORMAL} Microsoft Office 2016")
        print(f"{Style.DIM}     [2]{Style.NORMAL} Microsoft Office 2019")
        print(f"{Style.DIM}     [3]{Style.NORMAL} Microsoft Office 2021")
        print(f"{Style.DIM}     [4]{Style.NORMAL} Claves Genericas Para Office {Style.DIM}(Claves Manuales){Style.NORMAL}\n")

        while True:
            try:
                opc = int(input(Fore.YELLOW + Style.DIM + f"[¡] Elija la Opcion Requerida: "+ Style.NORMAL + Fore.WHITE))
                if (opc == 1):
                    Execute(Of_Version = "MicrosoftOffice2016")
                    break
                elif (opc == 2):
                    Execute(Of_Version = "MicrosoftOffice2019")
                    break
                elif (opc == 3):
                    Execute(Of_Version = "MicrosoftOffice2021")
                    break
                elif (opc == 4):
                    print(Fore.CYAN + "\n    [!] Abriendo en el Navegador, Listado de Claves Genericas Para Office...")
                    refresh(2)
                    os.system("start Scritps/Codes_Office.html")
                    break 
                else:
                    print(Fore.RED + "[!] Opcion Invalida.\n")
                    refresh(2)
            except ValueError:
                print(Fore.RED + "[!] Entrada no válida. Debe ser un número.\n")
                refresh(2)

def Main():
    Actualizacion = Update()
    Icon()
    print(f"    " + Actualizacion)
    print(f"{Style.BRIGHT}\n  - - - Menu De Opciones - - - {Style.NORMAL}")
    print(f"    {Fore.GREEN + Style.DIM}[1] {Fore.WHITE + Style.NORMAL}Activar Windows")
    print(f"    {Fore.GREEN + Style.DIM}[2] {Fore.WHITE + Style.NORMAL}Activar Office")
    print(f"    {Fore.GREEN + Style.DIM}[3] {Fore.WHITE + Style.NORMAL}Salir")

    while True:
        opc = input(Style.BRIGHT + "\n  - Ingrese Su Opcion: " + Style.NORMAL)

        if (opc == "1"):
            refresh(1)
            clear()
            Main_KMSWindows(SO_Version = SO_Windows(), SO_Bits = Bits_Windows())
            break

        elif (opc == "2"):
            refresh(1)
            clear()
            Main_KMSOffice(Of_Exists = Office_Version())
            break

        elif (opc == "3"):
            clear()
            print(Fore.YELLOW + Style.DIM + "[!] Cerrando...")
            refresh(1)
            break
        
        else:
            opc = print(Fore.RED + "[!] Error. Opcion Invalida." + Fore.WHITE)
        

if os.path.exists("Scritps/Start.py"):
    animation = subprocess.run(["python", "Scritps/Start.py"])
    if animation.returncode == 0:
        Main()
    else :
        clear()
        refresh(1)
else:
    Icon()
    print(Fore.RED + "\n[!] Error: No se encontró el archivo de iniciacion.")
    print(Fore.YELLOW + "[!] Integridad de archivos comprometida, Se Entiende como situacion de integridad Fatal")
    print(Fore.GREEN + Style.DIM + "[!] Descarge el programa para solucionarlo.")
    input(Fore.RED + Style.RESET_ALL + "\n \nPresione Enter Para Finalizar...")