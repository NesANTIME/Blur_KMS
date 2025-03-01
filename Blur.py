import sys
import winreg
import os
import subprocess
import time
import json
from colorama import Fore, init, Style
init()

def Icon():
    print(" ")
    print(Fore.CYAN + " ██████╗ ██╗     ██╗   ██╗██████╗               ██╗  ██╗███╗   ███╗███████╗")
    print(" ██╔══██╗██║     ██║   ██║██╔══██╗              ██║ ██╔╝████╗ ████║██╔════╝")
    print(" ██████╔╝██║     ██║   ██║██████╔╝    █████╗    █████╔╝ ██╔████╔██║███████╗")
    print(" ██╔══██╗██║     ██║   ██║██╔══██╗    ╚════╝    ██╔═██╗ ██║╚██╔╝██║╚════██║")
    print(" ██████╔╝███████╗╚██████╔╝██║  ██║              ██║  ██╗██║ ╚═╝ ██║███████║")
    print(" ╚═════╝ ╚══════╝ ╚═════╝ ╚═╝  ╚═╝              ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝")
    print(Fore.BLUE + "   Blur-KMS v1.5 By NesAnTime" + Fore.WHITE)

def clear():
    os.system("cls")

def comandPront(CMD):
    subprocess.run(CMD, shell=True)

def refresh(dat):
    time.sleep(dat)

def searh(dato, mode):
    if mode == 1:
       name = "Scritps/KMS_Codes_Windows.json"
    else:
       name = "Scritps/KMS_Codes_Office.json"

    with open(name, "r", encoding="utf-8") as file:
        kms_data = json.load(file)
        return kms_data.get(dato, "Clave no encontrada")
    
def edit_Sys():
    clear()
    Icon()
    print(Fore.CYAN + "\n[!] Elije La Version De Tu Sistema Operativo: " + Fore.WHITE)

    Versiones = ["Windows 11 Pro", "Windows 10 Pro", "Windows 11 Pro N", "Windows 10 Pro N", "Windows 11 Pro for Workstations", "Windows 10 Pro for Workstations", "Windows 11 Pro for Workstations N", "Windows 10 Pro for Workstations N", "Windows 11 Pro Education", "Windows 10 Pro Education", "Windows 11 Pro Education N", "Windows 10 Pro Education N", "Windows 11 Education", "Windows 10 Education", "Windows 11 Education N", "Windows 10 Education N", "Windows 11 Enterprise", "Windows 10 Enterprise", "Windows 11 Enterprise N", "Windows 10 Enterprise N", "Windows 11 Enterprise G", "Windows 10 Enterprise G", "Windows 11 Enterprise G N", "Windows 10 Enterprise G N", "Windows 11 Enterprise LTSC 2024", "Windows 10 Enterprise LTSC 2021", "Windows 10 Enterprise LTSC 2019", "Windows 11 Enterprise N LTSC 2024", "Windows 10 Enterprise N LTSC 2021", "Windows 10 Enterprise N LTSC 2019", "Windows 10 Home", "Windows 10 Home N", "Windows 10 Home Single Language", "Windows 10 Home Country Specific", "Windows 10 Professional", "Windows 10 Professional N", "Windows 10 Enterprise 2015 LTSB", "Windows 10 Enterprise 2015 LTSB N"]
    print("\n".join([f"{Style.DIM}[{i+1}] {Style.NORMAL}{Versiones[i]}" for i in range(len(Versiones))]))

    opc = input("\nSeleccione una opción: ")

    if opc.isdigit():
        opcion = int(opc)
        if 1 <= opcion <= len(Versiones):
            print(f"\nHas seleccionado: {Versiones[opcion - 1]}")
            return Versiones[opcion - 1]
        else:
            print("\nOpción inválida. Debe estar dentro del rango.")
    else:
        print("\nEntrada no válida. Debe ser un número.")

def version_windows():
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

def arquitectura():
    return "64 bits" if sys.maxsize > 2**32 else "32 bits"


def Main_KMSWindows(V_W, Arh):
    while True:
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
        print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + V_W}")
        print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + Arh} \n")
        lop = input(Fore.YELLOW + Style.DIM + f"[!] Si Considera un Error La Version Del Windows, Escriba (Windows) para corregir." + Fore.YELLOW + Style.DIM + f"\n[!] Si Considera un Error En la Arquitectura, Escriba (Arq) para corregir. \nDe lo contrario Presione la Tecla (Enter)...\n \n {Fore.GREEN}  [¡] Esperando: " + Style.NORMAL + Fore.WHITE)

        if (lop == "windows") or (lop == "Windows") or (lop == "WINDOWS"):
            V_W = edit_Sys()
            clear()

        elif (lop == "arq") or (lop == "Arq") or (lop == "ARQ"):
            print(Fore.GREEN +"\n[!] Su Sistema Operativo: "+ Style.DIM + V_W + Style.NORMAL + " Es de " + Style.DIM + Fore.WHITE + "[1] " + Style.NORMAL + "32 bits  -  " + Style.DIM + "[2] " + Style.NORMAL + "64 bits")
            Mip = int(input(f"  [!] Ingrese Su Opcion: {Style.DIM}"))
            while (Mip < 1) or (Mip > 2):
                print("Opcion Ingresada Invalida.")
                Mip = int(input(f"  [!] Ingrese Nuevamente Su Opcion: {Style.DIM}"))
            if Mip == 1:
                Arh = "32 bits"
            elif Mip == 2:
                Arh = "64 bits" 
        else:
            clear()
            Icon()
            print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
            print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + V_W}")
            print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + Arh} \n")

            print(Fore.GREEN + "[!] Iniciando Programa de Activacion...")
            if searh(dato = V_W ,mode = 1) == "Clave no encontrada":
                print(Fore.RED + "[!] Sistema Operativo Desconocido, Sin Clave KMS Valida.")
            else:
                refresh(2)
                print("\n[!] Iniciando Instalacion de Claves... ")
                comando = f"slmgr /ipk {searh(dato = V_W ,mode = 1)}" 
                comandPront(comando)

                refresh(4)
                print("\n[!] Configurando Administrador de Claves...")
                comandPront("slmgr /skms kms.digiboy.ir")
                print(Style.DIM + " - Administrador de Clave: kms.digiboy.ir" + Style.NORMAL)


                refresh(4)
                print("\n[!] Activando Clave... ")
                comandPront("slmgr /ato")

                refresh(4)
                print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)
                refresh(1)
                input(Fore.RED + "\nPresione Enter Para Finalizar...")
                break

def Main_KMSOffice(V_W, Arh):
    print(Fore.RED + "[!] En Desarrollo")

def Main():
    clear()
    Version_W = version_windows()
    Arh = arquitectura()

    Icon()
    print(f"{Style.BRIGHT}\n- - - Menu De Opciones - - - {Style.NORMAL}")
    print(f"   {Fore.GREEN + Style.DIM}[1] {Fore.WHITE + Style.NORMAL}Activar Windows")
    print(f"   {Fore.GREEN + Style.DIM}[2] {Fore.WHITE + Style.NORMAL}Activar Office")
    print(f"   {Fore.GREEN + Style.DIM}[3] {Fore.WHITE + Style.NORMAL}Salir")

    opc = input(Style.BRIGHT + "\n - Ingrese Su Opcion: " + Style.NORMAL)

    if (opc == "1"):
        clear()
        Main_KMSWindows(Version_W, Arh)

    elif (opc == "2"):
        clear()
        Main_KMSOffice()

    elif (opc == "3"):
        clear()
        print(Fore.YELLOW + Style.DIM + "[!] Cerrando...")
        refresh(1)
    
    else:
        opc = input(Fore.RED + "[!] Error. Opcion Invalida.")

        
    


    

    


    


    



   



Main()
