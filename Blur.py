import platform
import sys
import winreg
import os
import subprocess
import time
import json

def Icon():
    print("██████╗ ██╗     ██╗   ██╗██████╗               ██╗  ██╗███╗   ███╗███████╗")
    print("██╔══██╗██║     ██║   ██║██╔══██╗              ██║ ██╔╝████╗ ████║██╔════╝")
    print("██████╔╝██║     ██║   ██║██████╔╝    █████╗    █████╔╝ ██╔████╔██║███████╗")
    print("██╔══██╗██║     ██║   ██║██╔══██╗    ╚════╝    ██╔═██╗ ██║╚██╔╝██║╚════██║")
    print("██████╔╝███████╗╚██████╔╝██║  ██║              ██║  ██╗██║ ╚═╝ ██║███████║")
    print("╚═════╝ ╚══════╝ ╚═════╝ ╚═╝  ╚═╝              ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝")
    print("Blur-KMS v1.0 By NesAnTime")

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
    Icon()

    print("\n [!] Información del Sistema:")
    print(f"      - Sistema Operativo: {V_W}")
    print(f"      - Arquitectura: {Arh} \n")
    lop = input("Si hay Algun Error En La Version Del Windows o Arquitectura, Escriba (edit) para corregir. O Presione Enter para Continuar.. ")

    if (lop == "edit") or (lop == "Edit") or (lop == "EDIT"):
        #V_W = input("  [!] Ingrese su Sistema Operativo (Ejem: Windows 10 Pro): ")

        print("Su SO: " + V_W + " Es de [1] 32 bits  -  [2] 64 bits")
        Mip = int(input("  [!] Ingrese Su Opcion: "))
        while (Mip < 1) or (Mip > 2):
            print("Opcion Ingresada Invalida.")
            Mip = int(input("  [!] Ingrese Nuevamente Su Opcion: "))

        if Mip == 1:
            Arh = "32 bits"
        elif Mip == 2:
            Arh = "64 bits" 
    
    clear()
    Icon()
    print("\n [!] Información del Sistema:")
    print(f"      - Sistema Operativo: {V_W}")
    print(f"      - Arquitectura: {Arh} \n")

    print("[!] Iniciando Programa de Activacion...")
    
    refresh(2)
    print("[!] Iniciando Instalacion de Claves ")
    comando = f"slmgr /ipk {searh(dato = V_W ,mode = 1)}" 
    comandPront(comando)

    refresh(4)
    print("[!] Configurando Administrador de Claves ")
    comandPront("slmgr /skms kms.digiboy.ir")

    refresh(4)
    print("[!] Activando Clave ")
    comandPront("slmgr /ato")

    refresh(4)
    print("[!] Finalizando...")
    refresh(1)
    input("\nPresione Enter Para Finalizar...")

def Main_KMSOffice(V_W, Arh):
    return

def Main():
    Version_W = version_windows()
    Arh = arquitectura()

    Icon()
    print("- - - Menu De Opciones - - - ")
    print("[1] Activar Windows")
    print("[2] Activar Office")
    print("[3] Salir")

    opc = input(" - Ingrese Su Opcion: ")

    if (opc == "1"):
        clear()
        Main_KMSWindows(Version_W, Arh)

    elif (opc == "2"):
        clear()
        Main_KMSOffice()

    elif (opc == "3"):
        clear()
        print("Cerrando...")
        refresh(1)
    
    else:
        opc = input("Error. Opcion Invalida.")

        
    


    

    


    


    



   



Main()
