import sys
import winreg
import os
import subprocess
import time
import winreg
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
    print(Fore.BLUE + "   Blur-KMS v2.0 By NesAnTime" + Fore.WHITE)

def clear():
    os.system("cls")

def Command_Prompt(CMD):
    subprocess.run(CMD, shell=True)

def refresh(dat):
    time.sleep(dat)

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
    Codes = [
        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",  # Office Click-to-Run
        r"SOFTWARE\Microsoft\Office\16.0\Common\ProductVersion",  # Office 2016, 2019, 2021 y 365
        r"SOFTWARE\Microsoft\Office\15.0\Common\ProductVersion",  # Office 2013
        r"SOFTWARE\Microsoft\Office\14.0\Common\ProductVersion",  # Office 2010
    ]
    
    print("Buscando Microsoft Office en el registro de Windows...")
    
    for clave in Codes:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, clave):
                return True
        except FileNotFoundError:
            continue
    
    return None

def Search(dato):
    with open("Scritps/KMS_Codes_Windows.json", "r", encoding="utf-8") as file:
        kms_data = json.load(file)
        return kms_data.get(dato, "Clave no encontrada")
    
def Internet_Conexion():
    try:
        subprocess.check_call(["ping", "-n", "1", "8.8.8.8"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False


def Funcion_ActivationSO(SO_Version, KMS_Code):
    print("\n[!] Iniciando Instalacion de Claves... ")
    Command_Prompt(CMD = f"slmgr /ipk {KMS_Code}")

    refresh(4)
    print("\n[!] Configurando Administrador de Claves...")
    Command_Prompt(CMD = "slmgr /skms kms.digiboy.ir")
    print(Style.DIM + " - Administrador de Claves: kms.digiboy.ir" + Style.NORMAL)

    refresh(4)
    print("\n[!] Activando Clave... ")
    Command_Prompt(CMD = "slmgr /ato")

    refresh(4)
    print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)
    print(Fore.GREEN + Style.DIM + f"[!] Activacion Completada. {SO_Version} Activado Con Exito.")
    print(Fore.GREEN + f"\n[!] Es Nesesario Reiniciar, Para Completar. (Puede Reiniciar Mas Tarde)" + Style.NORMAL)
    refresh(1)
    input(Fore.RED + "\nPresione Enter Para Finalizar...")

def Funcion_ActivationOffice(Office):
    print("\n[!] Iniciando Instalacion de Claves KMS En Office... ")
    Command_Prompt(CMD = f"cd /d %ProgramFiles(x86)%\Microsoft Office\Office16")
    Command_Prompt(CMD = "cd /d %ProgramFiles%\Microsoft Office\Office16")

    refresh(2)
    print("\n[!] Convirtiendo Office a Licencia Volumen... ")
    if Office == "Microsoft Office 2016":
        Command_Prompt(CMD = "for /f %x in ('dir /b ..\\root\\Licenses16\\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:\"..\\root\\Licenses16\\%x\"")
    elif Office == "Microsoft Office 2019":
        Command_Prompt(CMD = "for /f %x in ('dir /b ..\\root\\Licenses16\\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:\"..\\root\\Licenses16\\%x\"")
    elif Office == "Microsoft Office 2021":
        Command_Prompt(CMD = "for /f %x in ('dir /b ..\\root\\Licenses16\\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:\"..\\root\\Licenses16\\%x\"")
    print(Style.DIM + " - Convertido Exitosamente..." + Style.NORMAL)

    refresh(2)
    print("\n[!] Activando Clave... ")
    if Office == "Microsoft Office 2016":
        Command_Prompt(CMD = "cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99")
        Command_Prompt(CMD = "cscript ospp.vbs /unpkey:BTDRB >nul")
        Command_Prompt(CMD = "cscript ospp.vbs /unpkey:KHGM9 >nul")
        Command_Prompt(CMD = "cscript ospp.vbs /unpkey:CPQVG >nul")
        Command_Prompt(CMD = "cscript ospp.vbs /sethst:e8.us.to")
        Command_Prompt(CMD = "cscript ospp.vbs /setprt:1688")
        Command_Prompt(CMD = "cscript ospp.vbs /act")
    elif Office == "Microsoft Office 2019":
        Command_Prompt(CMD = "cscript ospp.vbs /setprt:1688")
        Command_Prompt(CMD = "cscript ospp.vbs /unpkey:6MWKP >nul")
        Command_Prompt(CMD = "cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP")
        Command_Prompt(CMD = "cscript ospp.vbs /sethst:e8.us.to")
        Command_Prompt(CMD = "cscript ospp.vbs /act")
    elif Office == "Microsoft Office 2021":
        Command_Prompt(CMD = "cscript ospp.vbs /setprt:1688")
        Command_Prompt(CMD = "cscript ospp.vbs /unpkey:6F7TH >nul")
        Command_Prompt(CMD = "cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH")
        Command_Prompt(CMD = "cscript ospp.vbs /sethst:e8.us.to")
        Command_Prompt(CMD = "cscript ospp.vbs /act")

    refresh(2)
    print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)
    print(Fore.GREEN + Style.DIM + f"[!] Activacion Completada. Activado Con Exito.")
    refresh(1)
    input(Fore.RED + "\nPresione Enter Para Finalizar...")
    
def SO_Edit():
    clear()
    Icon()
    print(Fore.CYAN + "\n[!] Elije La Version De Tu Sistema Operativo: " + Fore.WHITE)

    Versiones = ["Windows 11 Pro", "Windows 10 Pro", "Windows 11 Pro N", "Windows 10 Pro N", "Windows 11 Pro for Workstations", "Windows 10 Pro for Workstations", "Windows 11 Pro for Workstations N", "Windows 10 Pro for Workstations N", "Windows 11 Pro Education", "Windows 10 Pro Education", "Windows 11 Pro Education N", "Windows 10 Pro Education N", "Windows 11 Education", "Windows 10 Education", "Windows 11 Education N", "Windows 10 Education N", "Windows 11 Enterprise", "Windows 10 Enterprise", "Windows 11 Enterprise N", "Windows 10 Enterprise N", "Windows 11 Enterprise G", "Windows 10 Enterprise G", "Windows 11 Enterprise G N", "Windows 10 Enterprise G N", "Windows 11 Enterprise LTSC 2024", "Windows 10 Enterprise LTSC 2021", "Windows 10 Enterprise LTSC 2019", "Windows 11 Enterprise N LTSC 2024", "Windows 10 Enterprise N LTSC 2021", "Windows 10 Enterprise N LTSC 2019", "Windows 10 Home", "Windows 10 Home N", "Windows 10 Home Single Language", "Windows 10 Home Country Specific", "Windows 10 Professional", "Windows 10 Professional N", "Windows 10 Enterprise 2015 LTSB", "Windows 10 Enterprise 2015 LTSB N"]
    print("\n".join([f"{Style.DIM}[{i+1}] {Style.NORMAL}{Versiones[i]}" for i in range(len(Versiones))]))

    try:
        opc = input("\nSeleccione una opción: ")

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
        print(Fore.RED + "[!] Error. Entrada no válida.")
        refresh(2)


def Main_KMSWindows(SO_Version, SO_Bits):
    while True:
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
        print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + SO_Version}")
        print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + SO_Bits} \n")
        opc = input(Fore.YELLOW + Style.DIM + f"[!] Si Considera un Error La Version Del Windows, Escriba (Windows) para corregir." + Fore.YELLOW + Style.DIM + f"\n[!] Si Considera un Error En la Arquitectura, Escriba (Bits) para corregir. \nDe lo contrario Presione la Tecla (Enter)...\n \n {Fore.GREEN}  [¡] Esperando: " + Style.NORMAL + Fore.WHITE)

        if (opc == "windows") or (opc == "Windows") or (opc == "WINDOWS"):
            SO_Version = SO_Edit()
            clear()

        elif (opc == "bits") or (opc == "Bits") or (opc == "BITS"):
            print(Fore.GREEN +"\n[!] Su Sistema Operativo: "+ Style.DIM + SO_Version + Style.NORMAL + " Es de " + Style.DIM + Fore.WHITE + "[1] " + Style.NORMAL + "32 bits  -  " + Style.DIM + "[2] " + Style.NORMAL + "64 bits")
            try:
                Mip = int(input(f"  [!] Ingrese Su Opcion: {Style.DIM}"))
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
        else:
            clear()
            Icon()
            print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
            print(f"      {Style.BRIGHT + Fore.CYAN}- Sistema Operativo: {Style.NORMAL + SO_Version}")
            print(f"      {Style.BRIGHT + Fore.CYAN}- Arquitectura: {Style.NORMAL + SO_Bits} \n")

            print(Fore.GREEN + "[!] Iniciando Programa de Activacion...")

            KMS_Code = Search(dato = SO_Version, mode = True)

            if KMS_Code == "Clave no encontrada":
                print(Fore.RED + "[!] Sistema Operativo Desconocido, Sin Clave KMS Valida.")
            else:
                refresh(2)
                Funcion_ActivationSO(SO_Version, KMS_Code)
                break

def Main_KMSOffice(Of_Exists):
    while True:
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [¡] Buscando Microsoft Office en el registro de Windows...\n" + Style.NORMAL)
        refresh(2)

        if Of_Exists == None:
            print(f"      {Style.BRIGHT + Fore.RED}- [!] Microsoft Office no está instalado en este sistema. {Style.NORMAL}")
        else: 
            print(f"{Style.BRIGHT + Fore.GREEN}  [!] Microsoft Office Detectado \n")
            print(Fore.CYAN + "***    Menu de Versiones de Office   ***\n" + Fore.WHITE)
            print(f"{Style.DIM}     [1]{Style.NORMAL} Microsoft Office 2016")
            print(f"{Style.DIM}     [2]{Style.NORMAL} Microsoft Office 2019")
            print(f"{Style.DIM}     [3]{Style.NORMAL} Microsoft Office 2021")
            print(f"{Style.DIM}     [4]{Style.NORMAL} Claves Genericas Para Office\n")

            try:
                opc = int(input(Fore.YELLOW + Style.DIM + f"[¡] Elija la Opcion Requerida: "+ Style.NORMAL + Fore.WHITE))
                if (opc == 1):
                    Of_Version = "Microsoft Office 2016"
                elif (opc == 2):
                    Of_Version = "Microsoft Office 2019"
                elif (opc == 3):
                    Of_Version = "Microsoft Office 2021"
                elif (opc == 4):
                    print(Fore.CYAN + "[!] Abriendo en el Navegador")
                    refresh(2)
                    os.system("start Scritps/Codes_Office.html")
                    break 
                else:
                    print(Fore.RED + "[!] Opcion Invalida.")
                    refresh(2)
            except ValueError:
                print(Fore.RED + "[!] Entrada no válida. Debe ser un número.")
                refresh(2)
                
            
            clear()
            Icon()
            print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
            print(f"      {Style.BRIGHT + Fore.GREEN}[!]  {Of_Version + Style.NORMAL} \n")

            print(Fore.GREEN + "[!] Iniciando Programa de Activacion KMS...\n")
            print(Fore.YELLOW + Style.DIM +  "[!] NOTA: La Herramienta de Claves KMS para Office tiene un 50% de posibilidades de \nresultar exitosa. Si tiene Un resultado Desfaborable Se Pueden le Proporcionar \nClaves Genericas Para Office." + Fore.WHITE + Style.NORMAL + " \n \n Presione Enter para Continuar...")
            input()

            while True:
                if Internet_Conexion() == True:
                    print(Fore.GREEN + "[!] Conectado a Internet." + Fore.WHITE)
                    break
                else:
                    print(Fore.RED + "[!] No se ha detectado conexión a Internet." + Fore.WHITE)
                    print("Intentelo Nuevamente Conectado a Internet, Se Intentara Conectar Automaticamente.")
                    refresh(2)
            
            Funcion_ActivationOffice(Of_Version)
            break




def Main():
    clear()

    Icon()
    print(f"{Style.BRIGHT}\n- - - Menu De Opciones - - - {Style.NORMAL}")
    print(f"   {Fore.GREEN + Style.DIM}[1] {Fore.WHITE + Style.NORMAL}Activar Windows")
    print(f"   {Fore.GREEN + Style.DIM}[2] {Fore.WHITE + Style.NORMAL}Activar Office")
    print(f"   {Fore.GREEN + Style.DIM}[3] {Fore.WHITE + Style.NORMAL}Salir")

    while True:
        opc = input(Style.BRIGHT + "\n - Ingrese Su Opcion: " + Style.NORMAL)

        if (opc == "1"):
            SO_Version = SO_Windows()
            SO_Bits = Bits_Windows()
            clear()
            Main_KMSWindows(SO_Version, SO_Bits)
            break

        elif (opc == "2"):
            Of_Exists = Office_Version()
            clear()
            Main_KMSOffice(Of_Exists)
            break

        elif (opc == "3"):
            clear()
            print(Fore.YELLOW + Style.DIM + "[!] Cerrando...")
            refresh(1)
            break
        
        else:
            opc = input(Fore.RED + "[!] Error. Opcion Invalida.")
        

Main()