import sys
import winreg
import os
import subprocess
import time
import winreg
import urllib.request
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
    print(Fore.BLUE + "   Blur-KMS v2.3 By NesAnTime" + Fore.WHITE)

def clear():
    os.system("cls")

def Rut_OSPP():
    rutas= [r"C:\Program Files\Microsoft Office\Office16\ospp.vbs", r"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"]
    for ruta in rutas:
        if os.path.exists(ruta):
            return ruta
    return None

def Command(CMD, mode):
    if mode == "Prompt":
        subprocess.run(CMD, shell=True)
    elif mode == "Shell":
        subprocess.run(CMD, shell=True, capture_output=True, text=True)
    else:
        Ospp_rut = Rut_OSPP()
        subprocess.run(f'cscript "{Ospp_rut}" {CMD}', shell=True, check=True)

def refresh(dat):
    time.sleep(dat)

def Bits_Windows():
    return "64 bits" if sys.maxsize > 2**32 else "32 bits"

#Updater--------
def Update():        
    VerLocal = 'Scritps/Version.txt'
    VerNew = "https://raw.githubusercontent.com/NesANTIME/Blur_KMS/refs/heads/main/Scritps/Version.txt"
    
    def V_NewVer(VerNew):
        try:
            with urllib.request.urlopen(VerNew) as response:
                contenido = response.read().decode('utf-8').strip()
                return contenido
        except Exception as e:
            print(f"{Fore.RED}[!] No se pudo obtener el contenido remoto: {e}")
            return None
        
    def V_LocalVer(VerLocal, V_NewUp):
        try:
            if not os.path.exists(VerLocal):
                print(f"{Fore.RED}[!] El archivo local {VerLocal} no existe.")
                return
            
            with open(VerLocal, 'r') as file:
                contenido_local = file.read().strip()
                if contenido_local == V_NewUp:
                    return f"{Fore.WHITE + Style.BRIGHT}[!] No hay actualizaciones disponibles."
                else:
                    return f"{Fore.CYAN + Style.BRIGHT}[!] La Nueva Version de {V_NewUp} Esta disponible ¡Descargala YA!\n"
        except Exception as e:
            print(f"{Fore.RED}[!] Error al leer el archivo local {VerLocal}: {e}")

    V_NewUp = V_NewVer(VerNew)
    if V_NewUp is not None:
        resultado = V_LocalVer(VerLocal, V_NewUp)
        return resultado
    else:
        return Fore.RED + "[!] No se pudo obtener el contenido remoto para comparar."
#---------------

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

def Office_Ruta():
    rut = [
        r"%ProgramFiles(x86)%\Microsoft Office\Office16",
        r"%ProgramFiles%\Microsoft Office\Office16"
    ]

    for path in rut:
        if os.path.exists(os.path.expandvars(path)):
            Command(CMD = f"cd /d {path}", mode = "Prompt")
            break
    else:
        print(Fore.RED + "[!] No se encontró la ruta de instalación de Microsoft Office." + Fore.WHITE)
        return

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
    Command(CMD = f"slmgr /ipk {KMS_Code}", mode = "Prompt")

    refresh(4)
    print("\n[!] Configurando Administrador de Claves...")
    Command(CMD = "slmgr /skms kms.digiboy.ir", mode = "Prompt")
    print(Style.DIM + " - Administrador de Claves: kms.digiboy.ir" + Style.NORMAL)

    refresh(4)
    print("\n[!] Activando Clave... ")
    Command(CMD = "slmgr /ato", mode = "Prompt")

    refresh(4)
    print(Style.BRIGHT + "\n[!] Finalizando..." + Style.NORMAL)
    print(Fore.GREEN + Style.DIM + f"[!] Activacion Completada. {SO_Version} Activado Con Exito.")
    print(Fore.GREEN + f"\n[!] Es Nesesario Reiniciar, Para Completar. (Puede Reiniciar Mas Tarde)" + Style.NORMAL)
    refresh(1)
    input(Fore.RED + "\nPresione Enter Para Finalizar...")

def Funcion_ActivationOffice(Office):
    print(Fore.GREEN + "\n[!] Iniciando Instalacion de Claves KMS En Office... " + Fore.WHITE)
    Office_Ruta()

    refresh(2)
    print(Fore.GREEN + "\n[!] Convirtiendo Office a Licencia Volumen... " + Fore.WHITE)
    if Office == "Microsoft Office 2016":
        Command(r'for /f %x in (\'dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms\') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"', mode = "Shell")
    elif Office == "Microsoft Office 2019":
        Command(r'for /f %x in (\'dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms\') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"', mode = "Shell")
    elif Office == "Microsoft Office 2021":
        Command(r'for /f %x in (\'dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms\') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"', mode = "Shell")
    print(Style.DIM + " - Convertido Exitosamente..." + Style.NORMAL)

    Ospp_rut = Rut_OSPP()

    refresh(2)
    print(Fore.GREEN + "\n[!] Activando Clave... " + Fore.WHITE)

    if not Ospp_rut:
        print(Fore.RED + "[ERROR] No se encontró ospp.vbs en las rutas esperadas." + Fore.WHITE)
    else:
        if Office == "Microsoft Office 2016":
            Command(CMD=" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99", mode="")
            Command(CMD=" /unpkey:BTDRB >nul", mode="")
            Command(CMD=" /unpkey:KHGM9 >nul", mode="")
            Command(CMD=" /unpkey:CPQVG >nul", mode="")
            Command(CMD=" /sethst:e8.us.to", mode="")
            Command(CMD=" /setprt:1688", mode="")
            Command(CMD=" /act", mode="")
        elif Office == "Microsoft Office 2019":
            Command(CMD=" /setprt:1688", mode="")
            Command(CMD=" /unpkey:6MWKP >nul", mode="")
            Command(CMD=" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", mode="")
            Command(CMD=" /sethst:e8.us.to", mode="")
            Command(CMD=" /act", mode="")
        elif Office == "Microsoft Office 2021":
            Command(CMD=" /setprt:1688", mode="")
            Command(CMD=" /unpkey:6F7TH >nul", mode="")
            Command(CMD=" /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH", mode="")
            Command(CMD=" /sethst:e8.us.to", mode="")
            Command(CMD=" /act", mode="")
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

            KMS_Code = Search(dato = SO_Version)

            if KMS_Code == "Clave no encontrada":
                print(Fore.RED + "[!] Sistema Operativo Desconocido, Sin Clave KMS Valida.")
            else:
                refresh(2)
                Funcion_ActivationSO(SO_Version, KMS_Code)
                break

def Main_KMSOffice(Of_Exists):
    def Execute(Of_Version):
        clear()
        Icon()
        print(Fore.YELLOW + Style.DIM + "\n [!] Información Obtenida del Sistema:" + Style.NORMAL)
        print(f"      {Style.BRIGHT + Fore.GREEN}[!]  {Of_Version + Style.NORMAL} \n")

        print(Fore.GREEN + "[!] Iniciando Programa de Activacion KMS...\n")
        print(Fore.YELLOW + Style.DIM +  "[!] NOTA: La Herramienta de Claves KMS para Office tiene un 50% de posibilidades de \nresultar exitosa. Si tiene Un resultado Desfaborable Se Pueden le Proporcionar \nClaves Genericas Para Office." + Fore.WHITE + Style.NORMAL + " \n \n Presione Enter para Continuar...")
        input()

        while True:
            if Internet_Conexion() == True:
                print(Fore.GREEN + " ----- [!] Conectado a Internet." + Fore.WHITE)
                break
            else:
                print(Fore.RED + " ----- [!] No se ha detectado conexión a Internet." + Fore.WHITE)
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
        print(Fore.CYAN + "***    Menu de Versiones de Office   ***\n" + Fore.WHITE)
        print(f"{Style.DIM}     [1]{Style.NORMAL} Microsoft Office 2016")
        print(f"{Style.DIM}     [2]{Style.NORMAL} Microsoft Office 2019")
        print(f"{Style.DIM}     [3]{Style.NORMAL} Microsoft Office 2021")
        print(f"{Style.DIM}     [4]{Style.NORMAL} Claves Genericas Para Office\n")

        while True:
            try:
                opc = int(input(Fore.YELLOW + Style.DIM + f"[¡] Elija la Opcion Requerida: "+ Style.NORMAL + Fore.WHITE))
                if (opc == 1):
                    Execute(Of_Version = "Microsoft Office 2016")
                    break
                elif (opc == 2):
                    Execute(Of_Version = "Microsoft Office 2019")
                    break
                elif (opc == 3):
                    Execute(Of_Version = "Microsoft Office 2021")
                    break
                elif (opc == 4):
                    print(Fore.CYAN + "[!] Abriendo en el Navegador")
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
    clear()

    Icon()
    print(f"{Style.BRIGHT}\n- - - Menu De Opciones - - - {Style.NORMAL}")
    print(f"   {Fore.GREEN + Style.DIM}[1] {Fore.WHITE + Style.NORMAL}Activar Windows")
    print(f"   {Fore.GREEN + Style.DIM}[2] {Fore.WHITE + Style.NORMAL}Activar Office")
    print(f"   {Fore.GREEN + Style.DIM}[3] {Fore.WHITE + Style.NORMAL}Buscar Actualizaciones")
    print(f"   {Fore.GREEN + Style.DIM}[4] {Fore.WHITE + Style.NORMAL}Salir")

    while True:
        opc = input(Style.BRIGHT + "\n - Ingrese Su Opcion: " + Style.NORMAL)

        if (opc == "1"):
            SO_Version = SO_Windows()
            SO_Bits = Bits_Windows()
            refresh(1)
            clear()
            Main_KMSWindows(SO_Version, SO_Bits)
            break

        elif (opc == "2"):
            Of_Exists = Office_Version()
            refresh(1)
            clear()
            Main_KMSOffice(Of_Exists)
            break

        elif (opc == "3"):
            resultado = Update()
            if resultado:
                refresh(1)
                clear()
                Icon()
                print("\n" + resultado)
            input(Fore.WHITE + Style.NORMAL + "\n[!] Presione Enter Para Continuar...")
            Main()

        elif (opc == "4"):
            clear()
            print(Fore.YELLOW + Style.DIM + "[!] Cerrando...")
            refresh(1)
            break
        
        else:
            opc = print(Fore.RED + "[!] Error. Opcion Invalida." + Fore.WHITE)
        

Main()