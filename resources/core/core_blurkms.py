import os
import sys
import json
import time
import socket
import ctypes
import msvcrt
import webbrowser
from colorama import Fore, init, Style
init()

from tools.windows_active import SO_Windows, Bits_Windows, buscador_clave_kms, activador_kms_windows_funcion
from tools.office_active import detect_office, Funcion_ActivationOffice

def load_json_blurkms(dat):
    ruta_base = os.path.dirname(__file__)
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


def icon():
    icon = load_coredata()
    version = load_configkms()
    clear()
    print(Fore.CYAN + "\n")
    for i in icon["icon"]:
        print(i)
    print(f"{" "*76} {Fore.WHITE + Style.BRIGHT}[{version["BlurKMS"][0]}]{Style.DIM} By NesAnTime{Style.RESET_ALL}\n")



def internet_conexion(hosts=["8.8.8.8", "kms.digiboy.ir", "1.1.1.1"], port=53, timeout=3):
    for host in hosts:
        try:
            socket.setdefaulttimeout(timeout)
            socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
            return True
        except socket.error:
            continue
    return False




def initiator_blurkms():
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hWnd = kernel32.GetConsoleWindow()
    pantalla_ancho = user32.GetSystemMetrics(0)
    pantalla_alto = user32.GetSystemMetrics(1)

    ancho_ventana = 800
    alto_ventana = 500
    x = int((pantalla_ancho - ancho_ventana) / 2)
    y = int((pantalla_alto - alto_ventana) / 2)
    ctypes.windll.user32.MoveWindow(hWnd, x, y, ancho_ventana, alto_ventana, True)
    Main()


        
def menu_windows(version_win, arq_win):
    icon()

    def version_manual_win():
        icon()
        datos = load_coredata()
        print(f"{" "*5}⚠️  Editando version de windows\n")
        print(f"{" "*8}{Fore.BLUE}[!] {Fore.CYAN}Elije la version de tu sistema operativo: " + Fore.WHITE)

        print("\n".join([f"{" "*11}{Style.DIM}[{i+1}] {Style.NORMAL}{datos["SO_EditCache"][i]}" for i in range(len(datos["SO_EditCache"]))]))

        while True:
            try:
                opc = input(f"\n{" "*10} [!] {Style.BRIGHT}Seleccione una opcion: {Style.RESET_ALL}".strip())

                if opc.isdigit():
                    opcion = int(opc)
                    if 1 <= opcion <= len(datos["SO_EditCache"]):
                        print(f"\n{" "*11}{Fore.GREEN}[!] {Fore.WHITE}Has seleccionado: {datos["SO_EditCache"][opcion - 1]}")
                        duration(1)
                        return datos["SO_EditCache"][opcion - 1]
                    else:
                        print(f"\n{" "*11}{Fore.RED}[!] Opción inválida. Debe estar dentro del rango.{Fore.WHITE}")
                else:
                    print(f"{" "*11}{Fore.RED}\n[!] Entrada no válida. Debe ser un número.{Fore.WHITE}")

            except Exception as e:
                print(Fore.RED + f"[!] Error inesperado: {e}" + Fore.WHITE)
                duration(2)

    def version_manual_arq():
        icon()
        print(f"{" "*5}⚠️  Editando arquitectura de windows\n")
        print(f"{" "*8}{Fore.BLUE}[!] {Fore.CYAN}Elije la version de la arquitectira del sistema operativo: " + Fore.WHITE)

        print(f"\n{" "*11}{Style.DIM}[1] {Style.NORMAL} 32 bits\n{" "*11}{Style.DIM}[2] {Style.NORMAL} 64 bits")

        while True:
            try:
                opc = input(f"\n{" "*10} [!] {Style.BRIGHT}Seleccione una opcion: {Style.RESET_ALL}".strip())

                if (opc == "1"):
                    print(f"\n{" "*11}{Fore.GREEN}[!] {Fore.WHITE}Has seleccionado: {Fore.CYAN} 32 bits")
                    duration(1)
                    return "32 bits"
                elif (opc == "2"):
                    print(f"\n{" "*11}{Fore.GREEN}[!] {Fore.WHITE}Has seleccionado: {Fore.CYAN} 64 bits")
                    duration(1)
                    return "64 bits"
                else:
                    print(f"{" "*11}{Fore.RED}\n[!] Entrada no válida. Debe ser un número.{Fore.WHITE}")
            except Exception as e:
                print(Fore.RED + f"[!] Error inesperado: {e}" + Fore.WHITE)
                duration(2)



    def start_win_kms(version_win, arq_win):
        icon()
        spinner = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]

        print(f"{Fore.YELLOW + Style.DIM}{" "*4}╔{"═"*24}{Style.NORMAL} [⚠️ ] {Style.BRIGHT}Información del Sistema{Style.NORMAL} [⚠️ ] {Style.DIM}{"═"*25}╗\n{" "*4}{Fore.YELLOW}║{" "*86}║{Style.RESET_ALL}")
        print(f"{Fore.YELLOW + Style.DIM}{" "*4}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.CYAN}💻 Sistema Operativo: {Style.NORMAL}{version_win}")
        print(f"{Fore.YELLOW + Style.DIM}{" "*4}║{" "*4}{Style.RESET_ALL}{Style.DIM + Fore.CYAN}📲 Arquitectura: {Style.NORMAL}{arq_win}\n{" "*4}{Fore.YELLOW + Style.DIM}║{" "*86}║")
        print(f"{" "*4}{Style.DIM}╚{"═"*86}╝{Style.RESET_ALL}\n")

        while True:
            print(f"{" "*8}🛜  Comprobando conexión a internet ", end="", flush=True)

            for _ in range(5):
                for frame in spinner:
                    sys.stdout.write(f"\r{" "*8}🛜  Comprobando conexión a internet {frame}")
                    sys.stdout.flush()
                    duration(0.1)

            if internet_conexion():
                print(f"\r{" "*10}🛜  {Fore.GREEN}Conexión establecida ✅   ")
                break
            else:
                print(f"\r{" "*10}🛜  {Fore.RED}Sin Conexión a internet ❌   {Fore.RESET}\n{" "*14}Intententandolo nuevamente\n")
                duration(2)

        print(f"\n{" "*6}{Fore.YELLOW}[!] {Fore.WHITE + Style.BRIGHT}Iniciando Programa de Activacion {Fore.YELLOW + Style.NORMAL}[!] {Style.RESET_ALL}")
        
        clave_kms = buscador_clave_kms(version_win)
        if (clave_kms != None):
            duration(2)
            icon()
            activador_kms_windows_funcion(version_win, clave_kms)
        else:
            print(f"{" "*8}{Fore.RED + Style.BRIGHT}⚠️  {Style.NORMAL}No se encontro la clave KMS valida para su version, es posible que el programa este modificado.")
            sys.exit(1)


    print(f"{Fore.YELLOW + Style.DIM}{" "*4}╔{"═"*20}{Style.NORMAL} [⚠️ ] {Style.BRIGHT}Información Obtenida del Sistema{Style.NORMAL} [⚠️ ] {Style.DIM}{"═"*20}╗\n{" "*4}{Fore.YELLOW}║{" "*86}║{Style.RESET_ALL}")
    print(f"{Fore.YELLOW + Style.DIM}{" "*4}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.CYAN}💻 Sistema Operativo: {Style.NORMAL}{version_win}")
    print(f"{Fore.YELLOW + Style.DIM}{" "*4}║{" "*4}{Style.RESET_ALL}{Style.DIM + Fore.CYAN}📲 Arquitectura: {Style.NORMAL}{arq_win}\n{" "*4}{Fore.YELLOW + Style.DIM}║{" "*86}║")
    print(f"{" "*4}{Style.DIM}╚{"═"*86}╝{Style.RESET_ALL}\n")

    print(f"{" "*9}{Style.BRIGHT}[!] {Style.NORMAL}Si Considera un error la version o arquitectura obtenida del sistema.\n{" "*12}Elija la opcion de su preferencia: \n")
    print(f"{" "*13}{Style.DIM}[1]{Style.NORMAL} 💱 Cambiar Version de Windows")
    print(f"{" "*13}{Style.DIM}[2]{Style.NORMAL} 💱 Cambiar Arquitectura de Windows")
    print(f"{" "*13}{Style.DIM}[3]{Style.NORMAL} 🚀 Continuar Con Activacion de Windows")

    while True:
        print(f"\n{Style.BRIGHT  + Fore.YELLOW}{" "*14}> {Style.NORMAL}[OPCION]:{Style.RESET_ALL} ", end="", flush=True)
        opc = ""
        while True:
            char = msvcrt.getwch()
            if char == "\r":
                print()
                break
            elif char == "\b":
                if opc:
                    opc = opc[:-1]
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
                continue
            opc += char

            if (char == "1") or (char == "2") or (char == "3"):
                sys.stdout.write(Fore.GREEN + Style.DIM + char + Style.RESET_ALL)
            else:
                sys.stdout.write(Fore.RED + char + Style.RESET_ALL)
            sys.stdout.flush()
        duration(0.5)
        if (opc == "1") or (opc == ""):
            menu_windows(version_manual_win(), arq_win)
        elif (opc == "2"):
            menu_windows(version_win, version_manual_arq())
        elif (opc == "3"):
            start_win_kms(version_win, arq_win)
            break
        else:
            print(f"{" "*14}{Fore.RED + Style.BRIGHT}[!] Error. Opcion Invalida.{Style.RESET_ALL}")







def menu_office(detect_office):
    icon()

    def start_office(office):
        icon()
        spinner = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]

        print(f"{Fore.YELLOW + Style.DIM}{" "*4}╔{"═"*24}{Style.NORMAL} [⚠️ ] {Style.BRIGHT}Información de office{Style.NORMAL} [⚠️ ] {Style.DIM}{"═"*25}╗\n{" "*4}{Fore.YELLOW}║{" "*84}║{Style.RESET_ALL}")
        print(f"{Fore.YELLOW + Style.DIM}{" "*4}║{" "*4}{Style.RESET_ALL}{Style.BRIGHT + Fore.CYAN}💻 Version de Office a Lotes: {Style.NORMAL}{office}\n{" "*4}{Fore.YELLOW + Style.DIM}║{" "*84}║{Style.RESET_ALL}")
        print(f"{" "*4}{Fore.YELLOW + Style.DIM}╚{"═"*84}╝{Style.RESET_ALL}\n")

        while True:
            print(f"{" "*8}🛜  Comprobando conexión a internet ", end="", flush=True)

            for _ in range(5):
                for frame in spinner:
                    sys.stdout.write(f"\r{" "*8}🛜  Comprobando conexión a internet {frame}")
                    sys.stdout.flush()
                    duration(0.1)

            if internet_conexion():
                print(f"\r{" "*10}🛜  {Fore.GREEN}Conexión establecida ✅   ")
                break
            else:
                print(f"\r{" "*10}🛜  {Fore.RED}Sin Conexión a internet ❌   {Fore.RESET}\n{" "*14}Intententandolo nuevamente\n")
                duration(2)

        print(f"\n{" "*6}{Fore.YELLOW}[!] {Fore.WHITE + Style.BRIGHT}Iniciando Programa de Activacion {Fore.YELLOW + Style.NORMAL}[!] {Style.RESET_ALL}")
        duration(2)
        Funcion_ActivationOffice(office)


    
    if (detect_office != True):
        print(f"{" "*4}{Fore.YELLOW + Style.NORMAL} [⚠️ ] {Style.BRIGHT}Información del Sistema (Office NO detectado){Style.NORMAL} [⚠️ ]\n")
        print(f"{" "*4} Cerrando...")
        sys.exit(1)

    else: 
        print(f"{" "*4}{Fore.YELLOW + Style.NORMAL} [⚠️ ] {Style.BRIGHT}Información del Sistema (Office detectado){Style.RESET_ALL} [⚠️ ]\n")

        print(f"{Style.DIM}{" "*4}╔{"═"*38}{Style.NORMAL} [ Menu De Opciones ] {Style.DIM}{"═"*38}╗{Style.NORMAL}")
        print(f"{" "*12}{Fore.GREEN}[1] {Fore.WHITE + Style.NORMAL}Microsoft Office 2016")
        print(f"{" "*12}{Fore.GREEN}[2] {Fore.WHITE + Style.NORMAL}Microsoft Office 2019")
        print(f"{" "*12}{Fore.GREEN}[2] {Fore.WHITE + Style.NORMAL}Microsoft Office 2019")
        print(f"{" "*12}{Fore.GREEN}[4] {Fore.WHITE + Style.NORMAL}Claves Genericas Para Office")
        print(f"{" "*4}{Style.DIM}╚{"═"*98}╝{Style.RESET_ALL}")

        while True:
            print(f"{Style.BRIGHT}{" "*14}> [OPCION]:{Style.RESET_ALL} ", end="", flush=True)
            opc = ""
            while True:
                char = msvcrt.getwch()
                if char == "\r":
                    print()
                    break
                elif char == "\b":
                    if opc:
                        opc = opc[:-1]
                        sys.stdout.write("\b \b")
                        sys.stdout.flush()
                    continue
                opc += char

                if (char == "1") or (char == "2") or (char == "3") or (char == "4"):
                    sys.stdout.write(Fore.GREEN + Style.DIM + char + Style.RESET_ALL)
                else:
                    sys.stdout.write(Fore.RED + char + Style.RESET_ALL)
                sys.stdout.flush()
            duration(0.5)

            validador = False

            if (opc == "1"):
                office = "MicrosoftOffice2016"
                validador = True
            elif (opc == "2"):
                office = "MicrosoftOffice2019"
                validador = True
            elif (opc == "3"):
                office = "MicrosoftOffice2016"
                validador = True
            elif (opc == "4"):
                duration(2)
                print(f"{" "*14}{Fore.BLUE + Style.BRIGHT}[!] Abriendo el navegador.{Style.RESET_ALL}")
                webbrowser.open("https://nesantimeproyect.netlify.app/proyectos/v/blur_kms/Codes_Office")
                break
            else:
                print(f"{" "*14}{Fore.RED + Style.BRIGHT}[!] Error. Opcion Invalida.{Style.RESET_ALL}")

            if (validador):
                start_office(office)
                break




def Main():
    icon()
    print(f"{Style.BRIGHT}{" "*4}╔{"═"*38} [ Menu De Opciones ] {"═"*38}╗{Style.NORMAL}")
    print(f"{" "*8}{Fore.GREEN}[1] {Fore.WHITE + Style.NORMAL}🚀 Activar Windows")
    print(f"{" "*8}{Fore.GREEN}[2] {Fore.WHITE + Style.NORMAL}📄 Activar Office")
    print(f"{" "*8}{Fore.GREEN}[3] {Fore.WHITE + Style.NORMAL}❎ Salir")
    print(f"{" "*4}{Style.BRIGHT}╚{"═"*98}╝{Style.NORMAL}")

    while True:
        print(f"\n{Style.BRIGHT  + Fore.YELLOW}{" "*10}> {Style.NORMAL}[OPCION]:{Style.RESET_ALL} ", end="", flush=True)
        opc = ""

        while True:
            char = msvcrt.getwch()
            if char == "\r":
                print()
                break
            elif char == "\b":
                if opc:
                    opc = opc[:-1]
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
                continue
            opc += char

            if (char == "1") or (char == "2") or (char == "3"):
                sys.stdout.write(Fore.GREEN + char + Style.RESET_ALL)
            else:
                sys.stdout.write(Fore.RED + char + Style.RESET_ALL)
            sys.stdout.flush()
        
        duration(1)

        if (opc == "1"):
            menu_windows(SO_Windows(), Bits_Windows())
            break

        elif (opc == "2"):
            menu_office(detect_office())
            break

        elif (opc == "3"):
            icon()
            print(f"{Fore.YELLOW}{" "*4}[!] {Fore.WHITE}Cerrando...")
            duration(1)
            break
        
        else:
            print(f"{" "*10}{Fore.RED + Style.BRIGHT}[!] Error. Opcion Invalida.{Style.RESET_ALL}")
        

initiator_blurkms()