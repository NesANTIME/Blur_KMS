import sys
import time
from colorama import Style, Fore, init

from resources.controller.modules import connect_internet, Launch_KMS
from resources.controller.animations import icon_, menu_information
from resources.files.load_config import return_data_repo


init(autoreset=True)


def main_key_kms_init(system: str, arq: str):
    icon_("clear")

    menu_information({
        "title": "[!] AVISO IMPORTANTE [!]",
        "opcs": [ 
            f"{Style.DIM}El uso de claves KMS no oficiales para activar productos de", 
            f"{Style.DIM}Microsoft es una violación de sus términos de servicio",
            f"{Style.DIM}y puede ser ilegal en muchas regiones.",
            " ",
            f"[!] El servidor que se usara para procesar la activacion,",
            "es ' kms.digibody.ir'."
        ]
    })

    try:
        input(f"{' '*4}[+] {Style.DIM}Enter para aceptar y continuar, o presione 'ctrl + c' en cualquier momento para cancelar...{Style.RESET_ALL}")
        pass
    except KeyboardInterrupt:
        print(f"\n{' '*4}Interrupción detectada (Ctrl+C)")
        sys.exit(1)


    
    icon_("clear")
    menu_information( 
        { 
            "title": "[!] Información Obtenida del Sistema [!]", 
            "opcs": [
                f"{Fore.BLUE}💻 Sistema Operativo:{Fore.RESET} {system}", 
                f"{Fore.BLUE + Style.BRIGHT}📲 Arquitectura: {Style.RESET_ALL}{arq}"
            ]
        }
    )


    try:

        while True:
            print(f"\n{' '*4}[🛜 ] Comprobando conexion a servidores", end="", flush=True)

            for _ in range(6):
                sys.stdout.write(f"\r{' '*4}[🛜 ] Comprobando conexion a servidores{"."*_}\033[K")
                sys.stdout.flush()
                time.sleep(0.5)

            if (connect_internet()):
                print(f"\r{' '*4}[🛜 ] [✅ ] Conexión establecida\033[K")
                break

            else:
                print(f"\r{' '*4}[🛜 ] [❌ ] Conexión no establecida\033[K")
                time.sleep(2)

            time.sleep(2)


        print(
            f"\n{Fore.YELLOW + Style.DIM}{' '*4}┌{'─'*24} {Style.BRIGHT}Iniciando Programa de Activacion{Style.DIM} {'─'*24}┐{Style.RESET_ALL}"
            )
        
        
        data_repo = return_data_repo("lista_key_kms")
        clave_kms = data_repo.get(system)

        controller_ = Launch_KMS(clave_kms)

        print(f"{' '*5}{Style.BRIGHT}[1] => {Style.NORMAL}Instalando clave KMS \n")
        salida = controller_.execute_P1()
        for line in salida.stdout:
            print(f"{' '*6}{Style.DIM}{line}")

        time.sleep(2)

        print(f"{' '*5}{Style.BRIGHT}[2] => {Style.NORMAL}Configurando administrador de claves KMS 'kms.digiboy.ir'\n")
        salida = controller_.execute_P2()
        for line in salida.stdout:
            print(f"{' '*6}{Style.DIM}{line}")

        print(f"{' '*5}{Style.BRIGHT}[3] => {Style.NORMAL}Activando Clave\n")
        salida = controller_.execute_P3()
        for line in salida.stdout:
            print(f"{' '*6}{Style.DIM}{line}")

        
        print(f"{' '*5}{Fore.GREEN + Style.BRIGHT}👍  Listo!, {system} activado con exito!")

        print(f"{" "*5}{Style.BRIGHT}NOTA: {Style.DIM}Es nesesario reiniciar, por favor cierre todas las aplicaciones abiertas\n{" "*10}antes de continuar.")
        input(f"{Fore.YELLOW + Style.DIM}\n{" "*12} <Presione Enter Para Reiniciar>")

        controller_.finalizar()

    except KeyboardInterrupt:
        print(f"\n{' '*4}Interrupción detectada (Ctrl+C)")
        sys.exit(1)