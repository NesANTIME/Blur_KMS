import sys
import readchar
from colorama import Style, Fore, init

from resources.controller.modules import clear
from resources.files.load_config import return_data
init(autoreset=True)

def icon_(msg = None):
    if (msg == "clear"):
        clear()

    lista_icon = return_data("icon")
    lista_icon[-1] = f"{lista_icon[-1]} V{return_data("version")}"

    print()
    for i in lista_icon:
        print(f"{' '*4}{i}")
    


def create_menu(content):
    opciones_dict = content.get("opc", {})
    keys = list(opciones_dict.keys())
    total = len(keys)
    idx = 0

    print(f"{Style.DIM}\n{' '*4}Usa [↑/↓] y 'Enter' para seleccionar.")
    
    print(f"{Style.BRIGHT}{' '*4}╔{'═'*30} [ Menu De Opciones ] {'═'*30}╗")
    for numero, opc in opciones_dict.items():
        print(f"{' '*4}║{' '*4}{Fore.GREEN}[{numero}] {Fore.WHITE + Style.NORMAL}{opc}{' ' * (73 - len(opc))}║")
    print(f"{' '*4}{Style.BRIGHT}╠{'═'*82}╣")
    print(f"{' '*4}║{' '*82}║")
    print(f"{' '*4}{Style.BRIGHT}╚{'═'*82}╝")

    sys.stdout.write("\033[?25l")
    sys.stdout.write("\033[2A")
    sys.stdout.flush()

    try:
        while True:
            current_key = keys[idx]
            contenido = opciones_dict[current_key]

            texto_seleccion = f"{' '*4}║  {Style.BRIGHT}{Fore.CYAN}[!] Opcion => {Fore.YELLOW}{contenido}"
            relleno = " " * (81 - (len(f"  [!] Opcion => {contenido}")))
            
            sys.stdout.write(f"\r{texto_seleccion}{relleno}{Style.BRIGHT}{Fore.WHITE}║")
            sys.stdout.flush()

            key_pressed = readchar.readkey()

            if key_pressed == readchar.key.UP:
                idx = (idx - 1) % total
            elif key_pressed == readchar.key.DOWN:
                idx = (idx + 1) % total
            elif key_pressed == readchar.key.ENTER:
                sys.stdout.write("\033[2B\n")
                return current_key
    finally:
        sys.stdout.write("\033[?25h")
        sys.stdout.flush()



def menu_information(content):
    title = content.get("title")

    print(f"\n{Style.BRIGHT}{' '*4}┌{'─'*20} {title} {'─'*20}┐")
    for i in content.get("opcs", []):
        print(f"{' '*6}{i}")

    print(f"{' '*4}{Style.BRIGHT}└{'─'* (42 + len(title))}┘")