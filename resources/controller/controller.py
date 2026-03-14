import sys
from colorama import Style, Fore, init

from resources.files.load_config import return_data_repo
from resources.controller.animations import icon_, menu_information, create_menu

from resources.controller.funciones.key_windows import main_key_kms_init


init(autoreset=True)

LISTA_KEY_KMS = return_data_repo("lista_key_kms")




def menu_system_windows_key(system_, arquitecture_):
    icon_("clear")

    menu_information( 
        { 
            "title": "[!] Información Obtenida del Sistema [!]", 
            "opcs": [
                f"{Fore.BLUE}💻 Sistema Operativo:{Fore.RESET} {system_}", 
                f"{Fore.BLUE + Style.BRIGHT}📲 Arquitectura: {Style.RESET_ALL}{arquitecture_}"
            ]
        }
    )

    opc_windows_menu = create_menu(
        { 
            "opc": { 
                1: "💱 Cambiar Version de Windows", 
                2: "💱 Cambiar Arquitectura de Windows", 
                3: "🚀 Continuar Con Activacion de Windows",
                4: "❎ Salir" 
            } 
        } 
    )

    if (opc_windows_menu == 1):
        icon_("clear")
        
        print(f"\n{' '*4}[!] Menu, para cambiar la version de windows.")
        cont = 1
        lista = {}

        for clave, valor in LISTA_KEY_KMS.items():
            lista[cont] = clave
            cont += 1
             
        win_manual = create_menu( {"opc": lista} )

        menu_system_windows_key(lista.get(win_manual), arquitecture_)


    elif (opc_windows_menu == 2):
        icon_("clear")

        print(f"\n{' '*4}[!] Menu, para cambiar la arquitectura de windows.")

        arq_manual = create_menu( { "opc": { 1: "64 bits ", 2: "32 bits " } } )

        if (arq_manual == 1):
            arq_manual = "64 bits"
        else:
            arq_manual = "32 bits"

        menu_system_windows_key(system_, arq_manual)

    elif (opc_windows_menu == 3):
        main_key_kms_init(system_, arquitecture_)

    else:
        sys.exit(0)