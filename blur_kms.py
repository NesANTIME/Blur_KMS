from colorama import Style, Fore, init

from resources.controller.modules import detect_controller
from resources.controller.animations import icon_, create_menu
from resources.controller.controller import menu_system_windows_key

init(autoreset=True)

def main():
    icon_("clear")
    opc_pricipal = create_menu(
        { "opc": { 1: "🚀 Activar Windows", 2: "📄 Activar Office", 3: "❎ Salir" } } 
    )

    if (opc_pricipal == 1):
        menu_system_windows_key( detect_controller("system"), detect_controller("arquitecture") )



    


main()