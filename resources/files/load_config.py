import os
import sys
import json
import requests
from pathlib import Path



# Retornar configuracion.json
def return_config():
    if (getattr(sys, 'frozen', False)):
        ruta_config = Path(sys.executable).parent
    else:
        ruta_config = Path(__file__).parent

    ruta_config = os.path.join(ruta_config, "config.json")

    if (not os.path.isfile(ruta_config)):
        sys.exit(1)
    
    with open(ruta_config, "r", encoding="utf-8") as archivo:
        return json.load(archivo)
    


# Retornar configuracion.json de el repo
def return_configRepo():
    try:
        response = requests.get(return_data("return_url_configuration"), timeout=10)
        response.raise_for_status()
        return response.json()
    
    except requests.exceptions.Timeout:
        raise("Tiempo de espera agotado al descargar el JSON")
    except requests.exceptions.HTTPError as e:
        raise(f"Error HTTP al descargar JSON: {e}")
    except requests.exceptions.RequestException as e:
        raise(f"Error de red: {e}")
    except ValueError:
        raise("El contenido descargado no es un JSON válido")
    



def return_data(opc):
    config_json = return_config()

    if (opc == "icon"):
        return config_json.get("information").get("icon", [])
    elif (opc == "version"):
        return config_json.get("information").get("version")
    

    elif (opc == "return_url_configuration"):
        return config_json.get("connections").get("configurations_url_windows")
    

    elif (opc == "return_hosts"):
        return config_json.get("servers", [])
    

def return_data_repo(opc):
    config_json_repo = return_configRepo()

    if (opc == "lista_key_kms"):
        return config_json_repo.get("claves_kms_for_windows", {})