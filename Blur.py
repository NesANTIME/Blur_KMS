# By NesAnTime the BlurKMS 5.0
import os
import sys
import time
import json
import ctypes
import subprocess
import tkinter as tk
from PIL import Image, ImageTk


def load_config(ruta):
    with open(ruta, "r", encoding="utf-8") as archivo:
        return json.load(archivo)

def run_a_admin():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:
        script = sys.executable
        params = " ".join([f'"{arg}"' for arg in sys.argv])
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", script, params, None, 1
            )
        except Exception as e:
            print(f"Error al intentar ejecutar como administrador: {e}")
        sys.exit(0)

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

    
def initiator_check():
    run_a_admin()

    def retur_os(archivo):
        ruta_base = os.path.dirname(__file__)
        dato = os.path.join(ruta_base, archivo)
        return dato

    var_result = True
    ruta_base = os.path.dirname(__file__)
    ruta_config = os.path.join(ruta_base, "resources", "core", "config_blurkms.json")
    ruta_python = os.path.join(ruta_base,"resources","core","interprete","python-3.13.7-amd64", "python.exe")

    if not retur_os(ruta_config):
        var_result = False
    else:
        config_var = load_config(retur_os(ruta_config))
        image = retur_os(config_var["Rutas_Archivos_BlurKMS"][0])

        if not os.path.isfile(image):
            var_result = None

        for archivo in config_var["Rutas_Archivos_BlurKMS"]:
            archivo = retur_os(archivo)

            if not os.path.isfile(archivo):
                var_result = False
                break
        
        core = retur_os(config_var["Rutas_Archivos_BlurKMS"][1])
        
    return var_result, image, core, ruta_python


def start(result, url_imagen, core, ruta_python):
    class SplashScreen(tk.Tk):
        def __init__(self):
            super().__init__()

            self.overrideredirect(True)
            self.attributes("-topmost", True)
            self.config(bg='black')
            self.wm_attributes('-transparentcolor', 'black')
            self.attributes('-alpha', 0.0)

            if (result == True) or (result == False):
                image = Image.open(url_imagen).convert("RGBA")
                image = image.resize((600, 300), Image.Resampling.LANCZOS)
                self.logo = ImageTk.PhotoImage(image)

                width, height = image.size
                total_height = height + 60

                x = (self.winfo_screenwidth() - width) // 2
                y = (self.winfo_screenheight() - total_height) // 2
                self.geometry(f"{width}x{total_height}+{x}+{y}")

                tk.Label(self, image=self.logo, bg='black', bd=0).pack()
                self.label_status = tk.Label(self, text="Cargando...", bg='black', fg='white', font=("Segoe UI", 12))
                self.label_status.pack(pady=5)

                self.progress = tk.Canvas(self, width=width, height=10, bg='black', highlightthickness=0)
                self.progress.pack()
                self.bar = self.progress.create_rectangle(0, 0, 0, 10, fill='#d9faf7')

                self.after(0, self.fade_in)
            else:
                width, height = 600, 100
                x = (self.winfo_screenwidth() - width) // 2
                y = (self.winfo_screenheight() - height) // 2
                self.geometry(f"{width}x{height}+{x}+{y}")

                self.label_status = tk.Label(self, text="Cargando...", bg='black', fg='white', font=("Segoe UI", 14))
                self.label_status.pack(pady=(25, 5))

                self.progress = tk.Canvas(self, width=width, height=10, bg='black', highlightthickness=0)
                self.progress.pack()
                self.bar = self.progress.create_rectangle(0, 0, 0, 10, fill='#d9faf7')

                self.after(0, self.fade_in)


        def fade_in(self):
            for i in range(0, 21):
                self.attributes("-alpha", i / 20)
                self.update()
                time.sleep(0.03)
            self.animate_bar()

        def animate_bar(self):
            self.actualizar_estado("Verificando Integridad...")
            if (result == True):
                self.progress_bar_animar()
                self.actualizar_estado("ℹ️ Integridad Completa")
                time.sleep(0.6)
                self.actualizar_estado("Iniciando")
                time.sleep(0.6)
                self.fade_out()
            else:
                if (result == False):
                    complet = 10
                    title_message = "⚠️ Integridad Crítica."
                    description_message = " Faltan archivos importantes."

                elif (result == None):
                    complet = 80
                    title_message = "⚠️ Integridad Comprometida."
                    description_message = " Faltan algunos archivos, Se puede continuar."

                self.mostrar_error(title_message, description_message)
                self.progress_bar_animar(ralentizar=True, porcentaje = complet)

                if (result == False):
                    sys.exit(1)

        def progress_bar_animar(self, ralentizar=False, porcentaje=100):
            width = self.progress.winfo_width()
            max_width = width * (porcentaje / 100)

            for i in range(0, int(max_width) + 1, 10):
                self.progress.coords(self.bar, 0, 0, i, 10)
                self.update()
                time.sleep(0.01 if not ralentizar else 0.03)
            self.progress.coords(self.bar, 0, 0, max_width, 10)
            self.update()

        def actualizar_estado(self, mensaje, color="white", barra="#d9faf7"):
            self.label_status.config(text=mensaje, fg=color)
            self.progress.itemconfig(self.bar, fill=barra)
            self.update()
        
        def mostrar_error(self, titulo, mensaje):
            root = tk.Tk()
            root.withdraw()
            self.actualizar_estado(titulo + " " + mensaje, color="red", barra="red",)
            self.update()
            root.attributes('-topmost', True)
            time.sleep(2)
            ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, 0x10 | 0x1000)
            root.destroy()

        def fade_out(self):
            for i in range(20, -1, -1):
                self.attributes("-alpha", i / 20)
                self.update()
                time.sleep(0.03)
            self.destroy()

            try:
                print(f"[!] Ejecutando corepy: \n{ruta_python} {core}")
                subprocess.run([ruta_python, core], check=True)
            except Exception as e:
                ctypes.windll.user32.MessageBoxW(0, f"Error al iniciar core:\n{e}", "Error", 0x10 | 0x1000)

            sys.exit(0)


    return SplashScreen


def core_check():
    estado, url_img, core, ruta_python = initiator_check()

    if (estado == None):
        sys.exit(1)

    SplashScreen = start(estado, url_img, core, ruta_python)
    app = SplashScreen()
    app.mainloop()


if __name__ == "__main__":
    core_check()