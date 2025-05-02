import os
import sys
import time
import json
import ctypes
import tkinter as tk
from PIL import Image, ImageTk


def Load_BD():
    with open("Scritps/BD_BlurKMS.json", "r", encoding="utf-8") as archivo:
        return json.load(archivo)
    
def Lanzador_Seguiridad():
    data = Load_BD()
    Archivos_Complementos_Faltantes = []
    Archivos_Rutas = data["Rutas_Archivos_BlurKMS"]

    if not os.path.isfile("Scritps/BD_BlurKMS.json"):
            return "ERROR: 19726KHYU2654&"
    else:
        for archivo in Archivos_Rutas:
            if not os.path.isfile(archivo):
                Archivos_Complementos_Faltantes.append(archivo)
    
    if Archivos_Complementos_Faltantes:
        return "ERROR: DHGFE983HME034FNJF490"
    else:
        return "OK: 19726KHYU2654&"
    

class SplashScreen(tk.Tk):
    def __init__(self):
        super().__init__()

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.config(bg='black')
        self.wm_attributes('-transparentcolor', 'black')
        self.attributes('-alpha', 0.0)

        if os.path.exists("Scritps/Banner_Lanzador.png"):
            image = Image.open("Scritps/Banner_Lanzador.png").convert("RGBA")
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
        resultado = Lanzador_Seguiridad()

        if resultado.startswith("OK: 19726KHYU2654&"):
            self.actualizar_estado("Verificando Integridad...")
            self.progress_bar_animar()
            self.actualizar_estado("ℹ️ Integridad Completa")
            time.sleep(0.6)
            self.actualizar_estado("Iniciando")
            time.sleep(0.6)
            self.fade_out()
        else:
            self.actualizar_estado("Verificando Integridad...")
            if resultado == "ERROR: 19726KHYU2654&":
                self.progress_bar_animar(ralentizar=True, porcentaje=40)
                self.mostrar_error("⚠️ Integridad Crítica.", " Faltan archivos importantes")
            elif resultado == "ERROR: DHGFE983HME034FNJF490":
                self.progress_bar_animar(ralentizar=True, porcentaje=80)
                self.mostrar_error("⚠️ Integridad Fatal.", " Descarge Nuevamente el programa")
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
        sys.exit(0)

if __name__ == "__main__":
    app = SplashScreen()
    app.mainloop()
