import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import sys
import threading

# Agregar ruta para importar desde core y utils
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.enums import Compania
from utils.extraer_pdf_atm import procesar_atm
from utils.extraer_pdf_federacion import procesar_federacion
from utils.extraer_pdf_rivadavia import procesar_rivadavia
from utils.extraer_pdf_mercantil import procesar_mercantil
from utils.extraer_pdf_rus import procesar_rus  # 👈 nuevo import

# === FUNCIONES DE RUTA ===
def obtener_ruta_logo():
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", "logo.png")

# === CONFIG VENTANA ===
root = tk.Tk()
root.title("Extractor de Pólizas PDF")
root.geometry("550x520")
root.configure(bg="#f5fff0")

# === ICONO BARRA DE TAREAS ===
try:
    icono_png = obtener_ruta_logo()
    if os.path.exists(icono_png):
        img_icon = ImageTk.PhotoImage(Image.open(icono_png).resize((32, 32)))
        root.iconphoto(True, img_icon)
        print("✅ Icono cargado con iconphoto (PNG)")
    else:
        print("❌ No se encontró el icono:", icono_png)
except Exception as e:
    print("❌ Error cargando icono con iconphoto:", e)

# === CONTENEDOR PRINCIPAL ===
frame = tk.Frame(root, bg="#f5fff0")
frame.pack(fill="both", expand=True)

# === LOGO VISUAL ===
try:
    logo_path = obtener_ruta_logo()
    if os.path.exists(logo_path):
        logo_img = ImageTk.PhotoImage(Image.open(logo_path).resize((100, 100)))
        logo_label = tk.Label(frame, image=logo_img, bg="#f5fff0", borderwidth=0, highlightthickness=0)
        logo_label.image = logo_img
        logo_label.pack(pady=10)
except Exception as e:
    print("⚠️ No se pudo cargar el logo:", e)

# === COMPAÑÍA ===
tk.Label(frame, text="Seleccioná la compañía aseguradora:", font=("Arial", 12), bg="#f5fff0").pack(pady=5)
combo = ttk.Combobox(frame, font=("Arial", 11), state="readonly")
combo["values"] = [c.value for c in Compania]
combo.pack(pady=5)

# === PDFs ===
entry_pdfs = tk.Entry(frame, width=60, font=("Arial", 10), state="readonly")
entry_pdfs.pack(pady=5)

btn_archivos = tk.Button(
    frame,
    text="Seleccionar PDFs",
    font=("Arial", 11),
    bg="#94c484",
    fg="white",
    activebackground="#7aa76f",
    relief="raised",
    bd=3,
    cursor="hand2",
    command=lambda: seleccionar_pdfs(),
    state="disabled"
)
btn_archivos.pack(pady=5)

def seleccionar_pdfs():
    archivos = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if archivos:
        entry_pdfs.config(state="normal")
        entry_pdfs.delete(0, tk.END)
        entry_pdfs.insert(0, ";".join(archivos))
        entry_pdfs.config(state="readonly")

def habilitar_pdf_inputs(event=None):
    entry_pdfs.config(state="readonly")
    btn_archivos.config(state="normal")

combo.bind("<<ComboboxSelected>>", habilitar_pdf_inputs)

# === CONSOLA LOGS ===
resultado = tk.Text(frame, height=10, font=("Consolas", 10), state="disabled", bg="#ffffff")
resultado.pack(padx=10, pady=10, fill="both", expand=True)

def logear(texto):
    resultado.config(state="normal")
    resultado.insert(tk.END, texto + "\n")
    resultado.see(tk.END)
    resultado.config(state="disabled")

# === EJECUTAR EXTRACCIÓN ===
def ejecutar_procesamiento():
    compania = combo.get()
    archivos = entry_pdfs.get().split(";")
    if not compania or not archivos or not archivos[0]:
        messagebox.showerror("Faltan datos", "Seleccioná una compañía y uno o más PDFs.")
        return

    logear(f"🔄 Procesando archivos para {compania}...")

    try:
        if compania == Compania.ATM.value:
            procesar_atm(archivos)
        elif compania == Compania.FEDERACION.value:
            procesar_federacion(archivos)
        elif compania == Compania.RIVADAVIA.value:
            procesar_rivadavia(archivos)
        elif compania == Compania.MERCANTIL.value:
            procesar_mercantil(archivos)
        elif compania == Compania.RIO_URUGUAY.value:
            procesar_rus(archivos)
        logear("✅ Extracción finalizada correctamente.")
    except Exception as e:
        logear(f"❌ Error durante la extracción: {e}")

# === BOTÓN EXTRAER ===
btn_extraer = tk.Button(
    frame,
    text="Extraer a Excel",
    font=("Arial", 12, "bold"),
    bg="#94c484",
    fg="white",
    activebackground="#7aa76f",
    relief="raised",
    bd=3,
    cursor="hand2",
    command=lambda: threading.Thread(target=ejecutar_procesamiento, daemon=True).start()
)
btn_extraer.pack(pady=10)

# === INICIAR APP ===
root.mainloop()
