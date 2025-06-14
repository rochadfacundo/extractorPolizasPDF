import pdfplumber
import pandas as pd
import re
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

# Ruta del PDF
pdf_path = "ARCHIVO[20250613152317355357].PDF"

# Estructura base de los datos
datos = {
    "Marca": "",
    "Modelo": "",
    "Año": "",
    "Suma Asegurada": "",
    "Premio": "--",
    "Cláusula de Ajuste": "--",
    "Cobertura": "--"
}

# Extraer texto completo del PDF
with pdfplumber.open(pdf_path) as pdf:
    texto_completo = ""
    for page in pdf.pages:
        texto_completo += page.extract_text() + "\n"

    def buscar(patron, multilinea=True):
        flags = re.MULTILINE | re.IGNORECASE if multilinea else re.IGNORECASE
        resultado = re.search(patron, texto_completo, flags)
        return resultado.group(1).strip() if resultado else ""

    # Marca, Modelo, Año
    marca_modelo = buscar(r"Marca.*?: ([^\n]+)")
    if marca_modelo:
        partes = marca_modelo.strip().split(" ")
        datos["Marca"] = partes[0]
        datos["Modelo"] = " ".join(partes[1:-1]) if len(partes) > 2 else partes[1] if len(partes) > 1 else ""
        datos["Año"] = partes[-1] if partes[-1].isdigit() else ""

    # Suma Asegurada
    datos["Suma Asegurada"] = buscar(r"Suma Asegurada:.*?\$\s*([0-9.,]+)")

    # Refacturación mensual o no
    es_mensual = re.search(r"Refactura.*Mensual", texto_completo, re.IGNORECASE)

    if es_mensual:
        premio = buscar(r"PREMIO TOTAL.*?\n.*?\n([0-9.]+,[0-9]{2})")
        if not premio:
            premio = buscar(r"Cuota\s+Vto\.Asegu\..*?\n\s*1\s+[0-9.]+\s+([0-9.,]+)")
        if not premio:
            premio = buscar(r"Importe\s*\n([0-9.]+,[0-9]{2})")
        if not premio:
            premio = buscar(r"Premio\s*[:]*\s*\$?\s*([0-9.,]+)")
        datos["Premio"] = premio if premio else "--"
        datos["Cláusula de Ajuste"] = "--"
    else:
        match = re.search(r"Suma Asegurada.*?\n(.{0,100}ajuste.{0,100})", texto_completo, re.IGNORECASE)
        datos["Cláusula de Ajuste"] = match.group(1).strip() if match else "--"
        datos["Premio"] = buscar(r"Premio\s*[:]*\s*\$?\s*([0-9.,]+)") or "--"

    # Cobertura precisa
    cobertura_match = re.search(
        r"Coberturas especif\.del riesgo\s*:?[\s\n]*((?:.+\n)+?)(?=(?:Descripción del Riesgo|Uso del vehículo|Suma Asegurada|VALOR DE REPOSICION))",
        texto_completo,
        re.IGNORECASE
    )
    datos["Cobertura"] = cobertura_match.group(1).strip() if cobertura_match else "--"

# Orden de columnas y exportación
columnas_ordenadas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura"]
df = pd.DataFrame([{col: datos[col] for col in columnas_ordenadas}])
nombre_archivo = "mercantil.xlsx"
df.to_excel(nombre_archivo, index=False)

# Estética en Excel
wb = load_workbook(nombre_archivo)
ws = wb.active

# Encabezado verde claro
fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
for cell in ws[1]:
    cell.fill = fill_verde

# Ajuste de columnas y estilo de celdas
for col in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if col[0].value == "Cobertura":
        ws.column_dimensions[col_letter].width = 60  # Ancho fijo para cobertura
    else:
        ws.column_dimensions[col_letter].width = max_length + 2

# Ajuste de altura de filas
for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
    max_lines = max(str(cell.value).count("\n") + 1 if cell.value else 1 for cell in row)
    ws.row_dimensions[row[0].row].height = max(15, max_lines * 15)

wb.save(nombre_archivo)
print(f"✅ Excel generado como {nombre_archivo}")
