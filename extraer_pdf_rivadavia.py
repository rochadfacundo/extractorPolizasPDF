import pdfplumber
import pandas as pd
import re
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

pdf_path = "C_POLAUT_4529039_20250613_1519275_18922_024002239342001.pdf"

# Marcas compuestas comunes
marcas_compuestas = [
    "MERCEDES BENZ", "ALFA ROMEO", "LAND ROVER", "ROLLS ROYCE", "ASTON MARTIN",
    "CHEVROLET CAPTIVA", "VOLKSWAGEN AMAROK", "GREAT WALL", "MINI COOPER"
]

# Estructura base
datos = {
    "Marca": "",
    "Modelo": "",
    "Año": "",
    "Suma Asegurada": "--",
    "Premio": "--",
    "Cláusula de Ajuste": "--",
    "Cobertura": "--"
}

# Extraer texto + buscar premio en tabla
with pdfplumber.open(pdf_path) as pdf:
    texto = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

    def buscar(patron, multilinea=True):
        flags = re.MULTILINE | re.IGNORECASE if multilinea else re.IGNORECASE
        match = re.search(patron, texto, flags)
        return match.group(1).strip() if match else ""

    # Año
    anio_match = re.search(r"MODELO[:\s]+(\d{4})", texto, re.IGNORECASE)
    if anio_match:
        datos["Año"] = anio_match.group(1)

    # Marca y Modelo
    marca_linea = re.search(r"MARCA[:\s]+(.+?)(?:\n|$)", texto, re.IGNORECASE)
    if marca_linea:
        linea = marca_linea.group(1).strip()
        linea = re.sub(r"ASIENTOS.*", "", linea).strip()
        matched = False
        for marca in marcas_compuestas:
            if linea.upper().startswith(marca):
                datos["Marca"] = marca
                datos["Modelo"] = linea[len(marca):].strip()
                matched = True
                break
        if not matched:
            partes = linea.split()
            datos["Marca"] = partes[0]
            datos["Modelo"] = " ".join(partes[1:]) if len(partes) > 1 else ""

    # Suma Asegurada
    suma_match = re.search(r"Suma máxima por Acontecimiento\s+\$?\s*([0-9.,]+)", texto, re.IGNORECASE)
    if suma_match:
        datos["Suma Asegurada"] = suma_match.group(1)

    # Premio
    premio_match = re.search(r"PREMIO\s+\$?\s*([0-9.,]+)", texto, re.IGNORECASE)
    if premio_match:
        datos["Premio"] = premio_match.group(1)
    else:
        # Buscar en tablas
        for page in pdf.pages:
            tablas = page.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    for celda in fila:
                        if celda and re.match(r"\d{5,},\d{2}", celda.strip()):
                            datos["Premio"] = celda.strip()
                            break
                    if datos["Premio"] != "--":
                        break
                if datos["Premio"] != "--":
                    break
            if datos["Premio"] != "--":
                break

    # Cobertura
    cobertura_match = re.search(r"(CG-RC\s*0?1\.1\s+Responsabilidad Civil.*?)\n", texto, re.IGNORECASE)
    if cobertura_match:
        datos["Cobertura"] = cobertura_match.group(1).strip()

# Crear Excel
columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura"]
df = pd.DataFrame([{col: datos[col] for col in columnas}])
nombre_archivo = "rivadavia.xlsx"
df.to_excel(nombre_archivo, index=False)

# Estética
wb = load_workbook(nombre_archivo)
ws = wb.active
fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

for cell in ws[1]:
    cell.fill = fill_verde

for col in ws.columns:
    col_letter = get_column_letter(col[0].column)
    max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
    for cell in col:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[col_letter].width = 60 if col[0].value == "Cobertura" else max_length + 2

for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
    max_lines = max(str(cell.value).count("\n") + 1 if cell.value else 1 for cell in row)
    ws.row_dimensions[row[0].row].height = max(15, max_lines * 15)

wb.save(nombre_archivo)
print(f"✅ Excel generado como {nombre_archivo}")
