import pdfplumber
import pandas as pd
import re
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

pdf_path = "4-33316078-0-2.pdf"

datos = {
    "Marca": "--",
    "Modelo": "--",
    "Año": "--",
    "Suma Asegurada": "--",
    "Premio": "--",
    "Cláusula de Ajuste": "--",
    "Cobertura": "--"
}

with pdfplumber.open(pdf_path) as pdf:
    texto = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

    def buscar(patron, multilinea=True):
        flags = re.MULTILINE | re.IGNORECASE if multilinea else re.IGNORECASE
        match = re.search(patron, texto, flags)
        return match.group(1).strip() if match else ""

    # Marca, Modelo y Año (línea completa)
    linea = buscar(r"Marca\s*-\s*modelo\s*-\s*a[ñn]o\s*/.*\n([^\n]+)")
    if linea:
        partes = linea.rsplit(" ", 1)
        if len(partes) == 2:
            datos["Año"] = partes[1].strip()
            marca_modelo = partes[0].strip()
            datos["Marca"] = marca_modelo.split()[0]
            datos["Modelo"] = " ".join(marca_modelo.split()[1:])

    # Cobertura
    cobertura = buscar(r"(TD3\s*-\s*TODO RIESGO CON FRANQUICIA FIJA)", multilinea=True)
    if cobertura:
        datos["Cobertura"] = cobertura

    # Cláusula de Ajuste
    ajuste = buscar(r"Ajuste Autom[aá]tico\s*\([^)]+\)[^\n]*\s+([0-9]{1,3}\s*%)")
    if ajuste:
        datos["Cláusula de Ajuste"] = ajuste

    # Suma Asegurada
    suma = buscar(r"SUMA ASEGURADA[\s\$]*([0-9.,]{8,})")
    if suma:
        datos["Suma Asegurada"] = suma
    else:
        suma = buscar(r"CGDA-DESTRUCCION TOTAL\s*\$?\s*([0-9.,]+)", multilinea=False)
        if suma:
            datos["Suma Asegurada"] = suma

    # Premio
    premio_texto = buscar(r"MONTO TOTAL DEL PREMIO\s*\$?\s*([0-9.,]+)")
    if premio_texto:
        datos["Premio"] = premio_texto
    else:
        for page in pdf.pages:
            tablas = page.extract_tables()
            for tabla in tablas:
                for fila in tabla:
                    for celda in fila:
                        if celda and re.match(r"\d{1,3}(?:\.\d{3})*,\d{2}", celda.strip()):
                            datos["Premio"] = celda.strip()
                            break
                    if datos["Premio"] != "--":
                        break
                if datos["Premio"] != "--":
                    break
            if datos["Premio"] != "--":
                break

# Exportar a Excel
columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura"]
df = pd.DataFrame([{col: datos[col] for col in columnas}])
nombre_excel = "federacion_patronal.xlsx"
df.to_excel(nombre_excel, index=False)

# Estética del Excel
wb = load_workbook(nombre_excel)
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

wb.save(nombre_excel)
print(f"✅ Excel generado como {nombre_excel}")
