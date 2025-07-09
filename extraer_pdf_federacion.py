import pdfplumber
import pandas as pd
import re
import os
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

pdf_path = "data/polizas/polizaFed4.pdf"
nombre_excel = "federacion_patronal.xlsx"

datos = {
    "Marca": "--",
    "Modelo": "--",
    "Año": "--",
    "Suma Asegurada": "--",
    "Premio": "--",
    "Cláusula de Ajuste": "--",
    "Cobertura": "--",
    "Archivo": os.path.basename(pdf_path)
}

with pdfplumber.open(pdf_path) as pdf:
    texto_completo = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

    # ---------- EXTRACCIONES ----------
    def buscar(patron, texto=texto_completo):
        match = re.search(patron, texto, re.IGNORECASE)
        return match.group(1).strip() if match else "--"

    # Marca
    datos["Marca"] = buscar(r"Marca\s+([A-Z0-9 /-]+)")

    # Modelo
    datos["Modelo"] = buscar(r"Modelo\s+([A-Z0-9 /.-]+)")

    # Año
    datos["Año"] = buscar(r"A[ÑN]O\s+(\d{4})")

    # Suma Asegurada: más confiable desde “CGDA-DESTRUCCION TOTAL SUMA ASEGURADA”
    suma_match = re.search(r"CGDA-DESTRUCCION TOTAL\s+SUMA ASEGURADA\s+([0-9.]+,[0-9]{2})", texto_completo, re.IGNORECASE)
    if not suma_match:
        suma_match = re.search(r"SUMA ASEGURADA\s+([0-9.]+,[0-9]{2})", texto_completo, re.IGNORECASE)
    datos["Suma Asegurada"] = suma_match.group(1) if suma_match else "--"

    # Premio del Endoso (único monto entero con coma)
    premio_match = re.search(r"Premio del Endoso\s*\$?\s*([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})", texto_completo, re.IGNORECASE)

    datos["Premio"] = premio_match.group(1) if premio_match else "--"

    # Cláusula de Ajuste (si existe)
    ajuste = buscar(r"Ajuste Autom[aá]tico.*?([0-9]{1,3}\s*%)")
    datos["Cláusula de Ajuste"] = ajuste if ajuste != "--" else "--"

    # Cobertura (busca línea del PLAN)
    match_plan = re.search(r"PLAN\s*\n?([A-Z0-9 \-]+)", texto_completo, re.IGNORECASE)
    datos["Cobertura"] = match_plan.group(1).strip() if match_plan else "--"

# ---------- EXPORTAR A EXCEL ----------
columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura", "Archivo"]
fila = {col: datos[col] for col in columnas}

if os.path.exists(nombre_excel):
    df_existente = pd.read_excel(nombre_excel)
    df = pd.concat([df_existente, pd.DataFrame([fila])], ignore_index=True)
else:
    df = pd.DataFrame([fila])

df.to_excel(nombre_excel, index=False)

# ---------- ESTÉTICA EXCEL ----------
wb = load_workbook(nombre_excel)
ws = wb.active
fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

for cell in ws[1]:
    cell.fill = fill

for col in ws.columns:
    col_letter = get_column_letter(col[0].column)
    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
    for cell in col:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[col_letter].width = 60 if col[0].value == "Cobertura" else max_len + 2

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    max_lines = max(str(c.value).count("\n") + 1 if c.value else 1 for c in row)
    ws.row_dimensions[row[0].row].height = max(15, max_lines * 15)

wb.save(nombre_excel)
print(f"✅ Excel generado correctamente como '{nombre_excel}'")
