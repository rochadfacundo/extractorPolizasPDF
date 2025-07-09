import pdfplumber
import pandas as pd
import re
import os
import json
from collections import Counter
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

# Archivos a procesar
pdfs = [
    "data/polizas/polizaFed.pdf",
    "data/polizas/polizaFed2.pdf",
    "data/polizas/polizaFed3.pdf",
    "data/polizas/polizaFed4.pdf"
]
nombre_excel = "federacion_patronal.xlsx"

# Cargar lista de marcas
with open("assets/marcasFiltradasId.json", "r", encoding="utf-8") as f:
    marcas_data = json.load(f)
lista_marcas = [m["marca"].upper() for m in marcas_data]
lista_marcas_ordenadas = sorted(lista_marcas, key=lambda x: len(x.split()), reverse=True)

columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura", "Archivo"]
filas = []

for pdf_path in pdfs:
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

        def buscar(patron, texto=texto_completo):
            match = re.search(patron, texto, re.IGNORECASE)
            return match.group(1).strip() if match else "--"

        # ------- Año -------
        datos["Año"] = buscar(r"Modelo\s+[^\n]*\n.*\b(\d{4})\b")

        # ------- Marca y Modelo -------
        modelo_match = re.search(r"Modelo\s+[^\n]*\n([^\n]+)", texto_completo, re.IGNORECASE)
        if modelo_match:
            linea = modelo_match.group(1).strip().upper()
            for marca in lista_marcas_ordenadas:
                if linea.startswith(marca):
                    datos["Marca"] = marca.title()
                    datos["Modelo"] = linea[len(marca):].strip().title()
                    break

            anio_match = re.search(r"(19|20)\d{2}$", datos["Modelo"])
            if anio_match:
                datos["Año"] = anio_match.group(0)
                datos["Modelo"] = re.sub(r"\s+(19|20)\d{2}$", "", datos["Modelo"])

        # ------- Suma Asegurada -------
        suma_matches = re.findall(r"SUMA ASEGURADA\s*\$?\s*([\d.,]+)", texto_completo)
        if suma_matches:
            conteo = Counter(suma_matches)
            datos["Suma Asegurada"] = conteo.most_common(1)[0][0]

        # ------- Premio -------
       # Nuevo patrón robusto
        premio_match = re.search(r"PREMIO DEL ENDOSO\s*-?\$?\s*(-?[\d.,]+)", texto_completo)
        if premio_match:
            datos["Premio"] = premio_match.group(1).lstrip("-")

        else:
            premio_fallback = re.findall(r"([0-9.]{1,6},[0-9]{2})", texto_completo)
            posibles_premios = [p for p in premio_fallback if float(p.replace('.', '').replace(',', '.')) < 999999]
            if posibles_premios:
                datos["Premio"] = max(posibles_premios, key=lambda x: float(x.replace('.', '').replace(',', '.')))

        # ------- Cláusula de Ajuste -------
        datos["Cláusula de Ajuste"] = buscar(r"Ajuste Autom[aá]tico.*?([0-9]{1,3}\s*%)")

        # ------- Cobertura -------
        cobertura_match = re.search(r"RIESGOS CUBIERTOS.*?(\n.*?)(?=\n[A-Z ]+|$)", texto_completo, re.IGNORECASE | re.DOTALL)
        if cobertura_match:
            cobertura = cobertura_match.group(1).strip()
            cobertura = re.sub(r"\s{2,}", " ", cobertura.replace("\n", " "))
            datos["Cobertura"] = cobertura

    filas.append({col: datos[col] for col in columnas})

# ---------- EXPORTAR A EXCEL ----------
if os.path.exists(nombre_excel):
    df_existente = pd.read_excel(nombre_excel)
    df = pd.concat([df_existente, pd.DataFrame(filas)], ignore_index=True)
else:
    df = pd.DataFrame(filas)

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
