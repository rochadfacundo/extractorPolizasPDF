# utils/extraer_pdf_atm.py
import pdfplumber
import pandas as pd
import re
import os
import json
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

with open("assets/marcas.json", "r", encoding="utf-8") as f:
    marcas_json = json.load(f)
marcas_lista = [m["marca"].upper() for m in marcas_json]

def buscar(texto, patron, multilinea=True):
    flags = re.MULTILINE | re.IGNORECASE if multilinea else re.IGNORECASE
    resultado = re.search(patron, texto, flags)
    return resultado.group(1).strip() if resultado else ""

def procesar_atm(pdfs):
    filas = []

    for pdf_path in pdfs:
        datos = {
            "Marca": "",
            "Modelo": "",
            "Año": "",
            "Suma Asegurada": "--",
            "Premio": "--",
            "Cláusula de Ajuste": "--",
            "Cobertura": "--",
            "Archivo": os.path.basename(pdf_path)
        }

        with pdfplumber.open(pdf_path) as pdf:
            texto_completo = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

            marca_modelo = buscar(texto_completo, r"MARCA/MODELO:\s+([^\n]+)")
            if marca_modelo:
                texto = marca_modelo.upper()
                for marca_posible in sorted(marcas_lista, key=len, reverse=True):
                    if texto.startswith(marca_posible):
                        datos["Marca"] = marca_posible.title()
                        datos["Modelo"] = texto[len(marca_posible):].strip().title()
                        break

            datos["Año"] = buscar(texto_completo, r"AÑO:\s*(\d{4})")
            suma = buscar(texto_completo, r"SUMA ASEGURADA:\s*([0-9.]+,[0-9]{2})") or \
                   buscar(texto_completo, r"SUMA ASEGURADA:\s*([0-9.]+)")
            datos["Suma Asegurada"] = suma if suma else "--"

            ajuste = buscar(texto_completo, r"CLAUSULA DE AJUSTE AUTOMATICO\s*:\s*(\d+ ?%)")
            datos["Cláusula de Ajuste"] = ajuste if ajuste else "--"

            cobertura = buscar(texto_completo, r"COBERTURA:\s*([^\n]+)")
            datos["Cobertura"] = cobertura if cobertura else "--"

            premio = buscar(texto_completo, r"PREMIO\s*DEL\s*PER[IÍ]ODO\s*\$?\s*([0-9.]+,[0-9]{2})") or \
                     buscar(texto_completo, r"PREMIO\s*\$?\s*([0-9.]+,[0-9]{2})")
            datos["Premio"] = premio if premio else "--"

        filas.append(datos)

    columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura", "Archivo"]
    df = pd.DataFrame(filas, columns=columnas)
    nombre_archivo = "atm.xlsx"
    df.to_excel(nombre_archivo, index=False)

    wb = load_workbook(nombre_archivo)
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

    wb.save(nombre_archivo)
    print(f"✅ Excel generado o actualizado como {nombre_archivo}")
