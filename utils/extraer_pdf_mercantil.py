import pdfplumber
import pandas as pd
import re
import os
import json
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

def procesar_mercantil(pdfs: list[str]):
    with open("assets/marcas.json", "r", encoding="utf-8") as f:
        marcas_data = json.load(f)

    lista_marcas = sorted(
        [m["marca"].upper() for m in marcas_data],
        key=lambda x: len(x.split()),
        reverse=True
    )

    columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura", "Archivo"]
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

            def buscar(patron, multilinea=True):
                flags = re.MULTILINE | re.IGNORECASE if multilinea else re.IGNORECASE
                resultado = re.search(patron, texto_completo, flags)
                return resultado.group(1).strip() if resultado else ""

            # Marca / Modelo / Año
            marca_modelo = buscar(r"Marca.*?: ([^\n]+)")
            if marca_modelo:
                texto = marca_modelo.upper()
                for marca in lista_marcas:
                    if texto.startswith(marca):
                        datos["Marca"] = marca.title()
                        restante = texto[len(marca):].strip()
                        partes = restante.split()
                        if partes and partes[-1].isdigit():
                            datos["Año"] = partes[-1]
                            partes = partes[:-1]
                        datos["Modelo"] = " ".join(partes).title()
                        break

            # Premio
            premio = (
                buscar(r"PREMIO TOTAL\s+([0-9.]+,[0-9]{2})") or
                buscar(r"Cuota\s+Vto\.Asegu\..*?\n\s*1\s+[0-9.]+\s+([0-9.,]+)") or
                buscar(r"Importe\s*\n([0-9.]+,[0-9]{2})") or
                buscar(r"Premio\s*[:]*\s*\$?\s*([0-9.,]+)")
            )
            datos["Premio"] = premio if premio else "--"

            # Suma Asegurada
            suma_asegurada = (
                buscar(r"Suma Asegurada:\s*\$?\s*([0-9.]+,[0-9]{2})") or
                buscar(r"Suma Asegurada:\s*\$?\s*([0-9.]+)")
            )
            datos["Suma Asegurada"] = suma_asegurada if suma_asegurada else "--"

            # Cobertura
            match_cobertura = re.search(
                r"Coberturas especif\.del riesgo\s*\n(.*?)\n\s*Descripción del Riesgo",
                texto_completo,
                re.DOTALL | re.IGNORECASE
            )
            if match_cobertura:
                cobertura = match_cobertura.group(1).strip()
                datos["Cobertura"] = cobertura

        filas.append({col: datos[col] for col in columnas})

    nombre_archivo = "mercantil.xlsx"
    if os.path.exists(nombre_archivo):
        df_existente = pd.read_excel(nombre_archivo)
        df = pd.concat([df_existente, pd.DataFrame(filas)], ignore_index=True)
    else:
        df = pd.DataFrame(filas)

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
