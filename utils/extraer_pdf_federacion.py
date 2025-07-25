
import pdfplumber
import pandas as pd
import re
import os
import json
from collections import Counter
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill

def procesar_federacion(pdf_paths: list[str]):
    # Cargar marcas
    with open("assets/marcas.json", "r", encoding="utf-8") as f:
        marcas_data = json.load(f)
    lista_marcas = [m["marca"].upper() for m in marcas_data]
    lista_marcas_ordenadas = sorted(lista_marcas, key=lambda x: len(x.split()), reverse=True)

    columnas = ["Marca", "Modelo", "Año", "Suma Asegurada", "Premio", "Cláusula de Ajuste", "Cobertura", "Archivo"]
    filas = []

    for pdf_path in pdf_paths:
        datos = dict.fromkeys(columnas, "--")
        datos["Archivo"] = os.path.basename(pdf_path)

        with pdfplumber.open(pdf_path) as pdf:
            texto = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

            def buscar(patron):
                match = re.search(patron, texto, re.IGNORECASE)
                return match.group(1).strip() if match else "--"

            # Año
            datos["Año"] = buscar(r"Modelo\s+[^\n]*\n.*\b(\d{4})\b")

            # Marca y modelo
            modelo_match = re.search(r"Modelo\s+[^\n]*\n([^\n]+)", texto, re.IGNORECASE)
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

            # Suma asegurada
            suma_matches = re.findall(r"SUMA ASEGURADA\s*\$?\s*([\d.,]+)", texto)
            if suma_matches:
                datos["Suma Asegurada"] = Counter(suma_matches).most_common(1)[0][0]

            # Premio
            premio_match = re.search(r"PREMIO DEL ENDOSO\s*-?\$?\s*(-?[\d.,]+)", texto)
            if premio_match:
                datos["Premio"] = premio_match.group(1).lstrip("-")
            else:
                posibles = re.findall(r"(\d{1,6}[.,]\d{2})", texto)
                candidatos = [p for p in posibles if float(p.replace(".", "").replace(",", ".")) < 999999]
                if candidatos:
                    datos["Premio"] = max(candidatos, key=lambda x: float(x.replace(".", "").replace(",", ".")))

            # Cláusula ajuste
            datos["Cláusula de Ajuste"] = buscar(r"Ajuste Autom[aá]tico.*?(\d{1,3}\s*%)")

        # Cobertura
        # Cargar planes de Federación
        with open("assets/planesFederacion.json", "r", encoding="utf-8") as f:
            planes_federacion = json.load(f)

        # Buscar plan en el texto completo (con tolerancia a número y espacios antes)
        # Buscar plan en el texto completo (normalizando para evitar errores de formato)
        texto_normalizado = texto.upper().replace("\n", " ")
        texto_normalizado = re.sub(r"[\s\-]+", " ", texto_normalizado)

        plan_encontrado = None
        for plan in planes_federacion:
            plan_normalizado = re.sub(r"[\s\-]+", " ", plan.upper())
            if plan_normalizado in texto_normalizado:
                plan_encontrado = plan
                break


        if plan_encontrado:
            datos["Cobertura"] = plan_encontrado
        else:
            # Fallback: buscar texto de cobertura si no encontró un plan
            cobertura_match = re.search(r"RIESGOS CUBIERTOS.*?(\n.*?)(?=\n[A-Z ]+|$)", texto, re.IGNORECASE | re.DOTALL)
            if cobertura_match:
                datos["Cobertura"] = re.sub(r"\s{2,}", " ", cobertura_match.group(1).replace("\n", " ")).strip()


        filas.append({col: datos[col] for col in columnas})

    df = pd.DataFrame(filas)
    nombre_excel = "federacion_patronal.xlsx"
    df.to_excel(nombre_excel, index=False)

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
