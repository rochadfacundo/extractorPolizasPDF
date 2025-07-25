# 📄 Extractor de Pólizas PDF

Este proyecto permite extraer automáticamente información estructurada desde archivos PDF de pólizas de seguros, generando planillas Excel con los datos relevantes.

Diseñado para trabajar con múltiples compañías aseguradoras, cada una con su propio script y lógica de extracción.

---

## 🧩 Características

- Extracción de:
  - Marca
  - Modelo
  - Año
  - Suma Asegurada
  - Premio
  - Cláusula de Ajuste
  - Cobertura (Plan)
  - Nombre del archivo fuente

- Soporte para múltiples compañías con formatos distintos.
- Lógica adaptable por compañía.
- Salida en Excel (`.xlsx`) con formato visual amigable.

---

# ✅ Requisitos

- Python 3.9 o superior
- Dependencias:

```bash
pip install pdfplumber pandas openpyxl
