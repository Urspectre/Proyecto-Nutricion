from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import re

# Configuración
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

pdf_path = r"C:\Users\nikol\Documents\GitHub\Proyecto-Nutricion\Proyecto Nutricion\ICBF\ICBF_unlocked.pdf"

# Convertir página del PDF a imagen
images = convert_from_path(
    pdf_path,
    first_page=29,
    last_page=29,
    poppler_path=r"C:\poppler-25.12.0\Library\bin"
)

# Extraer texto con OCR
text = pytesseract.image_to_string(images[0], lang='spa')

# Limpiar líneas vacías
lines = [l.strip() for l in text.split("\n") if l.strip()]

# Encontrar índices de las columnas
codigo_idx = lines.index("Código")
grupo_idx = lines.index("Grupo de Alimentos")

# Extraer valores de cada columna
codigos = lines[codigo_idx + 1:grupo_idx]
grupos = lines[grupo_idx + 1:]

# Filtrar solo códigos válidos (letras mayúsculas y números, 3-5 caracteres)
codigos = [c for c in codigos if re.match(r"[A-Z0-9]{3,5}", c)]

# Igualar longitudes
min_len = min(len(codigos), len(grupos))
codigos = codigos[:min_len]
grupos = grupos[:min_len]

# Crear DataFrame
df = pd.DataFrame({
    "codigo": codigos,
    "grupo": grupos
})

print(df.head())