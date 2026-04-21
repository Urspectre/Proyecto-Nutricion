"""
Union fuzzy de 'Valores_ingredientes' con 'tabla_nutricional_ICBF'
por la columna 'Nombre' / 'nombre'.
 
Diseño del scorer (versión 3):
  - Se LIMPIA el nombre ICBF quitando ruido OCR (|, números, frases como
    "Pulpa sin semillas", "crudo", "cruda", etc.) antes de tokenizar.
  - Los nombres de Valores_ingredientes se transliteran a su forma con
    tilde via MAPA_TILDES, y se conserva la ñ en todo momento.
  - Score = contención de tokens (peso alto) + Jaccard + SequenceMatcher
  - Bonus si el primer token sustantivo del ingrediente está en ICBF.
  - Penalización si no hay ningún token en común.
 
Genera:
  - tabla_unida.xlsx            : pares con match válido
  - ingredientes_sin_match.xlsx : ingredientes sin pareja (buscar en PDF)
"""
 
import re
import difflib
import unicodedata
import pandas as pd
 
# ── Parámetros ajustables ─────────────────────────────────────────────────────
RUTA_VALORES     = r"C:\Users\nikol\Documents\GitHub\Proyecto-Nutricion\Proyecto Nutricion\ICBF\Valores ingredientes.xlsx"
RUTA_ICBF        = r"C:\Users\nikol\Documents\GitHub\Proyecto-Nutricion\Proyecto Nutricion\ICBF\tabla_nutricional_ICBF.xlsx"
UMBRAL_SIMILITUD = 0.55   # 0.0–1.0
SALIDA_UNIDA     = "tabla_unida.xlsx"
SALIDA_SIN_MATCH = "ingredientes_sin_match.xlsx"
# ─────────────────────────────────────────────────────────────────────────────
 
# Palabras de relleno OCR en nombres ICBF que no aportan al match
STOPWORDS_ICBF = {
    "crudo","cruda","cocido","cocida","frita","frito","sin","con","sal",
    "pulpa","semillas","semilla","entera","entero","sin","piel","sin",
    "hueso","cabeza","cuerpo","flor","tallos","hojas","hoja","tallo",
    "médula","corteza","medula","pelado","pelada","picado","picada",
    "precocido","precocida","horneado","horneada","asado","asada",
    "ahumado","ahumada","enlatado","enlatada","congelado","congelada",
    "deshidratado","deshidratada","molido","molida","triturado","triturada",
    "maduro","madura","pintona","pintono","verde","amarillo","amarilla",
    "rojo","roja","blanco","blanca","negro","negra","claro","clara",
    "regular","extra","fino","fina","grueso","gruesa","pequeño","grande",
    "nacional","importado","importada","natural","fresco","fresca",
    "variedad","tipo","de","del","los","las","por","para","una","uno",
    "granulado","granulada","pulido","pulida","integral","enriquecida","enriquecido",
    "fortificada","fortificado","precocida","precocido",
}
 
# Mapa de transliteración: Valores_ingredientes (sin tilde) → forma con tilde
MAPA_TILDES = {
    "pina":          "piña",
    "brocoli":       "brócoli",
    "guanabana":     "guanábana",
    "maracuya":      "maracuyá",
    "melon":         "melón",
    "nispero":       "níspero",
    "platano":       "plátano",
    "limon":         "limón",
    "pimenton":      "pimentón",
    "rabano":        "rábano",
    "azucar":        "azúcar",
    "maiz":          "maíz",
    "frijol":        "fríjol",
    "frijoles":      "fríjoles",
    "salmon":        "salmón",
    "atun":          "atún",
    "arandano":      "arándano",
    "durazno":       "durazno",
    "higado":        "hígado",
    "riñon":         "riñón",
    "piñon":         "piñón",
    "naranja":       "naranja",
    "mandarina":     "mandarina",
    "zanahoria":     "zanahoria",
    "remolacha":     "remolacha",
    "espinaca":      "espinaca",
    "pepino":        "pepino",
    "acelga":        "acelga",
    "lechuga":       "lechuga",
    "berenjena":     "berenjena",
    "repollo":       "repollo",
    "cilantro":      "cilantro",
    "cebolla":       "cebolla",
    "ajo":           "ajo",
    "tomate":        "tomate",
    "habichuela":    "habichuela",
    "arveja":        "arveja",
    "garbanzo":      "garbanzo",
    "coliflor":      "coliflor",
    "ahuyama":       "ahuyama",
    "panela":        "panela",
    "mazorca":       "mazorca",
    "yuca":          "yuca",
    "papa":          "papa",
    "platano":       "plátano",
    "banano":        "banano",
    "mango":         "mango",
    "guayaba":       "guayaba",
    "aguacate":      "aguacate",
    "pitahaya":      "pitahaya",
    "granadilla":    "granadilla",
    "curuba":        "curuba",
    "feijoa":        "feijoa",
    "breva":         "breva",
    "lulo":          "lulo",
    "mora":          "mora",
    "fresa":         "fresa",
    "patilla":       "patilla",
    "papaya":        "papaya",
    "coco":          "coco",
    "arroz":         "arroz",
    "cebada":        "cebada",
    "trigo":         "trigo",
    "arracacha":     "arracacha",
    "costilla":      "costilla",
    "cadera":        "cadera",
    "lomo":          "lomo",
    "pierna":        "pierna",
    "sobrebarriga":  "sobrebarriga",
    "pollo":         "pollo",
    "cerdo":         "cerdo",
    "bagre":         "bagre",
    "cachama":       "cachama",
    "trucha":        "trucha",
    "corvina":       "corvina",
    "tilapia":       "tilapia",
    "mojarra":       "mojarra",
    "sierra":        "sierra",
    "nicuro":        "nicuro",
    "gualajo":       "gualajo",
    "bocachico":     "bocachico",
    "camaron":       "camarón",
    "caracol":       "caracol",
    "doncella":      "doncella",
    "leche":         "leche",
    "queso":         "queso",
    "huevo":         "huevo",
    "sal":           "sal",
    "aceite":        "aceite",
    "margarina":     "margarina",
    "manteca":       "manteca",
    "harina":        "harina",
    "pasta":         "pasta",
    "chocolate":     "chocolate",
    "cafe":          "café",
    "manzana":       "manzana",
    "uva":           "uva",
    "mandarina":     "mandarina",
    "durazno":       "durazno",
    "menudencias":   "menudencias",
    "pechuga":       "pechuga",
    "pernil":        "pernil",
    "uchuva":        "uchuva",
}
 
 
def normalizar(texto):
    """Minúsculas + strip + colapsar espacios. CONSERVA tildes y ñ."""
    if not isinstance(texto, str):
        return ""
    t = texto.lower().strip()
    t = re.sub(r"[\r\n_]+", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t
 
 
def limpiar_icbf(texto_lower):
    """
    Limpia ruido OCR de los nombres ICBF:
    - Quita | y caracteres no-letra/no-espacio que no sean tildes o ñ
    - Quita palabras de stopwords (crudo, cruda, pulpa, semillas, etc.)
    - Colapsa espacios
    """
    # Quitar caracteres que claramente son ruido OCR: | ' " , . ; : / \ 0-9
    t = re.sub(r"[|'\",.;:/\\0-9\(\)\[\]\{\}]", " ", texto_lower)
    # Conservar solo letras (incluyendo tildes y ñ) y espacios
    t = re.sub(r"[^a-záéíóúüñ\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    # Quitar stopwords
    tokens = [tok for tok in t.split() if tok not in STOPWORDS_ICBF and len(tok) >= 3]
    return " ".join(tokens)
 
 
def transliterar(nombre_lower):
    """Aplica MAPA_TILDES token por token al nombre de Valores_ingredientes."""
    tokens = nombre_lower.split()
    resultado = []
    for tok in tokens:
        base = re.sub(r"[^a-záéíóúüñ]", "", tok)
        resultado.append(MAPA_TILDES.get(base, tok))
    return " ".join(resultado)
 
 
def tokenizar(texto):
    """Extrae palabras de >= 3 chars (letras con tildes y ñ)."""
    return set(re.findall(r"[a-záéíóúüñ]{3,}", texto))
 
 
def score_match(nombre_val_tilde, nombre_icbf_limpio):
    """
    Score pensado para la asimetría nombre-corto vs nombre-largo.
 
    Componentes:
      1. Contencion: % de tokens de VAL presentes en ICBF  (peso 0.55)
      2. Jaccard sobre los tokens limpios                   (peso 0.20)
      3. SequenceMatcher de strings completos              (peso 0.20)
      4. Bonus si el primer token sustantivo de VAL está en ICBF (+0.10)
      5. Penalización si NINGÚN token de VAL está en ICBF  (-0.30)
    """
    tok_val  = tokenizar(nombre_val_tilde)
    tok_icbf = tokenizar(nombre_icbf_limpio)
 
    if not tok_val or not tok_icbf:
        return 0.0
 
    interseccion = tok_val & tok_icbf
 
    contencion = len(interseccion) / len(tok_val)
    jaccard    = len(interseccion) / len(tok_val | tok_icbf)
    seq        = difflib.SequenceMatcher(None, nombre_val_tilde, nombre_icbf_limpio).ratio()
 
    primer = re.sub(r"[^a-záéíóúüñ]", "", nombre_val_tilde.split()[0]) if nombre_val_tilde else ""
    bonus  = 0.10 if len(primer) >= 3 and primer in tok_icbf else 0.0
 
    penalizacion = -0.30 if len(interseccion) == 0 else 0.0
 
    score = contencion * 0.55 + jaccard * 0.20 + seq * 0.20 + bonus + penalizacion
    return round(score, 4)
 
 
def mejor_match(nombre_val_orig, opciones_icbf_limpio):
    """Devuelve (indice_en_opciones, score) para el mejor candidato."""
    nombre_lower = normalizar(nombre_val_orig)
    nombre_tilde = transliterar(nombre_lower)
 
    best_idx   = -1
    best_score = 0.0
 
    for idx, opcion in enumerate(opciones_icbf_limpio):
        s = score_match(nombre_tilde, opcion)
        if s > best_score:
            best_score = s
            best_idx   = idx
 
    return best_idx, best_score
 
 
def ajustar_anchos(ws):
    for col_cells in ws.columns:
        max_len = max(
            len(str(c.value)) if c.value is not None else 0
            for c in col_cells
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 45)
 
 
def main():
    df_val  = pd.read_excel(RUTA_VALORES)
    df_icbf = pd.read_excel(RUTA_ICBF)
 
    # Preparar ICBF: normalizar y limpiar ruido OCR para el matching,
    # pero conservar el nombre original para el output
    df_icbf["_nombre_orig_lower"] = df_icbf["nombre"].apply(normalizar)
    df_icbf["_nombre_limpio"]     = df_icbf["_nombre_orig_lower"].apply(limpiar_icbf)
 
    icbf_valido = df_icbf[df_icbf["_nombre_limpio"] != ""].copy().reset_index(drop=True)
    opciones_limpias = icbf_valido["_nombre_limpio"].tolist()
 
    resultados = []
    sin_match  = []
 
    for _, fila in df_val.iterrows():
        nombre_orig = fila["Nombre"]
        idx, score  = mejor_match(nombre_orig, opciones_limpias)
 
        if idx >= 0 and score >= UMBRAL_SIMILITUD:
            fila_icbf = icbf_valido.iloc[idx]
            row = fila.to_dict()
            row["similitud_%"]        = round(score * 100, 1)
            row["nombre_ICBF_match"]  = fila_icbf["nombre"]
            for col in df_icbf.columns:
                if col not in ("nombre", "_nombre_orig_lower", "_nombre_limpio"):
                    row[f"ICBF_{col}"] = fila_icbf[col]
            resultados.append(row)
        else:
            candidato = icbf_valido.iloc[idx]["nombre"] if idx >= 0 else ""
            sin_match.append({
                "Ingrediente_sin_match":    nombre_orig,
                "Mejor_candidato_rechazado": candidato,
                "Similitud_%":              round(score * 100, 1) if idx >= 0 else 0,
            })
 
    df_unida     = pd.DataFrame(resultados)
    df_sin_match = pd.DataFrame(sin_match)
 
    if "similitud_%" in df_unida.columns:
        df_unida.sort_values("similitud_%", ascending=False, inplace=True)
 
    with pd.ExcelWriter(SALIDA_UNIDA, engine="openpyxl") as writer:
        df_unida.to_excel(writer, index=False, sheet_name="Tabla_unida")
        ajustar_anchos(writer.sheets["Tabla_unida"])
 
    with pd.ExcelWriter(SALIDA_SIN_MATCH, engine="openpyxl") as writer:
        df_sin_match.to_excel(writer, index=False, sheet_name="Sin_match")
        ajustar_anchos(writer.sheets["Sin_match"])
 
    total       = len(df_val)
    con_match   = len(df_unida)
    sin_match_n = len(df_sin_match)
 
    print("=" * 65)
    print("  RESULTADO DEL MATCHING")
    print("=" * 65)
    print(f"  Total ingredientes        : {total}")
    print(f"  Con match válido (>={int(UMBRAL_SIMILITUD*100)}%) : {con_match}")
    print(f"  Sin match (buscar manual) : {sin_match_n}")
    print("=" * 65)
    print(f"\n  Tabla unida   -> {SALIDA_UNIDA}")
    print(f"  Sin match     -> {SALIDA_SIN_MATCH}")
 
    if sin_match_n:
        print("\n  Ingredientes que requieren búsqueda manual en el PDF:")
        for row in sin_match:
            nombre = row["Ingrediente_sin_match"]
            cand   = row.get("Mejor_candidato_rechazado", "")
            sim    = row.get("Similitud_%", 0)
            if cand:
                print(f"    - {nombre:<35} | rechazado: '{cand}' ({sim}%)")
            else:
                print(f"    - {nombre}")
 
    if not df_unida.empty and "similitud_%" in df_unida.columns:
        print("\n  --- Matches de menor score (verificar manualmente) ---")
        bajos = df_unida[["Nombre","nombre_ICBF_match","similitud_%"]].sort_values("similitud_%").head(25)
        print(bajos.to_string(index=False))
 
 
if __name__ == "__main__":
    main()