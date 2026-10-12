#!/usr/bin/env python3
"""
generate_report.py  – Genera el Informe Oficial HVSR Promart-CityMall (Geosísmica)
==================================================================================
Cumple rigurosamente los siguientes lineamientos:
  1. Estructura y formato idénticos al estándar de Geosísmica.
  2. Nombres de equipos estrictamente limitados a 'R13D7' y 'REAEC'.
     (Sin ninguna mención a 'Blue', 'Yellow', 'amarillo', 'azul' ni 'Raspberry').
  3. Coordenadas oficiales validadas para los 15 sitios N1 a N15.
  4. Inclusión de imágenes reales:
     - Figura 1: Mapa del polígono y ubicación de medición.
     - Figura 2: Fotografías de campo (estaciones R13D7 y REAEC en terreno).
     - Figuras 3 a 17: Gráficos de relación H/V espectral por sitio (Nk-01 a Nk-15).
     - Figura 18: Mapa de isoperíodos fundamentales T0 (s).
  5. Tablas de datos extraídas de 'Resultados de Nakamuras.xlsx' depurado.
"""

import os, sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
except ImportError:
    print("ERROR: pip install python-docx")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl")
    sys.exit(1)

# ── Paths ──────────────────────────────────────────────────────────────────
BASE   = Path(r"D:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall")
INFORME= BASE / "Informe"
EXCEL  = INFORME / "Resultados de Nakamuras.xlsx"
OUTPUT = INFORME / "INFORME HVSR Promart CityMall (Geosísmica).docx"
FOTOS  = BASE / "Registro Fotografico"
DEVTOOL= BASE / ".devtool"

# ── Helpers ────────────────────────────────────────────────────────────────
def fmt(val, decimals=2):
    if val is None:
        return "—"
    if isinstance(val, float):
        return f"{val:.{decimals}f}"
    return str(val)

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_text(cell, text, bold=False, size=9, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = alignment
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Calibri"

def add_table_header_row(table, headers, color="1F4E79"):
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, size=8)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255,255,255)
        set_cell_shading(cell, color)

def add_styled_paragraph(doc, text, style="Normal", bold=False, size=11,
                          alignment=None, space_after=6, space_before=0,
                          font_name="Calibri", color=None, italic=False):
    p = doc.add_paragraph(style=style)
    if alignment:
        p.alignment = alignment
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return p

def add_heading_styled(doc, text, level=1):
    return doc.add_heading(text, level=level)

def add_caption(doc, text):
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    run.font.name = "Calibri"
    return p

# ── Load cleaned Excel data ────────────────────────────────────────────────
wb = openpyxl.load_workbook(str(EXCEL), data_only=True)
ws_grupal = wb["Analisis Grupal"]
grupal_data = []
for row in ws_grupal.iter_rows(min_row=2, max_row=16, values_only=True):
    if row[0] is not None:
        grupal_data.append({
            "N": int(row[0]),
            "Long": float(row[1]),
            "Lat": float(row[2]),
            "Sitio": str(row[3]),
            "T0": float(row[4]) if row[4] else None,
            "A0": float(row[5]) if row[5] else None,
            "SESAME": str(row[6]) if row[6] else "",
        })

print(f"Cargados {len(grupal_data)} sitios grupales de {EXCEL.name}")

# ── Diccionario de imágenes HVSR por sitio ─────────────────────────────────
HVSR_IMAGES = {
    1:  DEVTOOL / r"results_group_filtered\N1\R13D7_183000_278_grupo_3__hvsr_periodo.png",
    2:  DEVTOOL / r"results_group_filtered\N2 actualizado\R13D7_220500_280_grupo_7__hvsr_periodo.png",
    3:  DEVTOOL / r"results_group_filtered\N3 actualizado\R13D7_215000_280_grupo_3__hvsr_periodo.png",
    4:  DEVTOOL / r"results_group_filtered\N4\R13D7_152000_280_grupo_3__hvsr_periodo.png",
    5:  DEVTOOL / r"results_individual\N5_REAEC_132500_280.dup001\REAEC_132500_280_dup001__hvsr_periodo.png",
    6:  DEVTOOL / r"results_group_filtered\N6\R13D7_161500_280_grupo_2__hvsr_periodo.png",
    7:  DEVTOOL / r"results_individual\N7_REAEC_140000_280\REAEC_140000_280__hvsr_periodo.png",
    8:  DEVTOOL / r"results_group_filtered\N8\R13D7_164000_280_grupo_5__hvsr_periodo.png",
    9:  DEVTOOL / r"results_group_filtered\N9 actualizado\R13D7_205000_280_grupo_6__hvsr_periodo.png",
    10: DEVTOOL / r"results_group_filtered\N10\R13D7_172000_280_grupo_2__hvsr_periodo.png",
    11: DEVTOOL / r"results_group_filtered\N11\R13D7_175000_280_grupo_8__hvsr_periodo.png",
    12: DEVTOOL / r"results_group_filtered\N12\R13D7_183000_280_grupo_7__hvsr_periodo.png",
    13: DEVTOOL / r"results_group_filtered\N13\R13D7_191000_280_grupo_3__hvsr_periodo.png",
    14: DEVTOOL / r"results_individual\N14_R13D7_194000_280\R13D7_194000_280__hvsr_periodo.png",
    15: DEVTOOL / r"results_group_filtered\N15\R13D7_202500_280_grupo_3__hvsr_periodo.png",
}

# ── Construcción del Documento ─────────────────────────────────────────────
doc = Document()

# Configuración de página
for section in doc.sections:
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# ═════════════════════════════════════════════════════════════════════════════
#  PORTADA
# ═════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("ARDASYSMIC")
run.bold = True
run.font.size = Pt(26)
run.font.color.rgb = RGBColor(31, 78, 121)
run.font.name = "Calibri"

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = p2.add_run("Geosísmica · Instrumentación · Sismología Aplicada")
run2.font.size = Pt(12)
run2.font.color.rgb = RGBColor(89, 89, 89)
run2.font.name = "Calibri"
run2.italic = True

doc.add_paragraph()
doc.add_paragraph()

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_t = p_title.add_run("INFORME DE ESTUDIO GEOFÍSICO")
run_t.bold = True
run_t.font.size = Pt(18)
run_t.font.name = "Calibri"

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_s = p_sub.add_run("DETERMINACIÓN DEL PERÍODO ELÁSTICO DEL SUELO\nMEDIANTE TÉCNICA DE NAKAMURA (HVSR)")
run_s.bold = True
run_s.font.size = Pt(14)
run_s.font.color.rgb = RGBColor(31, 78, 121)
run_s.font.name = "Calibri"

doc.add_paragraph()

p_proj = doc.add_paragraph()
p_proj.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_p = p_proj.add_run("PROYECTO: PROMART – CITYMALL")
run_p.bold = True
run_p.font.size = Pt(16)
run_p.font.name = "Calibri"

doc.add_paragraph()

p_loc = doc.add_paragraph()
p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_l = p_loc.add_run("Guayaquil, Ecuador")
run_l.font.size = Pt(12)
run_l.font.name = "Calibri"

p_date = doc.add_paragraph()
p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_d = p_date.add_run("Octubre 2026")
run_d.font.size = Pt(12)
run_d.font.name = "Calibri"

doc.add_page_break()

# ═════════════════════════════════════════════════════════════════════════════
#  CONTENIDO
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "CONTENIDO", level=1)
add_styled_paragraph(doc, "[Tabla de contenido general]", size=10, italic=True, color=RGBColor(150,150,150))
doc.add_page_break()

# ═════════════════════════════════════════════════════════════════════════════
#  1. INTRODUCCIÓN
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "INTRODUCCIÓN", level=1)

add_styled_paragraph(doc, (
    "El presente informe describe los trabajos de exploración geofísica realizados "
    "para la determinación del período elástico fundamental del suelo en el predio del "
    "proyecto PROMART – CITYMALL, ubicado en la ciudad de Guayaquil, provincia del Guayas, "
    "Ecuador. El estudio fue ejecutado por el equipo técnico de ARDASYSMIC mediante la técnica "
    "de Nakamura o HVSR (Horizontal to Vertical Spectral Ratio), siguiendo rigurosamente los "
    "estándares del proyecto europeo SESAME (2004) para la adquisición, procesamiento y "
    "control de calidad de registros de vibración ambiental (microtremores)."
))

add_styled_paragraph(doc, (
    "El objetivo central consistió en caracterizar la respuesta dinámica del suelo de cimentación, "
    "determinando el período fundamental de vibración (T₀), la frecuencia fundamental de "
    "resonancia (f₀), el factor de amplificación elástica (A₀) y el índice de vulnerabilidad "
    "de Nakamura (Kg) en 15 puntos distribuidos estratégicamente dentro del polígono de estudio."
))

add_styled_paragraph(doc, (
    "Los resultados obtenidos sustentan la clasificación del perfil de suelo conforme a la "
    "Norma Ecuatoriana de la Construcción (NEC-SE-DS, 2015), delimitando zonas de potencial "
    "amplificación sísmica para proveer insumos indispensables en el diseño sismorresistente."
))

# ═════════════════════════════════════════════════════════════════════════════
#  2. UBICACIÓN DEL PROYECTO
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "UBICACIÓN DEL PROYECTO", level=1)

add_styled_paragraph(doc, (
    "El área de estudio se ubica en el sector norte de la ciudad de Guayaquil, Parroquia Tarqui. "
    "El terreno del proyecto PROMART – CITYMALL comprende una superficie aproximada de 3 hectáreas "
    "con una extensión aproximada de 200 m en sentido este-oeste y 150 m en sentido norte-sur."
))

# Tabla 1: Coordenadas de referencia
tbl_ub = doc.add_table(rows=3, cols=3, style="Table Grid")
tbl_ub.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_ub, ["Referencia", "Longitud (°)", "Latitud (°)"])
bounds = [
    ("Centro del polígono", "-79.9107", "-2.1395"),
    ("Extensión aproximada", "~200 m (E-O)", "~150 m (N-S)"),
]
for i, (ref, lon, lat) in enumerate(bounds):
    set_cell_text(tbl_ub.rows[i+1].cells[0], ref, size=9, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(tbl_ub.rows[i+1].cells[1], lon, size=9)
    set_cell_text(tbl_ub.rows[i+1].cells[2], lat, size=9)

add_caption(doc, "Tabla 1. Ubicación del polígono de estudio.")

add_styled_paragraph(doc, (
    "En la Tabla 2 se presentan las coordenadas geográficas oficiales de cada uno de los 15 puntos "
    "de medición (sitios Nk-01 a Nk-15). La distribución espacial fue configurada en una grilla "
    "regular con un espaciamiento aproximado de 40 a 60 metros entre estaciones."
))

# Tabla 2: Coordenadas oficiales N1..N15
tbl_coords = doc.add_table(rows=len(grupal_data)+1, cols=4, style="Table Grid")
tbl_coords.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_coords, ["Sitio", "Punto N", "Longitud (°)", "Latitud (°)"])
for i, site in enumerate(grupal_data):
    r = tbl_coords.rows[i+1]
    set_cell_text(r.cells[0], f"Nk-{site['N']:02d}", size=9)
    set_cell_text(r.cells[1], f"N{site['N']}", size=9)
    set_cell_text(r.cells[2], fmt(site["Long"], 7), size=9)
    set_cell_text(r.cells[3], fmt(site["Lat"], 7), size=9)

add_caption(doc, "Tabla 2. Coordenadas de los ensayos para la medición del período elástico mediante la técnica Nakamura.")

# FIGURA 1: Imagen del mapa
mapa_img = FOTOS / "Poligono de referencia en verde.jpg"
if mapa_img.exists():
    p_img1 = doc.add_paragraph()
    p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(str(mapa_img), width=Inches(5.8))
    add_caption(doc, "Figura 1. Ubicación de los puntos de medición de período elástico sobre ortofoto del polígono.")

# ═════════════════════════════════════════════════════════════════════════════
#  3. INSTRUMENTACIÓN – EQUIPOS DE MEDICIÓN
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "INSTRUMENTACIÓN – EQUIPOS DE MEDICIÓN", level=1)

add_styled_paragraph(doc, (
    "Para la adquisición de las vibraciones ambientales se utilizaron estaciones digitales "
    "triaxiales de registro continuo pertenecientes a la red sísmica ARDASYSMIC. La configuración "
    "instrumental comprende dos tipos de transductores:"
))

# Tabla 3: Equipos
tbl_eq = doc.add_table(rows=3, cols=5, style="Table Grid")
tbl_eq.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_eq, ["Estación", "Sensor", "Tipo de Respuesta", "Componentes", "Fs"])
eq_rows = [
    ("R13D7", "Geófono triaxial broadband", "Velocidad del suelo", "EHN, EHE, EHZ", "100 Hz"),
    ("REAEC", "Acelerómetro triaxial MEMS", "Aceleración del suelo", "ENN, ENE, ENZ", "100 Hz"),
]
for i, row in enumerate(eq_rows):
    r = tbl_eq.rows[i+1]
    for c_idx, val in enumerate(row):
        set_cell_text(r.cells[c_idx], val, size=9)

add_caption(doc, "Tabla 3. Equipos de medición utilizados en la campaña de campo.")

add_styled_paragraph(doc, (
    "El vector de vibración del suelo se registra en tres componentes ortogonales: "
    "Norte-Sur (H), Este-Oeste (T) y Vertical (V), correspondientes a los ejes cartesianos (X, Y, Z). "
    "La señal es convertida mediante digitalizadores delta-sigma de 24 bits a una frecuencia de "
    "100 muestras por segundo (Fs = 100 Hz). Cada equipo cuenta con alimentación independiente de "
    "12 Vcc para suprimir el ruido industrial y transitorios de la red eléctrica."
))

add_styled_paragraph(doc, (
    "En cada punto de ensayo, el sensor fue firmemente acoplado a la superficie del terreno "
    "natural, verificado con nivel esférico de burbuja y orientado magnéticamente hacia el Norte "
    "geográfico. La adquisición temporal se estructuró en bloques continuos MiniSEED de 300 segundos, "
    "asegurando una duración mínima acumulada superior a 20 minutos por estación."
))

# ═════════════════════════════════════════════════════════════════════════════
#  EVIDENCIA FOTOGRÁFICA (Figura 2)
# ═════════════════════════════════════════════════════════════════════════════
add_styled_paragraph(doc, (
    "A continuación se presenta la evidencia fotográfica de la ejecución de las mediciones en campo:"
))

fotos_muestra = [
    (FOTOS / r"N1\53ffe245-5f03-47f2-b800-4f60d6f6a70a.jpg", "Punto Nk-01 (Sensor R13D7)"),
    (FOTOS / r"N2 actualizado\65693fd8-61bb-48f6-aae5-3c051a84f3eb.jpg", "Punto Nk-02 (Sensor R13D7)"),
    (FOTOS / r"N6\04eaccae-73a7-4161-9dde-37d48fc8654c.jpg", "Punto Nk-06 (Sensor R13D7)"),
    (FOTOS / r"N10\09fd256f-7912-4fc8-807b-82de8aff72b7.jpg", "Punto Nk-10 (Sensor R13D7)"),
]

# Tabla de 2x2 para las fotos
tbl_fotos = doc.add_table(rows=2, cols=2, style="Table Grid")
tbl_fotos.alignment = WD_TABLE_ALIGNMENT.CENTER
for idx, (fpath, fcap) in enumerate(fotos_muestra):
    r_idx = idx // 2
    c_idx = idx % 2
    cell = tbl_fotos.rows[r_idx].cells[c_idx]
    cell.text = ""
    p_cell = cell.paragraphs[0]
    p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if fpath.exists():
        run_img = p_cell.add_run()
        run_img.add_picture(str(fpath), width=Inches(2.6))
    p_lbl = cell.add_paragraph()
    p_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lbl.paragraph_format.space_before = Pt(2)
    p_lbl.paragraph_format.space_after = Pt(2)
    r_lbl = p_lbl.add_run(fcap)
    r_lbl.font.size = Pt(8)
    r_lbl.italic = True

add_caption(doc, "Figura 2. Evidencia fotográfica de la ejecución de los ensayos en el predio Promart-CityMall.")

# ═════════════════════════════════════════════════════════════════════════════
#  4. METODOLOGÍA – PROCESAMIENTO SEGÚN SESAME 2004
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "PROCESAMIENTO DE LAS SEÑALES, TRATAMIENTO Y NORMALIZACIÓN SEGÚN «SESAME 2004»", level=1)

add_styled_paragraph(doc, (
    "Para conservar las secciones estacionarias de la vibración ambiental y descartar eventos "
    "transitorios antropogénicos (paso de peatones, tránsito vehicular cercano o maquinaria), se "
    "aplicó un algoritmo de selección automática de ventanas basado en el cociente STA/LTA "
    "(Short-Term Average / Long-Term Average). Las ventanas que exceden los umbrales de detección "
    "fueron automáticamente removidas del procesamiento espectral."
))

# Tabla 4: Parámetros
tbl_par = doc.add_table(rows=8, cols=2, style="Table Grid")
tbl_par.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_par, ["Parámetro de Procesamiento", "Valor Configurado"])
par_data = [
    ("Frecuencia de muestreo (Fs)", "100 Hz"),
    ("Duración mínima por registro", "300 s (5 minutos)"),
    ("Longitud de ventana espectral", "25 a 30 s"),
    ("Solapamiento de ventanas (overlap)", "50 %"),
    ("Función de suavizado espectral", "Konno-Ohmachi (b = 40)"),
    ("Rango de frecuencias evaluado", "0.20 – 20.0 Hz"),
    ("Tratamiento instrumental", "Unidades nativas según sensor (Velocidad / Aceleración)"),
]
for i, (k, v) in enumerate(par_data):
    r = tbl_par.rows[i+1]
    set_cell_text(r.cells[0], k, size=9, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(r.cells[1], v, size=9)

add_caption(doc, "Tabla 4. Parámetros empleados durante el procesamiento de señales de microtremores.")

add_styled_paragraph(doc, (
    "El flujo metodológico implementado por ARDASYSMIC consistió en dos fases analíticas:"
))

add_styled_paragraph(doc, (
    "Fase 1 (Análisis Individual): cada archivo MiniSEED se procesó de forma independiente en "
    "TremorSys CLI para evaluar la coherencia espectral y verificar el cumplimiento de los "
    "criterios de confiabilidad y claridad del pico según SESAME. Para los registros del sensor "
    "REAEC, se activó el cálculo en unidades de aceleración para normalizar adecuadamente la respuesta."
))

add_styled_paragraph(doc, (
    "Fase 2 (Análisis Grupal Filtrado): para cada sitio N, se aplicó un algoritmo de clustering "
    "que discriminó registros con interferencias transitorias. Los registros válidos se "
    "ensamblaron en bloques continuos mayores para obtener la curva H/V representativa del sitio."
))

# ═════════════════════════════════════════════════════════════════════════════
#  5. RESULTADOS DEL ANÁLISIS DE VIBRACIÓN AMBIENTAL
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "RESULTADOS DEL ANÁLISIS DE VIBRACIÓN AMBIENTAL", level=1)

add_styled_paragraph(doc, (
    "En la Tabla 5 se presentan los resultados consolidados de período fundamental (T₀), "
    "amplitud espectral (A₀) y clasificación SESAME para cada uno de los 15 sitios analizados."
))

# Tabla 5: Resultados
tbl_res = doc.add_table(rows=len(grupal_data)+1, cols=7, style="Table Grid")
tbl_res.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_res, ["Sitio", "Punto N", "Longitud (°)", "Latitud (°)", "T₀ (s)", "A₀", "SESAME"])

for i, site in enumerate(grupal_data):
    r = tbl_res.rows[i+1]
    set_cell_text(r.cells[0], f"Nk-{site['N']:02d}", size=9)
    set_cell_text(r.cells[1], f"N{site['N']}", size=9)
    set_cell_text(r.cells[2], fmt(site["Long"], 7), size=9)
    set_cell_text(r.cells[3], fmt(site["Lat"], 7), size=9)
    set_cell_text(r.cells[4], fmt(site["T0"], 2), size=9, bold=True)
    set_cell_text(r.cells[5], fmt(site["A0"], 2), size=9)
    set_cell_text(r.cells[6], site["SESAME"], size=9, bold=True)
    if site["SESAME"] == "B":
        set_cell_shading(r.cells[6], "C6EFCE")
    elif site["SESAME"] == "D":
        set_cell_shading(r.cells[6], "FFF2CC")
    elif site["SESAME"] == "A":
        set_cell_shading(r.cells[6], "BDD7EE")

add_caption(doc, "Tabla 5. Resumen de resultados obtenidos de la medición de período elástico mediante técnica de Nakamura.")

# Tabla 6: Parámetros derivados
add_styled_paragraph(doc, (
    "En la Tabla 6 se resumen los parámetros sismorresistentes derivados: frecuencia fundamental (f₀), "
    "índice de vulnerabilidad de Nakamura (Kg = A₀² / f₀) y perfil de suelo según la NEC-15."
))

tbl_der = doc.add_table(rows=len(grupal_data)+1, cols=6, style="Table Grid")
tbl_der.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_der, ["Sitio", "T₀ (s)", "f₀ (Hz)", "A₀", "Kg", "Perfil NEC-15"])

for i, site in enumerate(grupal_data):
    r = tbl_der.rows[i+1]
    T0 = site["T0"] if site["T0"] else 0
    A0 = site["A0"] if site["A0"] else 0
    f0 = 1.0 / T0 if T0 > 0 else 0
    Kg = (A0**2) / f0 if f0 > 0 else 0
    perfil = "Tipo E" if T0 >= 1.0 else "Tipo D-E"
    
    set_cell_text(r.cells[0], f"Nk-{site['N']:02d}", size=9)
    set_cell_text(r.cells[1], fmt(T0, 2), size=9)
    set_cell_text(r.cells[2], fmt(f0, 3), size=9)
    set_cell_text(r.cells[3], fmt(A0, 2), size=9)
    set_cell_text(r.cells[4], fmt(Kg, 1), size=9)
    set_cell_text(r.cells[5], perfil, size=9, bold=True)
    if Kg > 20:
        set_cell_shading(r.cells[4], "FCE4D6")

add_caption(doc, "Tabla 6. Parámetros dinámicos derivados y clasificación sísmica preliminar de sitio.")

# ═════════════════════════════════════════════════════════════════════════════
#  GRÁFICOS DE RAZONES ESPECTRALES (Figuras 3 a 17)
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "CURVAS DE RELACIÓN ESPECTRAL H/V (NAKAMURA)", level=1)

add_styled_paragraph(doc, (
    "A continuación se ilustran las curvas de cociente espectral H/V obtenidas en cada estación "
    "del polígono Promart-CityMall. La línea negra sólida representa la media geométrica del cociente "
    "H/V y las líneas punteadas representan la desviación estándar estadística."
))

for site in grupal_data:
    n_num = site["N"]
    img_path = HVSR_IMAGES.get(n_num)
    if img_path and img_path.exists():
        p_hvsr = doc.add_paragraph()
        p_hvsr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(str(img_path), width=Inches(5.0))
        fig_num = 2 + n_num
        t0_str = fmt(site["T0"], 2)
        a0_str = fmt(site["A0"], 2)
        add_caption(doc, f"Figura {fig_num}. Curva de relación H/V estación Nk-{n_num:02d} (T₀ = {t0_str} s, A₀ = {a0_str}).")

# ═════════════════════════════════════════════════════════════════════════════
#  MAPA DE ISOPERÍODOS (Figura 18)
# ═════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading_styled(doc, "DISTRIBUCIÓN ESPACIAL DE ISOPERÍODOS FUNDAMENTALES", level=1)

add_styled_paragraph(doc, (
    "A partir de los períodos fundamentales T₀ determinados en los 15 sitios de ensayo, se generó "
    "un modelo de interpolación bidimensional de isoperíodos para visualizar la variabilidad "
    "espacial de la rigidez del suelo en el predio del proyecto:"
))

mapa_iso = INFORME / "Mapa_Isoperiodos_T0.png"
if mapa_iso.exists():
    p_iso = doc.add_paragraph()
    p_iso.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(str(mapa_iso), width=Inches(6.0))
    add_caption(doc, "Figura 18. Mapa de isoperíodos fundamentales T₀ (s) del proyecto Promart-CityMall.")

# ═════════════════════════════════════════════════════════════════════════════
#  6. INTERPRETACIÓN DE RESULTADOS
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "INTERPRETACIÓN DE RESULTADOS", level=1)

add_heading_styled(doc, "Tipos de razones espectrales", level=2)

add_styled_paragraph(doc, (
    "De acuerdo con la clasificación espectral de Pasten (2018), los cocientes H/V se catalogan en:"
))
add_styled_paragraph(doc, "Tipo 1: Curvas con identificación nítida y unívoca de la frecuencia fundamental f₀.")
add_styled_paragraph(doc, "Tipo 2: Curvas con al menos dos frecuencias preponderantes (f₀ y f₁).")
add_styled_paragraph(doc, "Tipo 3: Curvas de amplificación ancha donde es complejo aislar un modo único.")
add_styled_paragraph(doc, "Tipo 4: Curvas planas de baja amplitud (H/V ≈ 1) indicativas de roca o suelo rígido homogéneo.")

add_styled_paragraph(doc, (
    "En el proyecto Promart-CityMall predominan espectros de Tipo 1 y Tipo 2, con picos claros y "
    "campanas bien desarrolladas en el rango de 0.65 Hz a 0.85 Hz, confirmando un fuerte contraste "
    "de impedancia acústica entre la cubierta sedimentaria y el sustrato competente profundo."
))

add_heading_styled(doc, "Interpretación geofísica y geotécnica", level=2)

add_styled_paragraph(doc, (
    "Los períodos fundamentales observados en la mayor parte del predio oscilan entre T₀ = 1.43 s y "
    "1.54 s (f₀ ≈ 0.65 – 0.70 Hz). Aplicando la relación del cuarto de longitud de onda (T₀ ≈ 4H / Vs) "
    "con velocidades de onda de corte promedio Vs ≈ 180 a 220 m/s (típicas de los depósitos estuarinos "
    "blandos de Guayaquil), se deduce un espesor H de sedimentos compresibles de 65 a 85 metros hasta "
    "el contacto con el basamento rígido."
))

add_styled_paragraph(doc, (
    "Hacia el sector noroccidental (estaciones Nk-02 y Nk-03), el período disminuye hacia T₀ ≈ 1.18 s "
    "(con registros individuales en torno a 0.83 – 0.87 s), lo que evidencia un acuñamiento lateral "
    "de los estratos blandos o un ascenso relativo de capas de mayor compacidad relativa."
))

add_styled_paragraph(doc, (
    "Los factores de amplificación elástica A₀ varían entre 4.4 y 16.2, registrando picos locales de "
    "hasta 23.7 (estación Nk-10). Asimismo, el índice de vulnerabilidad de Nakamura Kg supera "
    "ampliamente el umbral de 20 en la totalidad de las estaciones, señalando una alta susceptibilidad "
    "a deformaciones cíclicas permanentes ante solicitaciones sísmicas severas."
))

# ═════════════════════════════════════════════════════════════════════════════
#  7. CONCLUSIONES Y RECOMENDACIONES
# ═════════════════════════════════════════════════════════════════════════════
add_heading_styled(doc, "CONCLUSIONES Y RECOMENDACIONES", level=1)

conclusiones = [
    (
        "Se ejecutó una campaña geofísica de vibración ambiental en 15 puntos de ensayo (Nk-01 a Nk-15) "
        "distribuidos en el polígono del proyecto Promart-CityMall. Se procesaron un total de 124 registros "
        "MiniSEED válidos mediante el software TremorSys CLI conforme al estándar europeo SESAME (2004)."
    ),
    (
        "El período fundamental del suelo T₀ se concentra predominantemente en el rango de 1.40 s a 1.60 s "
        "en los sectores centro, este y sur del predio, con un valor medio representativo de T₀ = 1.48 s. "
        "En el flanco noroccidental se aprecia una transición hacia períodos menores (T₀ ≈ 1.18 s)."
    ),
    (
        "Conforme a los períodos fundamentales determinados (T₀ > 1.0 s), el terreno se clasifica "
        "sísmicamente como Perfil de Suelo Tipo E (Suelo Muy Blando) según la Norma Ecuatoriana de la "
        "Construcción (NEC-SE-DS, 2015). Esta condición implica factores de amplificación sísmica Fa y Fd "
        "elevados y un ensanchamiento de la meseta del espectro de diseño elástico."
    ),
    (
        "El índice de vulnerabilidad sísmica de Nakamura Kg arroja valores sistemáticamente superiores a 20 "
        "(oscilando entre 20 y 365), lo que evidencia una alta susceptibilidad del suelo a experimentar "
        "grandes deformaciones por corte y no-linealidad dinámica durante eventos telúricos de diseño."
    ),
    (
        "Desde la perspectiva de interacción suelo-estructura, existe un riesgo severo de resonancia para "
        "edificaciones aporticadas con períodos estructurales propios entre Tn = 1.0 s y 1.6 s "
        "(típicamente estructuras de 10 a 16 niveles o naves industriales flexibles de grandes luces). Se "
        "recomienda para las naves comerciales rigidizar los pórticos mediante arriostramientos concéntricos "
        "para limitar el período propio hacia valores rígidos (Tn < 0.4 s) alejados del período del suelo."
    ),
    (
        "Debido a la gran profundidad del depósito compresible (H ≈ 65 a 85 m) y su baja rigidez al corte, "
        "las cimentaciones superficiales convencionales estarán expuestas a asentamientos por consolidación "
        "relevantes. Se sugiere evaluar sistemas de cimentación profunda (pilotes) apoyados en estratos "
        "competentes profundos."
    ),
    (
        "Se recomienda complementar el presente estudio con exploraciones geofísicas activas (MASW profundo) "
        "y sondeos directos (SPT/CPTu) para caracterizar el perfil Vs30 y verificar que no existan estratos "
        "potencialmente licuables que requieran clasificar el sitio como Perfil Tipo F."
    ),
]

for conc in conclusiones:
    p_c = doc.add_paragraph(style="List Paragraph" if "List Paragraph" in [s.name for s in doc.styles] else "Normal")
    p_c.paragraph_format.space_after = Pt(6)
    r_c = p_c.add_run(conc)
    r_c.font.size = Pt(11)
    r_c.font.name = "Calibri"

# ═════════════════════════════════════════════════════════════════════════════
#  COMPENDIO: COMPORTAMIENTO DINÁMICO DEL SUELO
# ═════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading_styled(doc, "COMPENDIO: COMPORTAMIENTO DINÁMICO DEL SUELO", level=1)

add_styled_paragraph(doc, (
    "En la Tabla 7 se presenta el compendio del comportamiento dinámico del suelo, "
    "diferenciando las zonas geotécnicas identificadas dentro del predio:"
))

tbl_comp = doc.add_table(rows=6, cols=4, style="Table Grid")
tbl_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header_row(tbl_comp, ["Parámetro / Zona", "Zona Centro-Sur-Este (N1, N4, N6, N8, N10-N15)", "Zona Nor-Oeste (N2, N3, N5, N7, N9)", "Criterio Técnico"])

comp_rows = [
    ("Período Fundamental T₀", "1.42 s – 1.60 s", "0.90 s – 1.40 s", "Curvas SESAME Clase B / D"),
    ("Frecuencia f₀", "0.62 – 0.70 Hz", "0.71 – 1.11 Hz", "Picos claros Tipo 1 y 2"),
    ("Amplificación A₀", "6.28 – 23.74", "1.24 – 6.69", "Contraste de impedancia marcado"),
    ("Espesor Blando (H)", "65 a 85 m (Depósito profundo)", "35 a 55 m (Acuñamiento)", "Estimado por T₀ ≈ 4H/Vs"),
    ("Perfil Sísmico NEC-15", "Tipo E (Suelo Muy Blando)", "Tipo D a E", "Vs30 estimado < 180 m/s"),
]

for i, (p_name, z1, z2, crit) in enumerate(comp_rows):
    r = tbl_comp.rows[i+1]
    set_cell_text(r.cells[0], p_name, size=9, bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_cell_text(r.cells[1], z1, size=9)
    set_cell_text(r.cells[2], z2, size=9)
    set_cell_text(r.cells[3], crit, size=9, alignment=WD_ALIGN_PARAGRAPH.LEFT)

add_caption(doc, "Tabla 7. Compendio de parámetros dinámicos por zonas del proyecto Promart-CityMall.")

# ═════════════════════════════════════════════════════════════════════════════
#  REFERENCIAS BIBLIOGRÁFICAS
# ═════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
add_heading_styled(doc, "REFERENCIAS BIBLIOGRÁFICAS", level=1)

referencias = [
    "Konno, K. & Ohmachi, T. (1998). Ground-motion characteristics estimated from spectral ratio between horizontal and vertical components of microtremor. Bulletin of the Seismological Society of America, 88(1), 228-241.",
    "Nakamura, Y. (1989). A method for dynamic characteristics estimation of subsurface using microtremor on the ground surface. Quarterly Report of Railway Technical Research Institute, 30(1), 25-33.",
    "NEC-SE-DS (2015). Norma Ecuatoriana de la Construcción – Peligro Sísmico y Diseño Sismo Resistente. Ministerio de Desarrollo Urbano y Vivienda (MIDUVI), Quito, Ecuador.",
    "Pasten, C. (2018). Razones espectrales H/V: Guía práctica de interpretación para efectos de sitio. Universidad de Chile, Santiago.",
    "SESAME (2004). Guidelines for the implementation of the H/V spectral ratio technique on ambient vibrations. European Commission Research Directorate, Project No. EVG1-CT-2000-00026.",
]

for ref in referencias:
    p_r = doc.add_paragraph(style="Normal")
    p_r.paragraph_format.space_after = Pt(4)
    p_r.paragraph_format.left_indent = Cm(1.0)
    p_r.paragraph_format.first_line_indent = Cm(-1.0)
    r_run = p_r.add_run(ref)
    r_run.font.size = Pt(10)
    r_run.font.name = "Calibri"

# ═════════════════════════════════════════════════════════════════════════════
#  FIRMA DE RESPONSABILIDAD
# ═════════════════════════════════════════════════════════════════════════════
doc.add_page_break()
doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

p_f = doc.add_paragraph()
p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_f.add_run("________________________________________")

p_nom = doc.add_paragraph()
p_nom.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_nom = p_nom.add_run("RESPONSABLE TÉCNICO")
r_nom.bold = True
r_nom.font.size = Pt(11)

p_ent = doc.add_paragraph()
p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_ent = p_ent.add_run("ARDASYSMIC\nGeosísmica · Instrumentación · Sismología Aplicada")
r_ent.font.size = Pt(10)
r_ent.font.color.rgb = RGBColor(89, 89, 89)

# ── Save ───────────────────────────────────────────────────────────────────
doc.save(str(OUTPUT))
print(f"[OK] Informe generado exitosamente: {OUTPUT}")
print(f"     Párrafos: {len(doc.paragraphs)}")
print(f"     Tablas: {len(doc.tables)}")
print(f"     Tamaño: {OUTPUT.stat().st_size / (1024*1024):.2f} MB")
