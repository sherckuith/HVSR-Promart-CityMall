"""
clean_excels.py
Depuración integral de los 3 archivos Excel del proyecto:
  1. Bitacora.xlsx
  2. Bitacora de Registros Nakamuras.xlsx
  3. Informe/Resultados de Nakamuras.xlsx
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime

# ── 1. COORDENADAS OFICIALES ACTIVAS (N1 a N15) ──────────────────────────
COORDENADAS = {
    1:  (-2.1389506, -79.9109948),
    2:  (-2.1391298, -79.9107470),
    3:  (-2.1392937, -79.9104110),
    4:  (-2.1391645, -79.9102661),
    5:  (-2.1398310, -79.9095325),
    6:  (-2.1398780, -79.9097595),
    7:  (-2.1397941, -79.9103107),
    8:  (-2.1387928, -79.9104797),
    9:  (-2.1394408, -79.9114019),
    10: (-2.1395091, -79.9115469),
    11: (-2.1394707, -79.9115146),
    12: (-2.1391612, -79.9112014),
    13: (-2.1399517, -79.9107717),
    14: (-2.1402488, -79.9103401),
    15: (-2.1402582, -79.9101212),
}

DISCARDED_MSEED = {
    # N2 base descartado (5 archivos)
    "R13D7_131000_280.mseed", "R13D7_131500_280.mseed", "R13D7_132000_280.mseed",
    "R13D7_132500_280.mseed", "R13D7_133000_280.mseed",
    # N3 base descartado (7 archivos)
    "REAEC_132500_280.mseed", "REAEC_133000_280.mseed", "REAEC_133500_280.mseed",
    "REAEC_134000_280.mseed", "REAEC_134500_280.mseed", "REAEC_135000_280.mseed",
    "REAEC_135500_280.mseed",
    # N9 base descartado (3 archivos)
    "REAEC_144500_280.mseed", "REAEC_145000_280.mseed", "REAEC_145500_280.mseed"
}

# ── 1. LIMPIEZA DE Bitacora.xlsx ───────────────────────────────────────────
print("--- Depurando Bitacora.xlsx ---")
wb1 = openpyxl.load_workbook("Bitacora.xlsx")

# Eliminar hojas obsoletas si existen
for old_sheet in ["Analisis Individual old", "Analisis Grupal"]:
    if old_sheet in wb1.sheetnames:
        del wb1[old_sheet]
        print(f"  Hoja eliminada: {old_sheet}")

ws1 = wb1["Bitacora"]

# Actualizar N2, N3, N9 con datos activos y coordenadas
# N2: R13D7 actualizado (22:06-22:40 UTC / 15:08-15:42 equipo)
n2_files = "R13D7_220500_280.mseed; R13D7_221000_280.mseed; R13D7_221500_280.mseed; R13D7_222000_280.mseed; R13D7_222500_280.mseed; R13D7_223000_280.mseed; R13D7_223500_280.mseed"
ws1.cell(3, 2).value = COORDENADAS[2][0]
ws1.cell(3, 3).value = COORDENADAS[2][1]
ws1.cell(3, 4).value = datetime.datetime(2026, 10, 7, 22, 6, 0)
ws1.cell(3, 5).value = datetime.datetime(2026, 10, 7, 22, 40, 0)
ws1.cell(3, 6).value = datetime.datetime(2026, 10, 7, 15, 8, 0)
ws1.cell(3, 7).value = datetime.datetime(2026, 10, 7, 15, 42, 0)
ws1.cell(3, 8).value = n2_files
ws1.cell(3, 9).value = None
ws1.cell(3, 10).value = None

# N3: R13D7 actualizado (21:29-22:03 UTC / 14:31-15:05 equipo)
n3_files = "R13D7_212500_280.mseed; R13D7_213000_280.mseed; R13D7_213500_280.mseed; R13D7_214000_280.mseed; R13D7_214500_280.mseed; R13D7_215000_280.mseed; R13D7_215500_280.mseed; R13D7_220000_280.mseed"
ws1.cell(4, 2).value = COORDENADAS[3][0]
ws1.cell(4, 3).value = COORDENADAS[3][1]
ws1.cell(4, 4).value = datetime.datetime(2026, 10, 7, 21, 29, 0)
ws1.cell(4, 5).value = datetime.datetime(2026, 10, 7, 22, 3, 0)
ws1.cell(4, 6).value = datetime.datetime(2026, 10, 7, 14, 31, 0)
ws1.cell(4, 7).value = datetime.datetime(2026, 10, 7, 15, 5, 0)
ws1.cell(4, 8).value = n3_files
ws1.cell(4, 9).value = None
ws1.cell(4, 10).value = None

# N9: R13D7 actualizado (20:51-21:25 UTC / 13:53-14:27 equipo)
n9_files = "R13D7_205000_280.mseed; R13D7_205500_280.mseed; R13D7_210000_280.mseed; R13D7_210500_280.mseed; R13D7_211000_280.mseed; R13D7_211500_280.mseed; R13D7_212000_280.mseed"
ws1.cell(10, 2).value = COORDENADAS[9][0]
ws1.cell(10, 3).value = COORDENADAS[9][1]
ws1.cell(10, 4).value = datetime.datetime(2026, 10, 7, 20, 51, 0)
ws1.cell(10, 5).value = datetime.datetime(2026, 10, 7, 21, 25, 0)
ws1.cell(10, 6).value = datetime.datetime(2026, 10, 7, 13, 53, 0)
ws1.cell(10, 7).value = datetime.datetime(2026, 10, 7, 14, 27, 0)
ws1.cell(10, 8).value = n9_files
ws1.cell(10, 9).value = None
ws1.cell(10, 10).value = None

# Asegurar que todas las coordenadas de N1..N15 sean las oficiales
for r in range(2, 17):
    n_num = ws1.cell(r, 1).value
    if n_num in COORDENADAS:
        ws1.cell(r, 2).value = COORDENADAS[n_num][0]
        ws1.cell(r, 3).value = COORDENADAS[n_num][1]

wb1.save("Bitacora.xlsx")
print("  Bitacora.xlsx guardada exitosamente.")

# ── 2. LIMPIEZA DE Bitacora de Registros Nakamuras.xlsx ───────────────────
print("\n--- Depurando Bitacora de Registros Nakamuras.xlsx ---")
wb2 = openpyxl.load_workbook("Bitacora de Registros Nakamuras.xlsx")
ws2 = wb2["Sheet1"]

# N2
ws2.cell(3, 2).value = COORDENADAS[2][0]
ws2.cell(3, 3).value = COORDENADAS[2][1]
ws2.cell(3, 4).value = n2_files
ws2.cell(3, 5).value = None
ws2.cell(3, 6).value = None
ws2.cell(3, 7).value = 7

# N3
ws2.cell(4, 2).value = COORDENADAS[3][0]
ws2.cell(4, 3).value = COORDENADAS[3][1]
ws2.cell(4, 4).value = n3_files
ws2.cell(4, 5).value = None
ws2.cell(4, 6).value = None
ws2.cell(4, 7).value = 8

# N9
ws2.cell(10, 2).value = COORDENADAS[9][0]
ws2.cell(10, 3).value = COORDENADAS[9][1]
ws2.cell(10, 4).value = n9_files
ws2.cell(10, 5).value = None
ws2.cell(10, 6).value = None
ws2.cell(10, 7).value = 7

# Asegurar coordenadas oficiales para todos los sitios
for r in range(2, 17):
    n_num = ws2.cell(r, 1).value
    if n_num in COORDENADAS:
        ws2.cell(r, 2).value = COORDENADAS[n_num][0]
        ws2.cell(r, 3).value = COORDENADAS[n_num][1]

# Recalcular totales en fila 17
# Conteo R13D7: N1(5), N2(7), N3(8), N4(7), N6(7), N8(7), N9(7), N10(7), N11(8), N12(7), N13(7), N14(6), N15(7) = 13 sitios con R13D7
# Conteo REAEC: N5(10), N7(9) = 2 sitios con REAEC
# Total sitios: 15
ws2.cell(17, 4).value = 13
ws2.cell(17, 5).value = 2
ws2.cell(17, 6).value = 0
ws2.cell(17, 7).value = "=SUM(G2:G16)"

wb2.save("Bitacora de Registros Nakamuras.xlsx")
print("  Bitacora de Registros Nakamuras.xlsx guardada exitosamente.")

# ── 3. LIMPIEZA DE Informe/Resultados de Nakamuras.xlsx ───────────────────
print("\n--- Depurando Informe/Resultados de Nakamuras.xlsx ---")
wb3 = openpyxl.load_workbook(r"Informe\Resultados de Nakamuras.xlsx")

# Hoja 'Analisis Individual': filtrar registros descartados
ws_ind = wb3["Analisis Individual"]
rows_keep = []
headers = [ws_ind.cell(1, c).value for c in range(1, ws_ind.max_column+1)]

deleted_count = 0
for r in range(2, ws_ind.max_row+1):
    reg = ws_ind.cell(r, 3).value
    n_val = ws_ind.cell(r, 2).value
    if reg in DISCARDED_MSEED:
        deleted_count += 1
        continue
    # Extraer valores de la fila
    row_data = [ws_ind.cell(r, c).value for c in range(1, ws_ind.max_column+1)]
    # Actualizar coordenadas según COORDENADAS oficial
    if n_val in COORDENADAS:
        row_data[3] = COORDENADAS[n_val][0]  # Lat
        row_data[4] = COORDENADAS[n_val][1]  # Long
    rows_keep.append(row_data)

print(f"  Filtrados {deleted_count} registros descartados. Quedan {len(rows_keep)} registros válidos.")

# Recrear hoja 'Analisis Individual' limpia
# Preservar estilos copiando a una nueva hoja
del wb3["Analisis Individual"]
ws_new = wb3.create_sheet("Analisis Individual", 0)

# Escribir encabezados
for col_idx, h in enumerate(headers, 1):
    cell = ws_new.cell(1, col_idx, h)
    cell.font = Font(bold=True, color="FFFFFF", size=9)
    cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Escribir filas re-enumeradas
for row_idx, rdata in enumerate(rows_keep, 2):
    rdata[0] = row_idx - 1  # Renumerar Nro 1..N
    for col_idx, val in enumerate(rdata, 1):
        cell = ws_new.cell(row_idx, col_idx, val)
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal="center", vertical="center")

# Hoja 'Analisis Grupal': verificar coordenadas
ws_grp = wb3["Analisis Grupal"]
for r in range(2, ws_grp.max_row+1):
    n_val = ws_grp.cell(r, 1).value
    if n_val in COORDENADAS:
        ws_grp.cell(r, 2).value = COORDENADAS[n_val][1]  # Long
        ws_grp.cell(r, 3).value = COORDENADAS[n_val][0]  # Lat

# Hoja 'Final': actualizar con datos limpios
if "Final" in wb3.sheetnames:
    del wb3["Final"]
ws_fin = wb3.create_sheet("Final")
# Headers
fin_headers = ["Sitio", "Longitud", "Latitud", "T0 (s)", "A0"]
for c_idx, h in enumerate(fin_headers, 1):
    c = ws_fin.cell(1, c_idx, h)
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

for r in range(2, 17):
    sitio = ws_grp.cell(r, 4).value
    lon = ws_grp.cell(r, 2).value
    lat = ws_grp.cell(r, 3).value
    t0 = ws_grp.cell(r, 5).value
    a0 = ws_grp.cell(r, 6).value
    ws_fin.cell(r, 1, sitio)
    ws_fin.cell(r, 2, lon)
    ws_fin.cell(r, 3, lat)
    ws_fin.cell(r, 4, t0)
    ws_fin.cell(r, 5, a0)

wb3.save(r"Informe\Resultados de Nakamuras.xlsx")
print("  Informe/Resultados de Nakamuras.xlsx guardada exitosamente.")
print("\n[OK] Todos los Excels depurados con éxito.")
