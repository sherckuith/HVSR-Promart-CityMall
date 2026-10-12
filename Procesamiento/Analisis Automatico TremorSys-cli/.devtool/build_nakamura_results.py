"""
============================================================================
SCRIPT MAESTRO: Análisis HVSR Individual y Grupal de registros MiniSEED
============================================================================
DESCRIPCIÓN:
  Este script implementa el procedimiento completo de análisis HVSR Nakamura
  usando TremorSys CLI para la red sísmica ARDASYSMIC en el proyecto
  Promart-CityMall. Ejecuta los siguientes pasos:

  PASO 1: Descubrimiento de TODOS los 139 registros mseed (124 principales
           más 15 descartados en subcarpetas 'Descartados/').
  PASO 2: Análisis individual de cada registro mseed con TremorSys CLI.
           Incluye fallback automático con `--units acceleration` para
           sensores de aceleración (como REAEC).
  PASO 3: Extracción exhaustiva de parámetros sismológicos de los JSON.
  PASO 4: Algoritmo de clasificación de candidatos a descarte basado en:
           - Detección de cluster principal en períodos de suelo (T0 < 2.0 s).
           - Filtrado de períodos anómalos o de ruido ambiental (T0 >= 2.0 s).
           - Descarte de picos transitorios o artefactos espectrales.
  PASO 5: Cálculo de T mean individual (mediana de registros válidos) y 
           T mean grupal (mediana de T Mediana de todo el grupo), veredicto
           y recomendación textual.
  PASO 6: Re-ejecución del análisis grupal filtrado ÚNICAMENTE con los 
           registros mseed válidos (no candidatos).
  PASO 7: Exportación a 'Resultados de Nakamuras.xlsx' con formato condicional,
           sombreado de candidatos en gris (#D9D9D9), estilos profesionales
           y orden de columnas especificado por el usuario.

ESTRUCTURA DE COLUMNAS (Hoja 'Analisis Individual'):
  1. Nro, 2. N, 3. Registro, 4. Lat, 5. Long, 6. T0, 7. A0, 8. Kg,
  9. T min, 10. T Med, 11. NC, 12. SESAME,
  13. Sitio, 14. Est, 15. Canales, 16. Fs Hz, 17. Dur s, 18. Vent, 19. Vent s,
  20. f0 Hz, 21. f0 Min, 22. f0 Max, 23. f0 Med, 24. T Medio, 25. T Max,
  26. Conf, 27. Hora, 28. Unidades, 29. XML, 30. XML OK, 31. Elev m,
  32. Estado, 33. Veredicto, 34. Recom, 35. T Indiv, 36. T Grup,
  37. Vrdcto Grp, 38. Recom Grp, 39. Conclusion T.

USO:
  python build_nakamura_results.py

AUTOR: Angel Yaguana, ARDASYSMIC
============================================================================
"""
import os
import sys
import glob
import json
import shutil
import subprocess
import math
import statistics
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
TREMORSYS_CLI = r"D:\2026\ARDASYSMIC\ARDASYSMIC_26\TremorSys\target\release\tremorsys-cli.exe"
PROFILE_TOML = r"D:\2026\ARDASYSMIC\ARDASYSMIC_26\TremorSys\examples\report_profile.toml"
BASE_NAKAMURA_DIR = r"d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\Sitios Nakamura"
OUTPUT_INDIV_DIR = r"d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\.devtool\results_individual"
OUTPUT_GROUP_DIR = r"d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\.devtool\results_group_filtered"
OUTPUT_EXCEL = r"d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\Resultados de Nakamuras.xlsx"

# Umbral T0 máximo para suelo en ruido ambiental
T0_MAX_THRESHOLD = 2.0


# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def get_sesame_class(f0, a0):
    """Clasificación SESAME textual basada en f0 y A0."""
    if f0 <= 0 or a0 <= 0:
        return "Desc"
    elif a0 < 1.0:
        return "E"
    elif a0 < 1.5:
        return "D"
    elif a0 > 2.0:
        return "A" if f0 > 1.0 else "B"
    else:
        return "C"


def natural_sort_key(s):
    """Ordena N1, N2, N2 actualizado, ..., N10, N11 correctamente."""
    import re
    parts = re.split(r'(\d+)', s)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def get_site_number(sitio_name):
    """Extrae el número de N del nombre del sitio. Ej: 'N3 actualizado' -> 3"""
    import re
    m = re.match(r'N(\d+)', sitio_name)
    return int(m.group(1)) if m else 999


def classify_outliers(t0_values, a0_values=None):
    """
    Algoritmo de detección de candidatos a descarte basado en clusters
    y física del suelo en ruido ambiental.
    
    METODOLOGÍA:
    1. En ruido ambiental en suelos sedimentarios, el período fundamental
       T0 típicamente oscila entre 0.4 s y 1.8 s.
    2. Valores T0 >= 2.0 s corresponden a picos secundarios, transitorios
       o artefactos de campanas secundarias en la curva HVSR.
    3. Se identifica el cluster bajo (T0 < 2.0 s, excluyendo valores aislados < 0.4 s).
    4. Si no existe cluster bajo confiable, TODO el grupo se marca candidato a descarte.
    5. Si existe cluster bajo, se calcula la mediana del cluster y un margen de aceptación
       (±50% de la mediana, mínimo ±0.3 s). Todo valor fuera del margen es candidato.
    6. Se marcan también registros con amplitud A0 anómala extrema (> 20) cuando existen
       otros registros del mismo sitio con A0 normal.
    """
    if not t0_values or all(v is None for v in t0_values):
        return [True] * len(t0_values)
    
    valid = [v for v in t0_values if v is not None and v > 0]
    if not valid:
        return [True] * len(t0_values)
    
    # Filtrar candidatos en rango bajo
    low_candidates = [v for v in valid if v < T0_MAX_THRESHOLD]
    
    # En casos como N2 original, un valor aislado de 0.24s cuando los demás son >=2.0s no es cluster
    low_cluster = [v for v in low_candidates if v >= 0.4 or len(low_candidates) >= 2]
    
    if not low_cluster:
        return [True] * len(t0_values)
    
    low_median = statistics.median(low_cluster)
    margin = max(low_median * 0.5, 0.3)
    lower_bound = max(0.4, low_median - margin)
    upper_bound = min(T0_MAX_THRESHOLD, low_median + margin)
    
    results = []
    for i, v in enumerate(t0_values):
        if v is None:
            results.append(True)
        elif v >= T0_MAX_THRESHOLD:
            results.append(True)
        elif v < lower_bound:
            results.append(True)
        else:
            # Caso especial: A0 extremadamente anómalo (> 20) si otro registro tiene misma T0 y A0 normal
            if (a0_values and a0_values[i] is not None and a0_values[i] > 20.0 and 
                any(a0_values[j] is not None and a0_values[j] < 10.0 and abs(t0_values[j] - v) < 0.1 for j in range(len(t0_values)) if j != i)):
                results.append(True)
            else:
                results.append(False)
    
    return results


# ============================================================================
# PASO 1: DESCUBRIMIENTO DE TODOS LOS MSEED
# ============================================================================

def discover_all_mseed():
    """
    Descubre todos los archivos mseed organizados por sitio,
    incluyendo los de carpetas Descartados/.
    
    Retorna: dict[sitio_name] -> list[{path, filename, is_discarded}]
    """
    sites = OrderedDict()
    site_dirs = sorted(
        [d for d in glob.glob(os.path.join(BASE_NAKAMURA_DIR, "N*")) if os.path.isdir(d)],
        key=lambda x: natural_sort_key(os.path.basename(x))
    )
    
    for site_dir in site_dirs:
        site_name = os.path.basename(site_dir)
        reg_dir = os.path.join(site_dir, "Registros")
        if not os.path.isdir(reg_dir):
            continue
        
        files = []
        # Archivos principales
        for f in sorted(glob.glob(os.path.join(reg_dir, "*.mseed"))):
            files.append({
                "path": f,
                "filename": os.path.basename(f),
                "is_discarded": False
            })
        
        # Archivos descartados previamente
        desc_dir = os.path.join(reg_dir, "Descartados")
        if os.path.isdir(desc_dir):
            for f in sorted(glob.glob(os.path.join(desc_dir, "*.mseed"))):
                files.append({
                    "path": f,
                    "filename": os.path.basename(f),
                    "is_discarded": True
                })
        
        if files:
            sites[site_name] = files
    
    return sites


# ============================================================================
# PASO 2: ANÁLISIS INDIVIDUAL CON TREMORSYS CLI
# ============================================================================

def run_individual_analysis(sites):
    """
    Ejecuta tremorsys-cli en modo individual sobre cada mseed.
    Si el comando falla por unidades de aceleración (ej. REAEC),
    reintenta automáticamente con `--units acceleration`.
    """
    total = sum(len(v) for v in sites.values())
    processed = 0
    
    for site_name, files in sites.items():
        for finfo in files:
            processed += 1
            prefix = "DESC_" if finfo["is_discarded"] else ""
            out_subdir = os.path.join(
                OUTPUT_INDIV_DIR,
                f"{site_name}_{prefix}{finfo['filename'].replace('.mseed', '')}"
            )
            
            # Verificar si ya existe resultado válido
            json_matches = glob.glob(os.path.join(out_subdir, "*__resultados.json"))
            if json_matches:
                print(f"  [{processed}/{total}] Ya existe: {site_name}/{finfo['filename']}")
                continue
            
            print(f"  [{processed}/{total}] Procesando: {site_name}/{finfo['filename']}")
            os.makedirs(out_subdir, exist_ok=True)
            
            cmd = [
                TREMORSYS_CLI,
                "--input", finfo["path"],
                "--output", out_subdir,
                "--profile", PROFILE_TOML,
                "--no-record", "--no-noise", "--no-watermark"
            ]
            res = subprocess.run(cmd, capture_output=True, text=True)
            
            # Fallback para archivos con acelerómetros (unidades m/s²)
            if res.returncode != 0:
                cmd_acc = [
                    TREMORSYS_CLI,
                    "--input", finfo["path"],
                    "--units", "acceleration",
                    "--output", out_subdir,
                    "--profile", PROFILE_TOML,
                    "--no-record", "--no-noise", "--no-watermark"
                ]
                subprocess.run(cmd_acc, capture_output=True, text=True)


# ============================================================================
# PASO 3: EXTRACCIÓN DE RESULTADOS INDIVIDUALES
# ============================================================================

def extract_individual_results(sites):
    """
    Lee los JSONs de resultados individuales y devuelve una lista
    de diccionarios con todos los campos necesarios.
    """
    all_rows = []
    
    for site_name in sites.keys():
        files = sites[site_name]
        for finfo in files:
            prefix = "DESC_" if finfo["is_discarded"] else ""
            out_subdir = os.path.join(
                OUTPUT_INDIV_DIR,
                f"{site_name}_{prefix}{finfo['filename'].replace('.mseed', '')}"
            )
            
            json_matches = glob.glob(os.path.join(out_subdir, "*__resultados.json"))
            if not json_matches:
                # Archivo sin resultado ejecutable
                all_rows.append({
                    "n": get_site_number(site_name),
                    "sitio": site_name,
                    "registro": finfo["filename"],
                    "path": finfo["path"],
                    "is_discarded": finfo["is_discarded"],
                    "error": True
                })
                continue
            
            with open(json_matches[0], "r", encoding="utf-8") as f:
                data = json.load(f)
            
            traces = data.get("traces") or []
            tr0 = traces[0] if traces else {}
            coords = data.get("coordinates") or {}
            stats = data.get("statistics") or {}
            
            f0 = data.get("f0_hz", 0.0)
            a0 = data.get("a0", 0.0)
            t0_s = data.get("t0_s", 0.0)
            
            lat = tr0.get("latitude") if tr0.get("latitude") is not None else coords.get("latitude")
            lon = tr0.get("longitude") if tr0.get("longitude") is not None else coords.get("longitude")
            elev = tr0.get("elevation_m") if tr0.get("elevation_m") is not None else coords.get("elevation_m")
            
            row = {
                "n": get_site_number(site_name),
                "sitio": site_name,
                "registro": finfo["filename"],
                "path": finfo["path"],
                "is_discarded": finfo["is_discarded"],
                "error": False,
                "lat": lat,
                "long": lon,
                "t0": t0_s,
                "a0": a0,
                "kg": data.get("kg", 0.0),
                "t_min": stats.get("period_low_s"),
                "t_med": stats.get("period_median_s"),
                "nc": data.get("num_cycles", 0.0),
                "sesame": get_sesame_class(f0, a0),
                # Columnas adicionales
                "estacion": tr0.get("station", ""),
                "canales": len(traces),
                "fs_hz": tr0.get("sample_rate_hz", 0.0),
                "dur_s": tr0.get("duration_s", 0.0),
                "ventanas": data.get("num_windows", 0),
                "ventana_s": data.get("window_length_s", 0.0),
                "f0_hz": f0,
                "f0_low": stats.get("f0_low_hz"),
                "f0_high": stats.get("f0_high_hz"),
                "f0_med": stats.get("f0_median_hz"),
                "t_mean_hz": stats.get("t_mean_hz"),
                "t_max": stats.get("period_high_s"),
                "confiable": data.get("reliable", False),
                "hora": tr0.get("start_time", ""),
                "unidades": data.get("units_status", ""),
                "xml": os.path.basename(data.get("stationxml", "") or ""),
                "xml_ok": data.get("stationxml_complete", False),
                "elev": elev,
            }
            all_rows.append(row)
    
    return all_rows


# ============================================================================
# PASO 4-5: CLASIFICACIÓN, ORDENAMIENTO Y VEREDICTOS
# ============================================================================

def apply_classification_and_verdicts(rows):
    """
    Aplica el algoritmo de clasificación de candidatos a descarte.
    Reordena las filas dentro de cada grupo para que los registros válidos
    aparezcan primero y los candidatos/descartados aparezcan al final
    (sombreados en gris), replicando exactamente la estructura de Bitacora.xlsx.
    
    Calcula T mean indiv y T mean grupal y emite veredictos por grupo.
    Retorna la lista final ordenada con 'nro' consecutivo (1..139).
    """
    # Agrupar por sitio conservando orden natural
    groups = OrderedDict()
    for row in rows:
        site = row["sitio"]
        if site not in groups:
            groups[site] = []
        groups[site].append(row)
    
    ordered_rows = []
    current_nro = 1
    
    for site_name, group_rows in groups.items():
        # Registros candidatos previos y principales
        non_disc_rows = [r for r in group_rows if not r["is_discarded"] and not r.get("error")]
        t0_values = [r.get("t0") for r in non_disc_rows]
        a0_values = [r.get("a0") for r in non_disc_rows]
        
        # Ejecutar clasificador de clusters
        if t0_values:
            is_outlier = classify_outliers(t0_values, a0_values)
        else:
            is_outlier = []
        
        # Asignar flag 'candidato'
        non_disc_idx = 0
        for row in group_rows:
            if row["is_discarded"] or row.get("error"):
                row["candidato"] = True
            else:
                row["candidato"] = is_outlier[non_disc_idx] if non_disc_idx < len(is_outlier) else True
                non_disc_idx += 1
        
        # Separar en válidos y candidatos para ordenarlos dentro del grupo
        valid_part = [r for r in group_rows if not r["candidato"] and not r["is_discarded"]]
        candidate_part = [r for r in group_rows if r["candidato"] or r["is_discarded"]]
        
        # Calcular T mean individual (mediana de válidos)
        valid_t0 = [r["t0"] for r in valid_part if r.get("t0") is not None and r.get("t0") > 0]
        t_mean_indiv = statistics.median(valid_t0) if valid_t0 else None
        
        # Calcular T mean grupal (mediana de T Mediana de todo el grupo)
        all_t_med = [r.get("t_med") for r in group_rows if r.get("t_med") is not None and r.get("t_med") > 0]
        t_mean_grupal = statistics.median(all_t_med) if all_t_med else None
        
        # Determinar veredicto de grupo
        if not valid_t0:
            veredicto_grupo = "REVISAR"
            recom_grupo = "GRUPO COMPLETO"
        else:
            if t_mean_indiv is not None and t_mean_grupal is not None:
                diff = abs(round(t_mean_indiv, 2) - round(t_mean_grupal, 2))
                if diff <= 0.15:
                    veredicto_grupo = "OK"
                    recom_grupo = None
                else:
                    veredicto_grupo = "REVISAR"
                    recom_grupo = None
            else:
                veredicto_grupo = "REVISAR"
                recom_grupo = "GRUPO COMPLETO"
        
        # Conclusión de T0 para el grupo
        if veredicto_grupo == "OK" and t_mean_indiv is not None:
            t0_conclusion = min(filter(None, [t_mean_indiv, t_mean_grupal]))
        elif t_mean_indiv is not None:
            t0_conclusion = t_mean_indiv
        else:
            t0_conclusion = t_mean_grupal
        
        # Combinar registros del grupo (válidos primero, candidatos después)
        sorted_group = valid_part + candidate_part
        
        for idx, row in enumerate(sorted_group):
            row["nro"] = current_nro
            current_nro += 1
            
            # Estado del registro
            if row["is_discarded"]:
                row["estado"] = "Descartado previo"
            elif row.get("candidato"):
                row["estado"] = "Candidato a descarte"
            else:
                row["estado"] = "Válido"
            
            # Veredicto y recomendación por fila
            if row.get("candidato") or row["is_discarded"] or row.get("error"):
                row["veredicto"] = "REVISAR"
                row["recomendacion"] = "GRUPO COMPLETO" if recom_grupo == "GRUPO COMPLETO" else "MSEED INDIVIDUAL"
            else:
                row["veredicto"] = "OK"
                row["recomendacion"] = None
            
            # Resumen del grupo en la primera fila
            if idx == 0:
                row["t_mean_indiv"] = t_mean_indiv
                row["t_mean_grupal"] = t_mean_grupal
                row["veredicto_grupo"] = veredicto_grupo
                row["recom_grupo"] = recom_grupo
                row["t0_conclusion"] = t0_conclusion
            else:
                row["t_mean_indiv"] = None
                row["t_mean_grupal"] = None
                row["veredicto_grupo"] = None
                row["recom_grupo"] = None
                row["t0_conclusion"] = None
            
            ordered_rows.append(row)
    
    return ordered_rows


# ============================================================================
# PASO 6: ANÁLISIS GRUPAL FILTRADO
# ============================================================================

def run_filtered_group_analysis(rows):
    """
    Para cada grupo N, copia ÚNICAMENTE los mseed válidos (no candidatos)
    a un directorio temporal, ejecuta tremorsys-cli en modo grupal,
    y devuelve los resultados para la hoja Analisis Grupal.
    """
    groups = OrderedDict()
    for row in rows:
        site = row["sitio"]
        if site not in groups:
            groups[site] = []
        groups[site].append(row)
    
    group_results = []
    
    for site_name, group_rows in groups.items():
        # Filtrar solo mseed válidos (no candidatos y no descartados previos)
        valid_rows = [r for r in group_rows 
                      if not r.get("candidato") and not r.get("error") and not r["is_discarded"]]
        
        # Conclusión T0 calculada previamente
        first_r = group_rows[0]
        t0_concl = first_r.get("t0_conclusion")
        verd_grp = first_r.get("veredicto_grupo")
        
        if len(valid_rows) < 2:
            print(f"  {site_name}: {len(valid_rows)} válido(s), análisis grupal no aplicable (< 2 mseed)")
            group_results.append({
                "sitio": site_name,
                "n": get_site_number(site_name),
                "mseed_count": len(valid_rows),
                "error": True,
                "veredicto": verd_grp,
                "t0_conclusion": t0_concl,
                "motivo": f"Sin datos válidos ({len(group_rows)} mseed candidatos a descarte)"
            })
            continue
        
        # Agrupar en bloques contiguos en el tiempo (300 s nominales) para evitar huecos en el CLI
        from datetime import datetime
        def parse_iso(t_str):
            if not t_str: return None
            try:
                return datetime.fromisoformat(t_str.split('.')[0].replace('Z', ''))
            except:
                return None

        sorted_valid = sorted(
            valid_rows, 
            key=lambda r: parse_iso(r.get("hora")) or datetime.min
        )
        
        blocks = []
        current_block = []
        for r in sorted_valid:
            t = parse_iso(r.get("hora"))
            if not current_block:
                current_block.append(r)
            else:
                prev_t = parse_iso(current_block[-1].get("hora"))
                if t and prev_t and 240 <= (t - prev_t).total_seconds() <= 360:
                    current_block.append(r)
                else:
                    blocks.append(current_block)
                    current_block = [r]
        if current_block:
            blocks.append(current_block)
            
        target_block = max(blocks, key=len) if blocks else []
        if len(target_block) < 2:
            print(f"  {site_name}: {len(valid_rows)} válidos pero sin bloque continuo >= 2 mseed")
            group_results.append({
                "sitio": site_name,
                "n": get_site_number(site_name),
                "mseed_count": len(valid_rows),
                "error": True,
                "veredicto": verd_grp,
                "t0_conclusion": t0_concl,
                "motivo": f"Sin bloque continuo >= 2 mseed ({len(valid_rows)} válidos no contiguos)"
            })
            continue

        # Crear directorio temporal con mseed del bloque continuo
        temp_dir = os.path.join(OUTPUT_GROUP_DIR, f"temp_{site_name}")
        shutil.rmtree(temp_dir, ignore_errors=True)
        os.makedirs(temp_dir, exist_ok=True)
        
        for vr in target_block:
            src = vr.get("path")
            if src and os.path.exists(src):
                shutil.copy2(src, temp_dir)
        
        # Copiar StationXML
        reg_dir = os.path.join(BASE_NAKAMURA_DIR, site_name, "Registros")
        for xml in glob.glob(os.path.join(reg_dir, "*.xml")):
            shutil.copy2(xml, temp_dir)
        
        out_dir = os.path.join(OUTPUT_GROUP_DIR, site_name)
        shutil.rmtree(out_dir, ignore_errors=True)
        os.makedirs(out_dir, exist_ok=True)
        
        group_size = len(target_block)
        print(f"  {site_name}: Grupal filtrado con {group_size} mseed continuos válidos")
        
        cmd = [
            TREMORSYS_CLI,
            "--group-input", temp_dir,
            "--group-size", str(group_size),
            "--output", out_dir,
            "--profile", PROFILE_TOML,
            "--no-record", "--no-noise", "--no-watermark"
        ]
        subprocess.run(cmd, capture_output=True, text=True)
        
        # Leer resultados del grupo
        json_matches = glob.glob(os.path.join(out_dir, "*__resultados.json"))
        if json_matches:
            with open(json_matches[0], "r", encoding="utf-8") as f:
                data = json.load(f)
            
            traces = data.get("traces") or []
            tr0 = traces[0] if traces else {}
            coords = data.get("coordinates") or {}
            stats = data.get("statistics") or {}
            
            f0 = data.get("f0_hz", 0.0)
            a0 = data.get("a0", 0.0)
            
            lat = tr0.get("latitude") if tr0.get("latitude") is not None else coords.get("latitude")
            lon = tr0.get("longitude") if tr0.get("longitude") is not None else coords.get("longitude")
            
            group_results.append({
                "sitio": site_name,
                "n": get_site_number(site_name),
                "mseed_count": group_size,
                "error": False,
                "registro": f"GRUPO_{site_name} ({group_size} mseed)",
                "estacion": tr0.get("station", ""),
                "dur_s": tr0.get("duration_s", 0.0),
                "ventanas": data.get("num_windows", 0),
                "f0_hz": f0,
                "t0": data.get("t0_s", 0.0),
                "a0": a0,
                "kg": data.get("kg", 0.0),
                "t_min": stats.get("period_low_s"),
                "t_med": stats.get("period_median_s"),
                "t_max": stats.get("period_high_s"),
                "nc": data.get("num_cycles", 0.0),
                "sesame": get_sesame_class(f0, a0),
                "confiable": data.get("reliable", False),
                "lat": lat,
                "long": lon,
                "t_mean_hz": stats.get("t_mean_hz"),
                "veredicto": verd_grp,
                "t0_conclusion": t0_concl,
            })
        else:
            group_results.append({
                "sitio": site_name,
                "n": get_site_number(site_name),
                "mseed_count": group_size,
                "error": True,
                "veredicto": verd_grp,
                "t0_conclusion": t0_concl,
                "motivo": "Error al procesar grupo"
            })
        
        # Limpiar carpeta temporal
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    return group_results


# ============================================================================
# PASO 7: EXPORTACIÓN A EXCEL ('Resultados de Nakamuras.xlsx')
# ============================================================================

def write_excel(indiv_rows, group_rows):
    """
    Escribe el archivo Excel 'Resultados de Nakamuras.xlsx' con dos hojas:
      1. 'Analisis Individual': Todos los 139 mseed con el orden exacto de columnas
         pedido por el usuario y sombreado en gris de candidatos a descarte.
      2. 'Analisis Grupal': Resultados consolidados del análisis grupal filtrado.
    """
    wb = openpyxl.Workbook()
    
    # ---- PALETA DE ESTILOS PROFESIONAL ----
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=9)
    bold_font = Font(name="Calibri", size=9, bold=True, color="002060")
    ok_font = Font(name="Calibri", size=9, bold=True, color="006600")
    rev_font = Font(name="Calibri", size=9, bold=True, color="CC0000")
    
    gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    right = Alignment(horizontal='right', vertical='center')
    
    # ================================================================
    # HOJA 1: Análisis Individual
    # ================================================================
    ws1 = wb.active
    ws1.title = "Analisis Individual"
    
    # 39 columnas en el orden estricto solicitado:
    # 1 a 12 según especificación del usuario, luego todas las columnas de bitácora,
    # y al final las columnas de validación y resumen.
    headers_indiv = [
        ("nro",              "Nro",          center, "0"),
        ("n",                "N",            center, "0"),
        ("registro",         "Registro",     left,   None),
        ("lat",              "Lat",          right,  "0.0000000"),
        ("long",             "Long",         right,  "0.0000000"),
        ("t0",               "T0",           right,  "0.0000"),
        ("a0",               "A0",           right,  "0.000"),
        ("kg",               "Kg",           right,  "0.00"),
        ("t_min",            "T min",        right,  "0.0000"),
        ("t_med",            "T Med",        right,  "0.0000"),
        ("nc",               "NC",           right,  "0.0"),
        ("sesame",           "SESAME",       center, None),
        # Columnas de bitácora
        ("sitio",            "Sitio",        left,   None),
        ("estacion",         "Est",          center, None),
        ("canales",          "Canales",      center, "0"),
        ("fs_hz",            "Fs Hz",        right,  "0.0"),
        ("dur_s",            "Dur s",        right,  "0.0"),
        ("ventanas",         "Vent",         center, "0"),
        ("ventana_s",        "Vent s",       right,  "0.0"),
        ("f0_hz",            "f0 Hz",        right,  "0.0000"),
        ("f0_low",           "f0 Min",       right,  "0.0000"),
        ("f0_high",          "f0 Max",       right,  "0.0000"),
        ("f0_med",           "f0 Med",       right,  "0.0000"),
        ("t_mean_hz",        "T Medio",      right,  "0.0000"),
        ("t_max",            "T Max",        right,  "0.0000"),
        ("confiable",        "Conf",         center, None),
        ("hora",             "Hora",         left,   None),
        ("unidades",         "Unidades",     center, None),
        ("xml",              "XML",          left,   None),
        ("xml_ok",           "XML OK",       center, None),
        ("elev",             "Elev m",       right,  "0.0"),
        # Columnas de validación y veredicto
        ("estado",           "Estado",       center, None),
        ("veredicto",        "Veredicto",    center, None),
        ("recomendacion",    "Recom",        center, None),
        ("t_mean_indiv",     "T Indiv",      right,  "0.0000"),
        ("t_mean_grupal",    "T Grup",       right,  "0.0000"),
        ("veredicto_grupo",  "Vrdcto Grp",   center, None),
        ("recom_grupo",      "Recom Grp",    center, None),
        ("t0_conclusion",    "Conclusion T", right,  "0.0000"),
    ]
    
    # Escribir encabezados
    for col_idx, (key, title, align, fmt) in enumerate(headers_indiv, 1):
        cell = ws1.cell(row=1, column=col_idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    
    # Escribir filas de datos
    for row_idx, rdata in enumerate(indiv_rows, 2):
        is_candidate = rdata.get("candidato", False)
        is_discarded = rdata.get("is_discarded", False)
        
        for col_idx, (key, title, align, fmt) in enumerate(headers_indiv, 1):
            val = rdata.get(key)
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.alignment = align
            cell.border = thin_border
            
            if fmt and isinstance(val, (int, float)):
                cell.number_format = fmt
            
            # Sombreado de candidatos / descartados en GRIS
            if is_candidate or is_discarded:
                cell.fill = gray_fill
            
            # Veredicto individual coloreado
            if key == "veredicto":
                if val == "OK":
                    cell.font = ok_font
                    cell.fill = green_fill
                elif val == "REVISAR":
                    cell.font = rev_font
                    cell.fill = red_fill if not is_candidate else gray_fill
            
            # T0 destacado en negrita
            if key == "t0" and val:
                cell.font = bold_font
            
            # Resumen de grupo destacado
            if key in ("t_mean_indiv", "t_mean_grupal", "t0_conclusion") and val is not None:
                cell.font = bold_font
            
            # Veredicto grupo destacado
            if key == "veredicto_grupo" and val:
                cell.font = ok_font if val == "OK" else rev_font
    
    # Auto-ajustar anchos
    for col_idx, (key, title, _, _) in enumerate(headers_indiv, 1):
        col_letter = get_column_letter(col_idx)
        max_len = len(title)
        for rdata in indiv_rows[:25]:
            v_str = str(rdata.get(key, "") or "")
            max_len = max(max_len, len(v_str))
        ws1.column_dimensions[col_letter].width = min(max(max_len + 3, 7), 24)
    
    ws1.freeze_panes = "A2"
    
    # ================================================================
    # HOJA 2: Análisis Grupal
    # ================================================================
    ws2 = wb.create_sheet("Analisis Grupal")
    
    headers_group = [
        ("n",              "N",            center, "0"),
        ("sitio",          "Sitio",        left,   None),
        ("registro",       "Grupo",        left,   None),
        ("estacion",       "Est",          center, None),
        ("dur_s",          "Dur s",        right,  "0.0"),
        ("ventanas",       "Vent",         center, "0"),
        ("f0_hz",          "f0 Hz",        right,  "0.0000"),
        ("t0",             "T0",           right,  "0.0000"),
        ("a0",             "A0",           right,  "0.000"),
        ("kg",             "Kg",           right,  "0.00"),
        ("t_min",          "T min",        right,  "0.0000"),
        ("t_med",          "T Med",        right,  "0.0000"),
        ("t_max",          "T Max",        right,  "0.0000"),
        ("nc",             "NC",           right,  "0.0"),
        ("sesame",         "SESAME",       center, None),
        ("confiable",      "Conf",         center, None),
        ("lat",            "Lat",          right,  "0.0000000"),
        ("long",           "Long",         right,  "0.0000000"),
        ("t_mean_hz",      "T Hz",         right,  "0.0000"),
        ("veredicto",      "Veredicto",    center, None),
        ("t0_conclusion",  "T0 Concl",     right,  "0.0000"),
    ]
    
    for col_idx, (key, title, align, fmt) in enumerate(headers_group, 1):
        cell = ws2.cell(row=1, column=col_idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    
    for row_idx, rdata in enumerate(group_rows, 2):
        if rdata.get("error"):
            # Fila de grupo no aplicable / sin registros válidos
            cell_n = ws2.cell(row=row_idx, column=1, value=rdata.get("n"))
            cell_n.font = data_font
            cell_n.border = thin_border
            cell_n.alignment = center
            
            cell_site = ws2.cell(row=row_idx, column=2, value=rdata.get("sitio"))
            cell_site.font = data_font
            cell_site.border = thin_border
            
            cell_msg = ws2.cell(row=row_idx, column=3, value=rdata.get("motivo", "Sin datos válidos"))
            cell_msg.font = rev_font
            cell_msg.fill = red_fill
            cell_msg.border = thin_border
            
            # Llenar celdas restantes con borde
            for c in range(4, len(headers_group) - 1):
                c_cell = ws2.cell(row=row_idx, column=c, value="-")
                c_cell.alignment = center
                c_cell.font = data_font
                c_cell.border = thin_border
            
            # Veredicto y T0 concluido
            cell_v = ws2.cell(row=row_idx, column=len(headers_group) - 1, value=rdata.get("veredicto"))
            cell_v.font = rev_font
            cell_v.fill = red_fill
            cell_v.alignment = center
            cell_v.border = thin_border
            
            cell_c = ws2.cell(row=row_idx, column=len(headers_group), value=rdata.get("t0_conclusion"))
            cell_c.font = bold_font
            cell_c.alignment = right
            cell_c.border = thin_border
            if isinstance(rdata.get("t0_conclusion"), (int, float)):
                cell_c.number_format = "0.0000"
            continue
        
        for col_idx, (key, title, align, fmt) in enumerate(headers_group, 1):
            val = rdata.get(key)
            cell = ws2.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.alignment = align
            cell.border = thin_border
            if fmt and isinstance(val, (int, float)):
                cell.number_format = fmt
            if key in ("t0", "t0_conclusion"):
                cell.font = bold_font
            if key == "veredicto":
                if val == "OK":
                    cell.font = ok_font
                    cell.fill = green_fill
                else:
                    cell.font = rev_font
                    cell.fill = red_fill
    
    # Auto-ajustar anchos hoja grupal
    for col_idx, (key, title, _, _) in enumerate(headers_group, 1):
        col_letter = get_column_letter(col_idx)
        ws2.column_dimensions[col_letter].width = max(len(title) + 4, 9)
    
    ws2.freeze_panes = "A2"
    
    wb.save(OUTPUT_EXCEL)
    print(f"\nExcel guardado exitosamente: {OUTPUT_EXCEL}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 70)
    print("ANÁLISIS HVSR NAKAMURA - PROMART CITYMALL (ARDASYSMIC)")
    print("=" * 70)
    
    # Paso 1: Descubrir todos los mseed
    print("\n[PASO 1] Descubriendo archivos mseed...")
    sites = discover_all_mseed()
    total = sum(len(v) for v in sites.values())
    print(f"  Encontrados: {len(sites)} sitios, {total} archivos mseed totales")
    
    # Paso 2: Análisis individual
    print("\n[PASO 2] Ejecutando análisis individual con TremorSys CLI...")
    run_individual_analysis(sites)
    
    # Paso 3: Extraer resultados individuales
    print("\n[PASO 3] Extrayendo resultados individuales de los JSON...")
    indiv_rows_raw = extract_individual_results(sites)
    print(f"  Extraídos: {len(indiv_rows_raw)} registros")
    
    # Paso 4-5: Clasificar, ordenar y emitir veredictos
    print("\n[PASO 4-5] Aplicando clasificación de candidatos, ordenamiento y veredictos...")
    indiv_rows = apply_classification_and_verdicts(indiv_rows_raw)
    
    n_ok = sum(1 for r in indiv_rows if r.get("veredicto") == "OK")
    n_rev = sum(1 for r in indiv_rows if r.get("veredicto") == "REVISAR")
    print(f"  Registros clasificados: OK = {n_ok}, REVISAR = {n_rev}")
    
    # Paso 6: Análisis grupal filtrado
    print("\n[PASO 6] Ejecutando análisis grupal filtrado (solo registros válidos)...")
    group_rows = run_filtered_group_analysis(indiv_rows)
    
    # Paso 7: Exportación a Excel
    print("\n[PASO 7] Exportando resultados a 'Resultados de Nakamuras.xlsx'...")
    write_excel(indiv_rows, group_rows)
    
    print("\n" + "=" * 70)
    print("PROCESO COMPLETADO SATISFACTORIAMENTE")
    print("=" * 70)


if __name__ == "__main__":
    main()
