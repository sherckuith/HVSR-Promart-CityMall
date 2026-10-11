# Manual de Automatización y Metodología Nakamura (HVSR) para Agentes Autónomos de IA
## Especificación Algorítmica, Contratos de Datos y Pipeline de Ejecución
### Red Sísmica ARDASYSMIC · Antigravity IDE · 2026

---

## 1. Alcance y Rol del Agente

Este manual define el protocolo de actuación autónomo para modelos de lenguaje y agentes de programación (Gemini, Claude, Codex, Antigravity) asignados a tareas de procesamiento, análisis de señales MiniSEED, control de calidad espectral, consolidación en bases de datos Excel y generación de informes sismológicos bajo la metodología Nakamura (HVSR / SESAME 2004).

---

## 2. Reglas Mandatorias de Gobernanza y Nomenclatura

El agente DEBE cumplir de forma irrestricta las siguientes reglas operativas:
1. **Atribución de Datos:** Todo dato, registro, traza o metadato pertenece a la red de monitoreo **ARDASYSMIC**.
2. **Prohibición de Hardware de Terceros:** NUNCA referirse a marcas externas de hardware doméstico o comercial (por ejemplo, "Raspberry" o "Raspberry Shake").
3. **Nomenclatura Estricta de Estaciones:** Las estaciones se denominan exclusivamente por su identificador de red:
   - `R13D7`: Estación sísmica con sensor velocimétrico triaxial (geófono broadband, componentes `EHN`, `EHE`, `EHZ`).
   - `REAEC`: Estación sísmica con sensor acelerométrico triaxial digital MEMS (componentes `ENN`, `ENE`, `ENZ`).
4. **Prohibición de Colores:** NUNCA utilizar denominaciones cromáticas (tales como "Blue", "Yellow", "azul", "amarillo") en variables, columnas, informes ni documentación.
5. **Gobernanza de Repositorio:** El archivo con la exportación del chat de mensajería (`+593 96 953 1673.zip`) se excluye explícitamente en `.gitignore`. Todos los demás archivos (RAW MiniSEED, fotos, scripts, modelos, manuales) se versionan en Git.

---

## 3. Pipeline Algorítmico y Pseudo-Código Maestro

El procesamiento completo se estructura en 5 etapas secuenciales:

```
[Descubrimiento MiniSEED] -> [Fase 1: CLI Individual] -> [Clustering y Outliers] -> [Fase 2: CLI Grupal Continuo] -> [Excel + DOCX]
```

### 3.1 Pseudo-Código del Pipeline Maestro
```python
def master_hvsr_pipeline(workspace_dir):
    # 1. Descubrimiento de archivos MiniSEED
    mseed_files = discover_all_mseed(workspace_dir / "Sitios Nakamura")
    
    # 2. Análisis Individual
    individual_results = []
    for f in mseed_files:
        is_accel = detect_accelerometer_metadata(f) # canal ENN/ENE/ENZ o estacion REAEC
        cmd = build_cli_command(f, is_acceleration=is_accel)
        run_cli_sync(cmd)
        json_data = parse_json_result(f)
        individual_results.append(json_data)
        
    # 3. Clustering y detección de candidatos a descarte
    classified_results = apply_clustering_algorithm(individual_results)
    
    # 4. Particionamiento en bloques continuos para análisis grupal
    for site, site_records in group_by_site(classified_results):
        valid_records = [r for r in site_records if r.status == "VALIDO"]
        if len(valid_records) >= 2:
            continuous_blocks = find_contiguous_blocks(valid_records, max_gap_seconds=330)
            largest_block = max(continuous_blocks, key=len)
            run_group_cli(largest_block)
            
    # 5. Generación de Base de Datos Excel y Reporte Word
    write_clean_excel(classified_results, output_path="Resultados de Nakamuras.xlsx")
    generate_word_report(classified_results, output_path="INFORME HVSR Promart CityMall (Geosísmica).docx")
```

---

## 4. Algoritmo de Clustering y Filtrado de Outliers

### 4.1 Justificación Física
En depósitos sedimentarios cuaternarios de la llanura aluvial y estuarina de Guayaquil, el período fundamental del suelo raramente supera los 2.0 segundos ($f_0 \ge 0.5\text{ Hz}$). Los picos espectrales con $T_0 \ge 2.0\text{ s}$ típicamente se deben a:
- Picos antropogénicos locales de muy baja frecuencia (tráfico pesado distante, camiones de volteo).
- Deriva instrumental térmica o de línea base en acelerómetros MEMS.
- Modos superiores espurios cuando la campana H/V es ancha o achatada.

### 4.2 Algoritmo Determinístico de Clasificación
```python
def classify_outliers(t0_values):
    # t0_values: lista de floats correspondientes al período T0 de cada mseed del sitio
    cluster_low = [t for t in t0_values if t is not None and t < 2.0]
    
    # Caso patológico: sin cluster bajo
    if not cluster_low or (len(cluster_low) == 1 and cluster_low[0] < 0.4):
        return ["CANDIDATO_DESCARTE" for _ in t0_values]
        
    median_low = statistics.median(cluster_low)
    margin = max(0.5 * median_low, 0.3)
    t_min = max(0.4, median_low - margin)
    t_max = min(2.0, median_low + margin)
    
    status = []
    for t in t0_values:
        if t is None or t >= 2.0 or t < t_min or t > t_max:
            status.append("CANDIDATO_DESCARTE")
        else:
            status.append("VALIDO")
    return status
```

---

## 5. Algoritmo de Particionamiento de Bloques Continuos

`tremorsys-cli` requiere que las trazas agrupadas no presenten saltos temporales superiores a 300 segundos entre archivos consecutivos. Cuando se descartan registros intermedios, el agente DEBE ejecutar:

```python
def find_contiguous_blocks(record_list, max_gap_seconds=330):
    # record_list ordenado cronológicamente por start_utc
    blocks = []
    current_block = [record_list[0]]
    
    for i in range(1, len(record_list)):
        prev_end = current_block[-1].end_utc
        curr_start = record_list[i].start_utc
        gap = (curr_start - prev_end).total_seconds()
        
        # Si el salto temporal está entre 0 y 330 segundos, es continuo
        if -10 <= gap <= max_gap_seconds:
            current_block.append(record_list[i])
        else:
            blocks.append(current_block)
            current_block = [record_list[i]]
            
    if current_block:
        blocks.append(current_block)
        
    return blocks
```

El CLI grupal se lanza únicamente sobre el bloque de mayor longitud:
```bash
tremorsys-cli --group-input <temp_dir_con_bloque> --group-size <len(bloque)> ...
```

---

## 6. Esquema de Base de Datos Excel (`Resultados de Nakamuras.xlsx`)

### Hoja 1: `Analisis Individual` (41 Columnas)
1. `Nro`: Entero incremental consecutivo (1..N).
2. `N`: Número de sitio (1..15).
3. `Registro`: Nombre del archivo MiniSEED.
4. `Lat`: Latitud oficial en grados decimales (6-7 decimales).
5. `Long`: Longitud oficial en grados decimales (6-7 decimales).
6. `T0`: Período fundamental extraído de `__resultados.json` (s).
7. `A0`: Amplitud espectral en $T_0$.
8. `Kg`: Índice de vulnerabilidad de Nakamura ($A_0^2 \cdot T_0$).
9. `T min`: Período inferior de ventana de confianza.
10. `T Med`: Período mediano de ventana.
11. `NC`: Ciclos SESAME acumulados ($N_c$).
12. `SESAME`: Clasificación textual (A, B, C, D o E).
13..32: Columnas de soporte (Estación, Canales, Fs, Duración, Ventanas, f0, Coordenadas, etc.).
33. `Veredicto`: `OK` (verde `#C6EFCE`) o `REVISAR` (rojo suave `#FFC7CE`).
34. `Recom`: `Válido` o `Candidato a descarte`.
35. `T Indiv`: Mediana de $T_0$ de los registros válidos.
36. `T Grup`: Mediana de las medianas grupales.
37. `Vrdcto Grp`: Veredicto del grupo.
38. `Recom Grp`: Recomendación de grupo.
39. `Conclusion T`: Período concluido definitivo.

### Hoja 2: `Analisis Grupal` (7 Columnas)
`N`, `Long`, `Lat`, `Sitio`, `T0`, `A0`, `SESAME`.
Contiene exactamente 15 filas correspondientes a los 15 sitios activos.

---

## 7. Generación del Informe Oficial DOCX (Geosísmica)

El agente debe automatizar la construcción del documento Word mediante `python-docx` asegurando:
1. Portada formal con logotipo textual ARDASYSMIC y datos del proyecto Promart-CityMall.
2. Inclusión de las 8 tablas formateadas con cabeceras azul marino (`#1F4E79`) y texto blanco.
3. Incrustación de 18 figuras de alta resolución:
   - Figura 1: Mapa del polígono con ortofoto satelital.
   - Figura 2: Tabla/grid de fotografías de campo georreferenciadas.
   - Figuras 3 a 17: Curvas de relación H/V por sitio (`*_hvsr_periodo.png`).
   - Figura 18: Modelo 2D de isoperíodos fundamentales $T_0$ (s).
4. Redacción en tercera persona técnica e inclusión de las conclusiones multidisciplinarias.
