<#
================================================================================
SCRIPT POWERSHELL: Flujo Maestro de Análisis HVSR (Individual y Grupal)
================================================================================
PROYECTO: ARDASYSMIC - Promart-CityMall
DESCRIPCIÓN:
  Este script coordina la ejecución automatizada del análisis espectral HVSR
  de registros MiniSEED mediante TremorSys CLI.
  
  METODOLOGÍA:
  1. Análisis Individual:
     - Ejecuta TremorSys CLI archivo por archivo (.mseed).
     - Si el sensor es de aceleración (ej. REAEC) y las componentes tienen
       unidades m/s², se activa automáticamente '--units acceleration'.
  2. Detección de Candidatos a Descarte (Clustering):
     - Identifica el período fundamental característico de suelo (T0 < 2.0 s).
     - Descarta picos secundarios y artefactos de ruido ambiental (T0 >= 2.0 s).
     - Calcula T mean individual (mediana de registros válidos) y T mean grupal.
     - Emite veredicto: OK (congruente <= 0.15s) o REVISAR / GRUPO COMPLETO.
  3. Análisis Grupal Filtrado:
     - Extrae el bloque continuo más grande de registros válidos (>= 2 archivos).
     - Ejecuta TremorSys CLI en modo grupal (--group-input y --group-size N).
     - Evita errores de huecos o solapamiento temporal.
  4. Generación de Bitácora Excel:
     - Produce 'Resultados de Nakamuras.xlsx' con 39 columnas en 'Analisis Individual'
       y 'Analisis Grupal' con sombreado de candidatos en gris.
================================================================================
#>

param(
    [string]$PythonExe = "python"
)

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "INICIANDO PROCESAMIENTO NAKAMURA HVSR - ARDASYSMIC" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan

$scriptPath = Join-Path $PSScriptRoot "build_nakamura_results.py"

if (-not (Test-Path $scriptPath)) {
    Write-Error "No se encontro el script maestro: $scriptPath"
    exit 1
}

Write-Host "`nEjecutando script maestro en Python..." -ForegroundColor Yellow
& $PythonExe $scriptPath

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n======================================================================" -ForegroundColor Green
    Write-Host "PROCESAMIENTO FINALIZADO CON EXITO" -ForegroundColor Green
    Write-Host "Archivo generado: Resultados de Nakamuras.xlsx" -ForegroundColor Green
    Write-Host "======================================================================" -ForegroundColor Green
} else {
    Write-Error "Error en la ejecucion del script maestro (Exit Code: $LASTEXITCODE)"
    exit $LASTEXITCODE
}
