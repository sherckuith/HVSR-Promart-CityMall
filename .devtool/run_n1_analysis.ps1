# Script para analizar individualmente cada mseed de N1 con TremorSys CLI
# No modifica el repositorio de TremorSys. Solo usa el binario compilado.

$tremorsysDir = "D:\2026\ARDASYSMIC\ARDASYSMIC_26\TremorSys"
$cli = "$tremorsysDir\target\release\tremorsys-cli.exe"
$profile = "$tremorsysDir\examples\report_profile.toml"
$baseOutputDir = "d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\.devtool\results_individual"

# N1 mseed files (no descartados)
$n1Dir = "d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\Sitios Nakamura\N1\Registros"
$mseedFiles = @(
    "R13D7_182000_278.mseed",
    "R13D7_182500_278.mseed",
    "R13D7_183000_278.mseed",
    "R13D7_183500_278.mseed",
    "R13D7_184000_278.mseed"
)

# Crear directorio de salida
New-Item -ItemType Directory -Force -Path $baseOutputDir | Out-Null

$results = @()

foreach ($mseed in $mseedFiles) {
    $inputPath = Join-Path $n1Dir $mseed
    $outSubdir = Join-Path $baseOutputDir ("N1_" + $mseed.Replace(".mseed",""))
    
    Write-Host "`n=== Procesando: $mseed ===" -ForegroundColor Cyan
    
    # Ejecutar CLI en modo individual, sin imagenes para rapido, solo datos
    & $cli --input $inputPath --output $outSubdir --profile $profile --no-record --no-noise --no-watermark 2>&1
    
    # Leer el JSON de resultados
    $baseName = $mseed.Replace(".mseed","")
    $jsonFile = Join-Path $outSubdir "${baseName}__resultados.json"
    
    if (Test-Path $jsonFile) {
        $json = Get-Content $jsonFile -Raw | ConvertFrom-Json
        
        $obj = [PSCustomObject]@{
            sitio = "N1"
            mseed = $mseed
            canales = $json.traces.Count
            fs_hz = $json.traces[0].sample_rate_hz
            duracion_s = $json.traces[0].duration_s
            ventanas = $json.num_windows
            ventana_s = $json.window_length_s
            f0_hz = $json.f0_hz
            t0_s = $json.t0_s
            a0 = $json.a0
            kg = $json.kg
            confiable = $json.reliable
            num_cycles = $json.num_cycles
            f0_low_hz = $json.statistics.f0_low_hz
            f0_high_hz = $json.statistics.f0_high_hz
            f0_median_hz = $json.statistics.f0_median_hz
            t_mean_hz = $json.statistics.t_mean_hz
            units_status = $json.units_status
            station = $json.traces[0].station
        }
        
        $results += $obj
        Write-Host "  f0=$($obj.f0_hz) Hz, T0=$($obj.t0_s) s, A0=$($obj.a0), Kg=$($obj.kg), reliable=$($obj.confiable)" -ForegroundColor Green
    } else {
        Write-Host "  ERROR: No se encontro $jsonFile" -ForegroundColor Red
    }
}

Write-Host "`n=== RESUMEN N1 ===" -ForegroundColor Yellow
$results | Format-Table sitio, mseed, f0_hz, t0_s, a0, kg, confiable, ventanas, duracion_s -AutoSize

# Exportar a JSON para leer despues
$results | ConvertTo-Json -Depth 5 | Out-File (Join-Path $baseOutputDir "N1_summary.json") -Encoding UTF8
Write-Host "`nResultados guardados en: $baseOutputDir\N1_summary.json"
