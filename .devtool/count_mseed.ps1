$baseDir = "d:\2026\ARDASYSMIC\ARDASYSMIC_26\Trabajos\Promart-CityMall\Sitios Nakamura"
$sites = Get-ChildItem -Path $baseDir -Directory | Sort-Object Name
foreach ($site in $sites) {
    $regDir = Join-Path $site.FullName "Registros"
    $descDir = Join-Path $regDir "Descartados"
    $mseedCount = 0
    $descCount = 0
    if (Test-Path $regDir) {
        $mseedCount = @(Get-ChildItem -Path $regDir -Filter "*.mseed" -File).Count
    }
    if (Test-Path $descDir) {
        $descCount = @(Get-ChildItem -Path $descDir -Filter "*.mseed" -File).Count
    }
    Write-Host ("{0,-20} registros={1}  descartados={2}" -f $site.Name, $mseedCount, $descCount)
}
