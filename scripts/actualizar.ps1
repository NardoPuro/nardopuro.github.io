# Nardo Puro - Guarda los productos, regenera productos.js y el PDF, y pregunta si subir a la web.
# Se corre solo al final de agregar-productos.ps1, o a mano después de cambiar precios en productos.json.
param($Data, [string[]]$Nuevos = @())

$ErrorActionPreference = 'Stop'
$Repo = Split-Path $PSScriptRoot -Parent
$Json = Join-Path $Repo 'productos.json'
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$Edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

if (-not $Data) { $Data = [IO.File]::ReadAllText($Json, [Text.Encoding]::UTF8) | ConvertFrom-Json }
$Data.productos = @($Data.productos)
$out = $Data | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText($Json, $out, $Utf8)
[IO.File]::WriteAllText((Join-Path $Repo 'productos.js'), "window.NP = $out;`n", $Utf8)

Write-Host ''
Write-Host '=== Generando el PDF ===' -ForegroundColor Cyan
$prof = Join-Path $env:TEMP 'np-edge-pdf'
$src = 'file:///' + ((Join-Path $Repo 'fuente\catalogo.html') -replace '\\', '/')
$pdf = Join-Path $Repo 'catalogo-nardo-puro.pdf'
Start-Process -Wait -WindowStyle Hidden -FilePath $Edge -ArgumentList '--headless=new', '--disable-gpu', '--allow-file-access-from-files', '--no-pdf-header-footer', "--user-data-dir=`"$prof`"", '--virtual-time-budget=20000', "--print-to-pdf=`"$pdf`"", "`"$src`""
Write-Host "PDF actualizado: $pdf"

if ($env:NP_AUTO -or -not (Test-Path (Join-Path $Repo '.git'))) { exit }

Write-Host ''
Write-Host '=== Subir a la web ===' -ForegroundColor Cyan
if ($Nuevos.Count) { Write-Host "Productos nuevos: $($Nuevos -join ', ')" }
$r = (Read-Host '¿Subir los cambios a la página web? (S/N)').Trim()
if ($r -match '^[sS]') {
  Push-Location $Repo
  try {
    git add -A
    $msg = if ($Nuevos.Count) { "Agregar: $($Nuevos -join ', ')" } else { 'Actualizar catálogo' }
    git commit -q -m $msg
    git push -q
    Write-Host 'Subido. La página se actualiza en 1 o 2 minutos.' -ForegroundColor Green
  } finally { Pop-Location }
} else {
  Write-Host 'No se subió nada. Los cambios quedan guardados en tu PC.'
}
