# Nardo Puro - Agregar perfumes desde links de cualquier tienda
# Lee los links de links.txt y por cada uno:
#   1. Descarga las fotos y los datos del producto (nombre, marca, descripción, ml)
#   2. Te pregunta lo que falte: nombre, marca, género, presentación, precio
#   3. Crea Productos\<Marca - Nombre>\ con las fotos en JPG y el texto para Instagram
#   4. Lo agrega a la web (productos.json + fotos en img\) y regenera el PDF
#
# En links.txt:
#   - Un link de producto por línea (Notino, Sephora, Amazon, tiendas Shopify, la página de la marca...).
#   - Si un sitio no deja bajar las fotos, pega debajo del link los links directos de las fotos
#     (clic derecho en la foto > "Copiar dirección de imagen") o la ruta de una foto en tu PC.
#     Esas líneas se agregan como fotos del producto de arriba.
#   - Las líneas que empiezan con # se ignoran.

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.Drawing, PresentationCore, WindowsBase

$Repo      = Split-Path $PSScriptRoot -Parent
$Raiz      = Split-Path $Repo -Parent
$Productos = Join-Path $Raiz 'Productos'
$Links     = Join-Path $Raiz 'links.txt'
$Json      = Join-Path $Repo 'productos.json'
$Registro  = Join-Path $Raiz 'registro.csv'
$ImgDir    = Join-Path $Repo 'img'
$UA        = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'
$Utf8      = New-Object System.Text.UTF8Encoding($false)
$Edge      = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
$Auto      = [bool]$env:NP_AUTO   # para pruebas: acepta todas las sugerencias sin preguntar
$MaxFotos  = 8

function Write-Titulo($t) { Write-Host ''; Write-Host "=== $t ===" -ForegroundColor Cyan }

function Preguntar($texto, $sugerido) {
  if ($Auto) { return $sugerido }
  $r = (Read-Host $texto).Trim()
  if ($r) { return $r } else { return $sugerido }
}

# ---------- Leer la página del producto ----------

function Get-Pagina($url) {
  try {
    $r = Invoke-WebRequest -Uri $url -UserAgent $UA -UseBasicParsing -TimeoutSec 30 -Headers @{
      'Accept' = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
      'Accept-Language' = 'es-ES,es;q=0.9,en;q=0.8'
    }
    $ms = New-Object IO.MemoryStream
    $r.RawContentStream.CopyTo($ms)
    $html = [Text.Encoding]::UTF8.GetString($ms.ToArray())
    if ($html -match 'og:image|ld\+json|cdn\.shopify\.com') { return $html }
  } catch {}
  # Plan B: abrir la página con Edge, para sitios que arman la página con JavaScript
  Write-Host '  Abriendo la página con Edge...' -ForegroundColor DarkGray
  $prof = Join-Path $env:TEMP 'np-edge'
  $out  = Join-Path $env:TEMP 'np-dom.html'
  $err  = Join-Path $env:TEMP 'np-dom.err'
  if (Test-Path $out) { Remove-Item $out }
  $pr = Start-Process -PassThru -WindowStyle Hidden -FilePath $Edge -RedirectStandardOutput $out -RedirectStandardError $err -ArgumentList '--headless=new', '--disable-gpu', "--user-data-dir=`"$prof`"", '--virtual-time-budget=15000', "--user-agent=`"$UA`"", '--dump-dom', "`"$url`""
  if (-not $pr.WaitForExit(45000)) {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*np-edge*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
  }
  for ($i = 0; $i -lt 10; $i++) {
    try { return [IO.File]::ReadAllText($out, [Text.Encoding]::UTF8) } catch { Start-Sleep -Milliseconds 500 }
  }
  return ''
}

# Sitios con verificación anti-robots: no se intenta saltarla, se pide la foto a mano
function Es-Bloqueo($html) {
  return (-not $html) -or ($html -match '(?i)<title>\s*(just a moment|un momento|attention required|access denied|acceso denegado|robot check|are you a robot)' -or $html -match 'cf-chl-|challenge-platform|captcha-delivery|px-captcha')
}

function Get-Meta($html, $key) {
  foreach ($m in [regex]::Matches($html, '<meta\b[^>]*>', 'IgnoreCase')) {
    $t = $m.Value
    if ($t -match "(?:property|name|itemprop)\s*=\s*[""']$([regex]::Escape($key))[""']") {
      if ($t -match 'content\s*=\s*"([^"]*)"' -or $t -match "content\s*=\s*'([^']*)'") {
        return [Net.WebUtility]::HtmlDecode($matches[1]).Trim()
      }
    }
  }
  return $null
}

function Find-Producto($o) {
  if ($null -eq $o) { return $null }
  if ($o -is [array]) {
    foreach ($x in $o) { $r = Find-Producto $x; if ($r) { return $r } }
    return $null
  }
  if ($o -is [pscustomobject]) {
    if ((@($o.'@type') -join ',') -match 'Product') { return $o }
    foreach ($k in '@graph', 'mainEntity', 'itemListElement', 'item') {
      if ($o.$k) { $r = Find-Producto $o.$k; if ($r) { return $r } }
    }
  }
  return $null
}

function Lista-Urls($v) {
  $out = @()
  foreach ($x in @($v)) {
    if ($x -is [string]) { $out += $x }
    elseif ($x.url) { $out += $x.url }
    elseif ($x.contentUrl) { $out += $x.contentUrl }
    elseif ($x.src) { $out += $x.src }
  }
  return $out
}

function Url-Completa($u, $base) {
  if (-not $u) { return $null }
  $u = [Net.WebUtility]::HtmlDecode($u.Trim())
  if ($u.StartsWith('//')) { return "https:$u" }
  try { return ([Uri]::new([Uri]$base, $u)).AbsoluteUri } catch { return $null }
}

function Limpiar-Html($html) {
  if (-not $html) { return '' }
  $t = $html -replace '<(br|/p|/li|/h\d|/div)[^>]*>', "`n"
  $t = $t -replace '<[^>]+>', ''
  $t = [Net.WebUtility]::HtmlDecode($t)
  $t = ($t -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join "`n"
  return $t.Trim()
}

function Leer-Producto($url) {
  $html = Get-Pagina $url
  $p = [ordered]@{ nombre = ''; marca = ''; desc = ''; fotos = @(); precioOrigen = ''; bloqueado = $false }
  if (Es-Bloqueo $html) { $p.bloqueado = $true; return $p }

  # 1) Tiendas Shopify: tienen un .js con todos los datos y todas las fotos
  if ($html -match 'cdn\.shopify\.com|Shopify\.shop') {
    try {
      $base = ($url -split '[?#]')[0].TrimEnd('/')
      $s = Invoke-WebRequest -Uri "$base.js" -UserAgent $UA -UseBasicParsing -TimeoutSec 30
      $ms = New-Object IO.MemoryStream; $s.RawContentStream.CopyTo($ms)
      $j = [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
      $p.nombre = $j.title; $p.marca = $j.vendor; $p.desc = Limpiar-Html $j.description
      $p.fotos += @($j.images)
      if ($j.price) { $p.precioOrigen = '{0:N2}' -f ($j.price / 100) }
    } catch {}
  }

  # 2) Datos estructurados (JSON-LD) que usan casi todas las tiendas
  foreach ($m in [regex]::Matches($html, '(?is)<script[^>]*application/ld\+json[^>]*>(.*?)</script>')) {
    try { $ld = $m.Groups[1].Value.Trim() | ConvertFrom-Json } catch { continue }
    $prod = Find-Producto $ld
    if (-not $prod) { continue }
    if (-not $p.nombre) { $p.nombre = $prod.name }
    if (-not $p.marca -and $prod.brand) { $p.marca = if ($prod.brand -is [string]) { $prod.brand } else { $prod.brand.name } }
    if (-not $p.desc) { $p.desc = Limpiar-Html $prod.description }
    $p.fotos += Lista-Urls $prod.image
    if (-not $p.precioOrigen -and $prod.offers) {
      $o = @($prod.offers)[0]
      $pr = if ($o.price) { $o.price } elseif ($o.lowPrice) { $o.lowPrice } else { $null }
      if ($pr) { $p.precioOrigen = "$pr $($o.priceCurrency)".Trim() }
    }
    if ($prod.hasVariant) { foreach ($v in @($prod.hasVariant)) { $p.fotos += Lista-Urls $v.image } }
  }

  # 3) Galerías de WooCommerce y otras
  foreach ($m in [regex]::Matches($html, 'data-(?:large_image|zoom-image|zoom|large|full)\s*=\s*"([^"]+)"')) { $p.fotos += $m.Groups[1].Value }

  # 4) Etiquetas para redes sociales (og:)
  if (-not $p.nombre) { $p.nombre = Get-Meta $html 'og:title' }
  if (-not $p.nombre -and $html -match '(?is)<title[^>]*>(.*?)</title>') { $p.nombre = [Net.WebUtility]::HtmlDecode($matches[1]).Trim() }
  if (-not $p.marca) { $p.marca = Get-Meta $html 'product:brand' }
  if (-not $p.desc) { $p.desc = Get-Meta $html 'og:description' }
  if (-not $p.desc) { $p.desc = Get-Meta $html 'description' }
  if (-not $p.precioOrigen) {
    $pa = Get-Meta $html 'product:price:amount'
    if ($pa) { $p.precioOrigen = "$pa $(Get-Meta $html 'product:price:currency')".Trim() }
  }
  $p.fotos += @(Get-Meta $html 'og:image'; Get-Meta $html 'og:image:secure_url'; Get-Meta $html 'twitter:image')

  # Quita repetidas (la misma foto suele venir con distinto tamaño o servidor)
  $vistas = @{}
  $p.fotos = @($p.fotos | ForEach-Object { Url-Completa $_ $url } | Where-Object { $_ -and $_ -notmatch '\.svg(\?|$)' -and $_ -notmatch '(?i)og-default|placeholder|logo' } | Where-Object {
    $k = ((($_ -split '[?#]')[0] -split '/')[-1]).ToLower() -replace '_\d+x\d*(?=\.)', '' -replace '\.(jpe?g|png|webp|avif)$', ''
    if ($vistas.ContainsKey($k)) { $false } else { $vistas[$k] = 1; $true }
  })
  $p.nombre = if ($p.nombre) { ($p.nombre -replace '\s+', ' ').Trim() } else { '' }
  $p.marca = if ($p.marca) { $p.marca.Trim() } else { '' }
  return $p
}

# ---------- Fotos ----------

function Bajar-Foto($origen, $destino, $referer) {
  if (Test-Path -LiteralPath $origen) { Copy-Item -LiteralPath $origen $destino; return }
  $wc = New-Object Net.WebClient
  $wc.Headers.Add('User-Agent', $UA)
  $wc.Headers.Add('Accept', 'image/jpeg,image/png,image/webp;q=0.8,image/*;q=0.5')
  if ($referer) { $wc.Headers.Add('Referer', $referer) }
  try { $wc.DownloadFile($origen, $destino) } finally { $wc.Dispose() }
}

function Abrir-Imagen($path) {
  $bytes = [IO.File]::ReadAllBytes($path)
  try { return [System.Drawing.Image]::FromStream((New-Object IO.MemoryStream(, $bytes))) } catch {}
  # WebP, AVIF y otros formatos: se abren con el lector de imágenes de Windows
  try {
    $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create((New-Object IO.MemoryStream(, $bytes)), 'None', 'OnLoad')
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($dec.Frames[0]))
    $ms = New-Object IO.MemoryStream; $enc.Save($ms); $ms.Position = 0
    return [System.Drawing.Image]::FromStream($ms)
  } catch { return $null }
}

$JpgCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }

function Guardar-Jpg($img, $destino, $maxW, $calidad) {
  $w = [Math]::Min($maxW, $img.Width); $h = [int]($img.Height * $w / $img.Width)
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.DrawImage($img, 0, 0, $w, $h)
  $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
  $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$calidad)
  $bmp.Save($destino, $JpgCodec, $ep)
  $g.Dispose(); $bmp.Dispose()
}

# ---------- Adivinar datos del perfume ----------

function Adivinar-Genero($t) {
  # Privé y otras tiendas marcan (M) masculino, (F)/(W) femenino, (U) unisex
  if ($t -match '\((U|Unisex)\)') { return 'unisex' }
  if ($t -match '\((F|W|D)\)') { return 'mujer' }
  if ($t -match '\((M|H)\)') { return 'hombre' }
  $t = $t.ToLower()
  if ($t -match 'unisex') { return 'unisex' }
  if ($t -match '\b(pour femme|femme|for women|women|woman|mujer|dama|her|elle|donna)\b') { return 'mujer' }
  if ($t -match '\b(pour homme|homme|for men|men|man|hombre|caballero|him|uomo)\b') { return 'hombre' }
  return ''
}

function Adivinar-Tipo($t) {
  if ($t -match '(?i)extrait|elixir') { return 'Extrait de Parfum' }
  if ($t -match '(?i)eau de parfum|\bedp\b') { return 'Eau de Parfum' }
  if ($t -match '(?i)eau de toilette|\bedt\b') { return 'Eau de Toilette' }
  if ($t -match '(?i)eau de cologne|\bedc\b|colonia') { return 'Eau de Cologne' }
  if ($t -match '(?i)\bparfum\b') { return 'Parfum' }
  if ($t -match '(?i)body mist|splash') { return 'Body Mist' }
  return ''
}

function Adivinar-Ml($t) {
  $m = [regex]::Match($t, '(?i)(\d{1,3}(?:[.,]\d)?)\s?ml')
  if ($m.Success) { return "$($m.Groups[1].Value) ml" }
  $m = [regex]::Match($t, '(?i)(\d(?:[.,]\d)?)\s?(?:fl\.?\s?oz|oz)')
  if ($m.Success) { return "$([Math]::Round([double]($m.Groups[1].Value -replace ',', '.') * 29.5735)) ml" }
  return ''
}

function Nombre-Carpeta($texto) {
  $n = ($texto -replace '[<>:"/\\|?*]', '' -replace '\s+', ' ').Trim().TrimEnd('.')
  if ($n.Length -gt 70) { $n = $n.Substring(0, 70).Trim() }
  return $n
}

function Slug($t) { return ($t.ToLower().Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', '' -replace '[^a-z0-9]+', '').Trim() }

if ($env:NP_SOLO_FUNCIONES) { return }   # para pruebas: solo cargar las funciones

# ---------- Inicio ----------
foreach ($d in $Productos, $ImgDir) { if (-not (Test-Path $d)) { New-Item -ItemType Directory $d | Out-Null } }

$lineas = [IO.File]::ReadAllLines($Links, [Text.Encoding]::UTF8)
# Agrupa: cada link de producto + las fotos extra que tenga debajo
$pedidos = @()
foreach ($l in $lineas) {
  $t = $l.Trim()
  if (-not $t -or $t.StartsWith('#')) { continue }
  $esFoto = ($t -match '(?i)\.(jpe?g|png|webp|avif|gif)(\?|$)') -or (Test-Path -LiteralPath $t -PathType Leaf)
  if ($esFoto -and $pedidos.Count) { $pedidos[-1].extra += $t }
  elseif ($esFoto) { $pedidos += @{ link = $null; extra = @($t); lineas = @($l) } }
  else { $pedidos += @{ link = $t; extra = @(); lineas = @($l) } }
  if ($esFoto -and $pedidos.Count) { $pedidos[-1].lineas += $l }
}
if ($pedidos.Count -eq 0) {
  Write-Host 'No hay links en links.txt. Pega los links (uno por línea), guarda el archivo y vuelve a correr esto.' -ForegroundColor Yellow
  if (-not $Auto) { Start-Process notepad.exe $Links }
  exit
}

$data = [IO.File]::ReadAllText($Json, [Text.Encoding]::UTF8) | ConvertFrom-Json
if (-not $data.productos) { $data.productos = @() }
$yaRegistrados = @()
if (Test-Path $Registro) { $yaRegistrados = @(Import-Csv $Registro -Encoding UTF8 | ForEach-Object { $_.link }) }
$fallidos = @()
$agregados = @()

foreach ($ped in $pedidos) {
  $link = $ped.link
  Write-Titulo $(if ($link) { $link } else { 'Producto con fotos propias' })
  try {
    if ($link -and $yaRegistrados -contains $link) {
      Write-Host 'Este link ya está en el catálogo, se salta.' -ForegroundColor DarkYellow
      continue
    }
    $info = if ($link) { Leer-Producto $link } else { @{ nombre = ''; marca = ''; desc = ''; fotos = @(); precioOrigen = '' } }
    # Las fotos que pegaste a mano van primero
    $fotos = @(@($ped.extra) + @($info.fotos) | Select-Object -Unique | Select-Object -First $MaxFotos)
    Write-Host "Encontrado: $($info.nombre)" -ForegroundColor White
    Write-Host "Marca: $($info.marca)   Fotos: $($fotos.Count)   Precio en la tienda: $($info.precioOrigen)" -ForegroundColor DarkGray
    if ($info.bloqueado -and $fotos.Count -eq 0) { throw 'Esta tienda no deja que programas lean la página. Pega debajo del link la dirección de la foto (clic derecho > Copiar dirección de imagen) o guarda la foto en tu PC y pega la ruta.' }
    if ($fotos.Count -eq 0) { throw 'No se encontraron fotos. Pega debajo del link la dirección de la foto (clic derecho > Copiar dirección de imagen).' }

    $todo = "$($info.nombre) $($info.desc)"
    $marca = Preguntar "Marca (Enter = '$($info.marca)')" $info.marca
    $nombreSug = $info.nombre -replace '^(?i)(perfume|fragancia|colonia|set|estuche)\s+', ''
    $nombreSug = $nombreSug -replace '\s*\((M|F|W|D|H|U|Unisex)\)', '' -replace '\s*/\s*\d+(\.\d)?\s?ml\b', ''
    if ($marca -and $nombreSug -match "^(?i)$([regex]::Escape($marca))\s*[-–:|]?\s*(.+)$") { $nombreSug = $matches[1] }
    $nombreSug = ($nombreSug -replace '(?i)\s*[-–|,]?\s*(eau de (parfum|toilette|cologne)|edp|edt|extrait de parfum|parfum|\d+(\.\d)?\s?ml|for (wo)?men|pour (homme|femme)|spray)\b.*$', '').Trim(' ', '-', '–', ',')
    if (-not $nombreSug) { $nombreSug = $info.nombre }
    $nombre = Preguntar "Nombre del perfume (Enter = '$nombreSug')" $nombreSug
    if (-not $nombre) { throw 'El producto necesita un nombre.' }
    if (-not $marca) { $marca = 'Otras marcas' }
    # Misma marca escrita igual siempre: "DIOR" y "Dior" quedan juntas
    $igual = @($data.productos | Where-Object { $_.marca -ieq $marca } | Select-Object -First 1)
    if ($igual.Count) { $marca = $igual[0].marca }
    elseif ($marca.Length -ge 4 -and $marca -ceq $marca.ToUpper()) { $marca = (Get-Culture).TextInfo.ToTitleCase($marca.ToLower()) }

    $gSug = Adivinar-Genero $todo
    $genero = $null
    while (-not $genero) {
      $r = Preguntar "Género: 1) Hombre  2) Mujer  3) Unisex  (Enter = '$gSug')" $gSug
      switch -regex ($r) { '^(1|h)' { $genero = 'hombre' } '^(2|m)' { $genero = 'mujer' } '^(3|u)' { $genero = 'unisex' } }
      if ($Auto -and -not $genero) { $genero = 'unisex' }
    }
    $tSug = Adivinar-Tipo $todo
    $tipo = Preguntar "Concentración (EDP, EDT, Parfum...) (Enter = '$tSug')" $tSug
    switch -regex ($tipo) { '^(?i)edp$' { $tipo = 'Eau de Parfum' } '^(?i)edt$' { $tipo = 'Eau de Toilette' } '^(?i)edc$' { $tipo = 'Eau de Cologne' } }
    $mlSug = Adivinar-Ml $todo
    $ml = Preguntar "Presentación (ej. 100 ml) (Enter = '$mlSug')" $mlSug
    if ($ml -match '^\d+([.,]\d)?$') { $ml = "$ml ml" }

    $precio = $null
    # De Privé se toma su precio (ya viene en lempiras); de otras tiendas se escribe a mano
    $esPrive = $link -match 'priveperfumes\.com'
    $pSug = if ($esPrive -and $info.precioOrigen) { ($info.precioOrigen -replace '[^\d.]', '') -replace '\.00$', '' } else { '' }
    if (-not $esPrive -and $info.precioOrigen) { Write-Host "  (precio en esa tienda: $($info.precioOrigen), no es lempiras)" -ForegroundColor DarkGray }
    while ($true) {
      $msg = if ($pSug) { "Precio en lempiras (Enter = L $pSug, el de Privé)" } else { 'Precio en lempiras, solo el número (Enter = "Consultar")' }
      $r = Preguntar $msg $pSug
      if (-not $r) { break }
      $n = $r -replace '[^\d.]', ''
      if ($n -match '^\d+(\.\d+)?$') { $precio = [double]$n; break }
      Write-Host 'Escribe solo el número, por ejemplo 2500' -ForegroundColor Yellow
    }

    $auto = ($info.desc -split "`n" | Where-Object { $_.Length -gt 30 } | Select-Object -First 1)
    if (-not $auto) { $auto = '' }
    if ($auto.Length -gt 180) { $auto = $auto.Substring(0, $auto.LastIndexOf(' ', 180)).TrimEnd(',', ';', ':') + '...' }
    if (-not $Auto -and $auto) { Write-Host "Descripción sugerida: $auto" -ForegroundColor DarkGray }
    $desc = Preguntar 'Descripción corta (Enter = usar la sugerida)' $auto

    # Fotos: carpeta para Instagram + versión liviana para la web
    $id = 1 + (@($data.productos | ForEach-Object { [int]$_.id }) + 0 | Measure-Object -Maximum).Maximum
    $dir = Join-Path $Productos (Nombre-Carpeta "$marca - $nombre")
    if (Test-Path $dir) { $dir = "$dir ($id)" }
    New-Item -ItemType Directory -Force $dir | Out-Null
    $ok = 0
    foreach ($u in $fotos) {
      $tmp = Join-Path $env:TEMP "np-foto.tmp"
      try {
        Bajar-Foto $u $tmp $link
        $img = Abrir-Imagen $tmp
        if (-not $img) { Write-Host "  (no se pudo abrir una foto, se salta)" -ForegroundColor DarkYellow; continue }
        if ($img.Width -lt 250 -or $img.Height -lt 250) { $img.Dispose(); continue }   # íconos y miniaturas
        Guardar-Jpg $img (Join-Path $dir ('{0:D2}.jpg' -f ($ok + 1))) 1440 92
        Guardar-Jpg $img (Join-Path $ImgDir "$id-$ok.jpg") 900 82
        $img.Dispose()
        $ok++
        Write-Host "  foto $ok"
      } catch {
        Write-Host "  (no se pudo bajar una foto: $($_.Exception.Message))" -ForegroundColor DarkYellow
        if ($env:NP_DEBUG) { Write-Host "  DEBUG dir=[$dir] img=[$ImgDir] id=[$id] $($_.ScriptStackTrace)" }
      } finally { if (Test-Path $tmp) { Remove-Item $tmp } }
    }
    if ($ok -eq 0) {
      Remove-Item -Recurse $dir
      throw 'Ninguna foto se pudo bajar. Pega debajo del link la dirección directa de la foto o guárdala en tu PC y pega la ruta.'
    }

    # Texto para Instagram
    $generoTxt = @{ hombre = 'para él'; mujer = 'para ella'; unisex = 'unisex' }[$genero]
    $precioTxt = if ($precio) { 'L ' + $precio.ToString('N0', [Globalization.CultureInfo]'en-US') } else { 'Consulta el precio por WhatsApp' }
    $detalle = (@($tipo, $ml) | Where-Object { $_ }) -join ' · '
    $tagGen = @{ hombre = '#perfumeshombre #fraganciasmasculinas'; mujer = '#perfumesmujer #fraganciasfemeninas'; unisex = '#perfumesunisex' }[$genero]
    $ig = @"
✨ $marca – $nombre ✨
$detalle · $generoTxt

$desc

💯 Perfume 100% original
💰 $precioTxt
📲 Pedidos por WhatsApp: $($data.whatsappTexto)

#nardopuro #perfumes #perfumesoriginales #honduras #$(Slug $marca) $tagGen
"@
    $txt = @"
$marca – $nombre

===== TEXTO PARA INSTAGRAM (copiar y pegar) =====

$ig

===== DESCRIPCIÓN COMPLETA DE LA TIENDA =====

$($info.desc)

===== DE DÓNDE SALIÓ =====
Link: $link
Precio en la tienda: $($info.precioOrigen)
Agregado: $(Get-Date -Format 'yyyy-MM-dd')
"@
    [IO.File]::WriteAllText((Join-Path $dir 'descripcion.txt'), $txt, $Utf8)

    # Solo datos públicos: la web y productos.json los puede ver cualquiera
    $data.productos += [pscustomobject]@{
      id = $id; nombre = $nombre; marca = $marca; genero = $genero; tipo = $tipo; ml = $ml
      precio = $precio; desc = $desc; fotos = $ok; agotado = $false
    }

    # Registro PRIVADO de todo lo descargado (link, tienda, precio de la tienda), fuera de la web.
    # Sirve para cruzarlo con las facturas.
    $tienda = if ($link) { ([Uri]$link).Host -replace '^www\.', '' } else { '' }
    if (-not (Test-Path $Registro)) {
      [IO.File]::WriteAllText($Registro, "fecha,id,marca,nombre,genero,ml,tienda,precio_tienda,precio_venta,link`r`n", (New-Object System.Text.UTF8Encoding($true)))
    }
    $csv = { param($v) '"' + ("$v" -replace '"', '""') + '"' }
    $fila = @((Get-Date -Format 'yyyy-MM-dd'), $id, $marca, $nombre, $genero, $ml, $tienda, $info.precioOrigen, $precio, $link) | ForEach-Object { & $csv $_ }
    [IO.File]::AppendAllText($Registro, ($fila -join ',') + "`r`n", $Utf8)
    $agregados += "$marca $nombre"
    Write-Host "Listo: $dir" -ForegroundColor Green
  } catch {
    Write-Host "No se pudo: $($_.Exception.Message)" -ForegroundColor Red
    $fallidos += $ped.lineas
  }
}

# Deja en links.txt solo los comentarios y los que fallaron
$quedan = @($lineas | Where-Object { $_.Trim().StartsWith('#') }) + $fallidos
[IO.File]::WriteAllLines($Links, [string[]]$quedan, $Utf8)

if ($agregados.Count -eq 0) {
  Write-Titulo 'Terminado'
  Write-Host 'No se agregó ningún producto.'
  exit
}

& (Join-Path $PSScriptRoot 'actualizar.ps1') -Data $data -Nuevos $agregados
