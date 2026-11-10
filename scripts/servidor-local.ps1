# Sirve la web en http://localhost:8642 para verla en el navegador antes de subirla
param([int]$Port = 8642)
$root = Split-Path $PSScriptRoot -Parent
$types = @{ '.html' = 'text/html; charset=utf-8'; '.js' = 'text/javascript; charset=utf-8'; '.json' = 'application/json'; '.css' = 'text/css'
  '.svg' = 'image/svg+xml'; '.jpg' = 'image/jpeg'; '.png' = 'image/png'; '.webp' = 'image/webp'; '.pdf' = 'application/pdf' }
$l = New-Object Net.HttpListener
$l.Prefixes.Add("http://localhost:$Port/")
$l.Start()
Write-Host "Web en http://localhost:$Port/"
while ($l.IsListening) {
  $c = $l.GetContext()
  $path = [Uri]::UnescapeDataString($c.Request.Url.AbsolutePath.TrimStart('/'))
  if (-not $path) { $path = 'index.html' }
  $file = [IO.Path]::GetFullPath((Join-Path $root $path))
  if ($file.StartsWith($root) -and (Test-Path -LiteralPath $file -PathType Leaf)) {
    $b = [IO.File]::ReadAllBytes($file)
    $t = $types[[IO.Path]::GetExtension($file).ToLower()]
    $c.Response.ContentType = if ($t) { $t } else { 'application/octet-stream' }
    $c.Response.OutputStream.Write($b, 0, $b.Length)
  } else { $c.Response.StatusCode = 404 }
  $c.Response.Close()
}
