# Convertit tous les EMF/WMF (vecteur Windows) sous <root> en PNG « <fichier>.png » (x2 pour la netteté).
# Préprocessing pour le moteur de rendu (imageHref sert le .png). Usage : convert-metafiles.ps1 <root>
param([string]$root = "renderer/_work")
Add-Type -AssemblyName System.Drawing
$files = Get-ChildItem -Path $root -Recurse -Include *.emf, *.wmf -ErrorAction SilentlyContinue
$ok = 0; $err = 0
foreach ($f in $files) {
  $out = $f.FullName + ".png"
  if (Test-Path $out) { $ok++; continue }
  try {
    $img = [System.Drawing.Image]::FromFile($f.FullName)
    $scale = 2.0
    $w = [int]($img.Width * $scale); $h = [int]($img.Height * $scale)
    if ($w -lt 1) { $w = 1 }; if ($h -lt 1) { $h = 1 }
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'; $g.InterpolationMode = 'HighQualityBicubic'
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.DrawImage($img, 0, 0, $w, $h)
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose(); $ok++
  } catch { Write-Output ("ERR {0} : {1}" -f $f.Name, $_.Exception.Message); $err++ }
}
Write-Output ("EMF/WMF -> PNG : {0} OK, {1} erreurs" -f $ok, $err)
