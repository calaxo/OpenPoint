# Dump de la vue "PowerPoint" du moteur d'animation via COM (object model).
# Sert de vérité terrain pour croiser avec le XML brut (ppt/slides/slideN.xml <p:timing>).
param(
  [string]$PptxPath = (Join-Path $PSScriptRoot "..\specimen\sample.pptx"),
  [string]$OutPath  = (Join-Path $PSScriptRoot "..\specimen\com-dump.json")
)

$PptxPath = (Resolve-Path $PptxPath).Path

# Tente de charger l'interop pour avoir les NOMS des enums (sinon on garde les ints)
$effectEnum = $null; $typeEnum = $null; $trigEnum = $null
try {
  Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint -ErrorAction Stop
  $effectEnum = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]
  $typeEnum   = [Microsoft.Office.Interop.PowerPoint.MsoAnimType]
  $trigEnum   = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]
} catch { Write-Warning "Interop PowerPoint introuvable, enums en int brut" }

function EnumName($enum, $val) {
  if ($null -eq $enum) { return $null }
  try { return [Enum]::GetName($enum, [int]$val) } catch { return $null }
}

$pp = New-Object -ComObject PowerPoint.Application
try {
  # Open(FileName, ReadOnly:=msoTrue, Untitled:=msoFalse, WithWindow:=msoFalse)
  $pres = $pp.Presentations.Open($PptxPath, -1, 0, 0)
  $out = New-Object System.Collections.Generic.List[object]
  foreach ($slide in $pres.Slides) {
    $seq = $slide.TimeLine.MainSequence
    $effects = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $seq.Count; $i++) {
      $e = $seq.Item($i)
      $behaviors = New-Object System.Collections.Generic.List[object]
      for ($b = 1; $b -le $e.Behaviors.Count; $b++) {
        $bh = $e.Behaviors.Item($b)
        $behaviors.Add(@{ type = [int]$bh.Type; typeName = EnumName $typeEnum $bh.Type })
      }
      $para = $null; try { $para = $e.Paragraph } catch {}
      $info = $e.EffectInformation
      $effects.Add([ordered]@{
        index        = $i
        shape        = $e.Shape.Name
        shapeId      = [int]$e.Shape.Id
        effectType   = [int]$e.EffectType
        effectName   = EnumName $effectEnum $e.EffectType
        exit         = [int]$e.Exit
        trigger      = [int]$e.Timing.TriggerType
        triggerName  = EnumName $trigEnum $e.Timing.TriggerType
        delay        = [double]$e.Timing.TriggerDelayTime
        duration     = [double]$e.Timing.Duration
        autoReverse  = [int]$e.Timing.AutoReverse
        repeatCount  = [double]$e.Timing.RepeatCount
        paragraph    = $para
        afterEffect  = [int]$info.AfterEffect
        textUnit     = [int]$info.TextUnitEffect
        buildByLevel = [int]$info.BuildByLevelEffect
        behaviors    = $behaviors
      })
    }
    $out.Add([ordered]@{ slide = $slide.SlideIndex; effectCount = $effects.Count; effects = $effects })
  }
  $pres.Close()
  $out | ConvertTo-Json -Depth 8 | Out-File $OutPath -Encoding utf8
  Write-Host "OK -> $OutPath ($($out.Count) diapos)"
} finally {
  $pp.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
