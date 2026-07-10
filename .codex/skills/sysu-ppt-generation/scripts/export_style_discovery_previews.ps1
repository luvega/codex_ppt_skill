param(
  [Parameter(Mandatory = $true)]
  [string]$Pptx,
  [Parameter(Mandatory = $true)]
  [string]$OutputDir
)

$ErrorActionPreference = "Stop"
$pptxPath = (Resolve-Path -LiteralPath $Pptx).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$outputPath = (Resolve-Path -LiteralPath $OutputDir).Path

$powerPoint = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {
  $presentation = $powerPoint.Presentations.Open($pptxPath, 0, 0, 0)
  if ($presentation.Slides.Count -ne 3) {
    throw "Visual Discovery PPTX must contain exactly three slides."
  }
  $names = @("option-a.png", "option-b.png", "option-c.png")
  for ($index = 1; $index -le 3; $index++) {
    $destination = Join-Path $outputPath $names[$index - 1]
    $presentation.Slides.Item($index).Export($destination, "PNG", 1600, 900) | Out-Null
    Write-Host "Wrote $destination"
  }
}
finally {
  if ($null -ne $presentation) {
    $presentation.Close()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
  }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
