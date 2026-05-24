param(
  [string]$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..\..")).Path
)

$ErrorActionPreference = "Stop"

$previewDir = Join-Path $Root "docs\previews"
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

$items = @(
  @{
    Pptx = "outputs\style-showcase\template-elements\template-elements-gallery.pptx"
    Slide = 1
    Out = "strict-template-gallery.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-inspired\beamer-sysu-blue-showcase.pptx"
    Slide = 4
    Out = "beamer-inspired-blue.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-inspired\beamer-sysu-green-showcase.pptx"
    Slide = 5
    Out = "beamer-inspired-green.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-inspired\beamer-sysu-red-showcase.pptx"
    Slide = 6
    Out = "beamer-inspired-red.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-candidates\simpleplus-sysu-clean-showcase.pptx"
    Slide = 3
    Out = "candidate-simpleplus.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-candidates\ustc-thu-sysu-institutional-showcase.pptx"
    Slide = 3
    Out = "candidate-ustc-thu.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-candidates\moloch-sysu-minimal-showcase.pptx"
    Slide = 3
    Out = "candidate-moloch.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-candidates\sleek-sysu-research-showcase.pptx"
    Slide = 3
    Out = "candidate-sleek.png"
  },
  @{
    Pptx = "outputs\style-showcase\beamer-candidates\river-sysu-atelier-showcase.pptx"
    Slide = 3
    Out = "candidate-river.png"
  }
)

$powerPoint = New-Object -ComObject PowerPoint.Application
$opened = @()
try {
  foreach ($item in $items) {
    $src = Join-Path $Root $item.Pptx
    if (-not (Test-Path -LiteralPath $src)) {
      throw "Missing PPTX: $src"
    }
    $dst = Join-Path $previewDir $item.Out
    $presentation = $powerPoint.Presentations.Open($src, 0, 0, 0)
    $opened += $presentation
    try {
      $presentation.Slides.Item([int]$item.Slide).Export($dst, "PNG", 1600, 900) | Out-Null
      Write-Host "Wrote $dst"
    }
    finally {
      $presentation.Close()
    }
  }
}
finally {
  foreach ($presentation in $opened) {
    try {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    catch {
    }
  }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
