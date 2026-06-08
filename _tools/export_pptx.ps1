param(
    [Parameter(Mandatory=$true)][string]$Pptx,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [int]$Width = 1280,
    [int]$Height = 720
)
$ErrorActionPreference = "Stop"
$Pptx = (Resolve-Path $Pptx).Path
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path $OutDir).Path

$pp = New-Object -ComObject PowerPoint.Application
try {
    $deck = $pp.Presentations.Open($Pptx, $true, $false, $false)  # ReadOnly, Untitled, WithWindow=false
    $count = $deck.Slides.Count
    Write-Output "SLIDES=$count"
    $deck.Export($OutDir, "PNG", $Width, $Height)
    $deck.Close()
} finally {
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
