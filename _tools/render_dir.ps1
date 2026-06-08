param($InDir, $OutDir, $Width=640, $Height=360)
New-Item -ItemType Directory -Force $OutDir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$decks = Get-ChildItem $InDir -Filter *.pptx
foreach ($d in $decks) {
  try {
    $deck = $pp.Presentations.Open($d.FullName, $true, $false, $false)
    $png = Join-Path $OutDir ($d.BaseName + ".PNG")
    $deck.Slides.Item(1).Export($png, "PNG", $Width, $Height)
    $deck.Close()
  } catch {
    Write-Output ("FAIL {0}: {1}" -f $d.Name, $_.Exception.Message)
  }
}
$pp.Quit()
Write-Output ("rendered {0} decks" -f $decks.Count)
