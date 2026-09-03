# Export the deck to PDF and per-slide PNGs using PowerPoint.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$pptx = Join-Path $root "deck\Gridlint.pptx"
$pdf  = Join-Path $root "deck\Gridlint.pdf"
$png  = Join-Path $root "deck\png"

$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($pptx, $true, $false, $false)
    $pres.SaveAs($pdf, 32)            # ppSaveAsPDF
    if (Test-Path $png) { Remove-Item $png -Recurse -Force }
    $pres.SaveAs($png, 18)            # ppSaveAsPNG
    $pres.Close()
    Write-Output "wrote $pdf and $png"
} finally {
    $app.Quit()
}
