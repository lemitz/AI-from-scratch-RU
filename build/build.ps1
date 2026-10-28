# Builds dist/neuroseti-s-nulya-A4.pdf and dist/neuroseti-s-nulya-A5.pdf from src/Нейросети_с_нуля.docx.
# Requirements: Windows, Microsoft Word, Python 3 with lxml (pip install lxml).
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root 'src\Нейросети_с_нуля.docx'
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force $dist | Out-Null

# Word opens files from network/mapped drives in Protected View (a modal dialog that hangs
# automation), so everything is staged through the local temp folder.
$stage = Join-Path $env:TEMP ('book-build-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $stage | Out-Null

try {
    $a4 = Join-Path $stage 'book-A4.docx'
    $a5 = Join-Path $stage 'book-A5.docx'
    Copy-Item -LiteralPath $src -Destination $a4

    Write-Host 'A5: deriving booklet layout...'
    python (Join-Path $PSScriptRoot 'make_a5.py') $a4 $a5
    if ($LASTEXITCODE -ne 0) { throw 'make_a5.py failed' }

    foreach ($name in 'A4', 'A5') {
        $docx = Join-Path $stage "book-$name.docx"
        $pdf = Join-Path $stage "book-$name.pdf"
        Write-Host "PDF: $name..."
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'export_pdf.ps1') $docx $pdf
        if (-not (Test-Path $pdf)) { throw "export failed: $name" }
        Copy-Item -LiteralPath $pdf -Destination (Join-Path $dist "neuroseti-s-nulya-$name.pdf") -Force
    }
}
finally {
    Remove-Item -Recurse -Force $stage -ErrorAction SilentlyContinue
}
Write-Host 'Done.'
