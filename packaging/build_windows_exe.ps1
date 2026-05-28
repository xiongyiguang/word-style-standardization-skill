$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).ProviderPath
Push-Location $RepoRoot

try {
    $TemplatePath = (Get-ChildItem -Path (Join-Path $RepoRoot "assets") -Filter "*.docx" | Select-Object -First 1).FullName
    if (-not $TemplatePath) {
        throw "No .docx template found in assets."
    }
    $BuildDir = Join-Path $RepoRoot "build\pyinstaller"
    $BuildTemplatePath = Join-Path $BuildDir "standard_template.docx"
    New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
    Copy-Item -Force $TemplatePath $BuildTemplatePath

    py -3 -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    py -3 -m pip install -r requirements.txt pyinstaller
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    py -3 -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name "WordStyleStandardizer" `
        --add-data="$BuildTemplatePath`:assets" `
        --distpath "dist\windows" `
        --workpath "build\pyinstaller" `
        --specpath "build\pyinstaller" `
        "scripts\word_style_standardizer_gui.py"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    Write-Host ""
    Write-Host "Generated: dist\windows\WordStyleStandardizer.exe"
}
finally {
    Pop-Location
}
