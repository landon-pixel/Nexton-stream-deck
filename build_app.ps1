$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$iconPng = Join-Path $PSScriptRoot "assets\logo\Nextdeck logo.png"
$iconIco = Join-Path $PSScriptRoot "assets\icons\NextDeck.ico"
$distDir = Join-Path $PSScriptRoot "dist"
$exePath = Join-Path $distDir "NextDeck.exe"

if (Test-Path $iconPng) {
    Add-Type -AssemblyName System.Drawing
    $bitmap = [System.Drawing.Bitmap]::FromFile($iconPng)
    $icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
    $stream = [System.IO.File]::Open($iconIco, [System.IO.FileMode]::Create)
    $icon.Save($stream)
    $stream.Close()
    $icon.Dispose()
    $bitmap.Dispose()
}

if (Test-Path $exePath) {
    try {
        Remove-Item -LiteralPath $exePath -Force
    }
    catch {
        throw "Could not replace dist\NextDeck.exe. Close NextDeck if it is open, then run build_app.ps1 again."
    }
}

$args = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onefile",
    "--name", "NextDeck",
    "--add-data", "assets;assets"
)

if (Test-Path $iconIco) {
    $args += @("--icon", $iconIco)
}

$args += "app.py"

python @args
