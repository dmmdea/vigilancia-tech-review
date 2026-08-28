# Instalador para Windows. Deja la máquina lista para correr el skill.
#
#   powershell -ExecutionPolicy Bypass -File install.ps1
#
# Instala lo que falte con winget (Python, ffmpeg, LibreOffice), luego las
# dependencias de Python, y termina corriendo preflight.py — que es la única
# verificación en la que hay que confiar, porque prueba los backends de
# verdad y no lo que diga este script.
#
# No toca nada que ya esté instalado: si Python 3.10+ ya existe, no lo
# actualiza. Se puede volver a correr sin miedo.
[CmdletBinding()]
param([switch]$SinLibreOffice)

$ErrorActionPreference = "Stop"
$fallos = @()

function Escribe($msg, $color = "Gray") { Write-Host $msg -ForegroundColor $color }

function Existe($cmd) {
    $null -ne (Get-Command $cmd -ErrorAction SilentlyContinue)
}

function VersionPythonOk {
    if (-not (Existe "python")) { return $false }
    try {
        $v = (& python -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null)
        if (-not $v) { return $false }
        $p = $v.Split(".")
        return ([int]$p[0] -gt 3) -or ([int]$p[0] -eq 3 -and [int]$p[1] -ge 10)
    } catch { return $false }
}

function Instala($id, $nombre) {
    if (-not (Existe "winget")) {
        $script:fallos += "$nombre (winget no está disponible: instálalo a mano)"
        Escribe "  ! winget no está disponible; instala $nombre a mano" "Yellow"
        return
    }
    Escribe "  instalando $nombre ..." "Cyan"
    # --silent evita ventanas; --accept-* evita que se quede esperando un Enter
    & winget install --id $id --silent --accept-package-agreements `
        --accept-source-agreements --disable-interactivity 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        $script:fallos += "$nombre (winget salió con código $LASTEXITCODE)"
        Escribe "  ! winget no pudo instalar $nombre" "Yellow"
    }
}

Escribe "`n== Dependencias del sistema ==" "White"

if (VersionPythonOk) {
    $v = (& python -c "import sys; print('%d.%d.%d' % sys.version_info[:3])")
    Escribe "  ok  Python $v" "Green"
} else {
    Instala "Python.Python.3.12" "Python 3.12"
}

if (Existe "ffmpeg") { Escribe "  ok  ffmpeg" "Green" } else { Instala "Gyan.FFmpeg" "ffmpeg" }

if ($SinLibreOffice) {
    Escribe "  -   LibreOffice omitido (--SinLibreOffice)" "Gray"
} elseif ((Existe "soffice") -or (Test-Path "$env:ProgramFiles\Microsoft Office") -or
          (Test-Path "$env:ProgramFiles\LibreOffice")) {
    Escribe "  ok  backend de Office para pasar .pptx/.docx a PDF" "Green"
} else {
    Instala "TheDocumentFoundation.LibreOffice" "LibreOffice"
}

# winget deja los ejecutables nuevos en el PATH del sistema, pero ESTA sesión
# de PowerShell sigue con el PATH viejo: recargarlo evita el clásico
# "lo instalé y sigue diciendo que falta".
$env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [Environment]::GetEnvironmentVariable("Path", "User")

Escribe "`n== Dependencias de Python ==" "White"
if (VersionPythonOk) {
    & python -m pip install --quiet --upgrade pip
    & python -m pip install --quiet -r (Join-Path $PSScriptRoot "requirements.txt")
    if ($LASTEXITCODE -eq 0) { Escribe "  ok  openpyxl, pypdf, pymupdf" "Green" }
    else { $fallos += "pip install -r requirements.txt" }
} else {
    $fallos += "Python 3.10+ (cierra y vuelve a abrir la terminal y repite)"
    Escribe "  ! Python todavía no responde en esta terminal." "Yellow"
    Escribe "    Cierra PowerShell, ábrelo de nuevo y vuelve a correr este script." "Yellow"
}

Escribe "`n== Verificación real (preflight) ==" "White"
if (VersionPythonOk) {
    & python (Join-Path $PSScriptRoot "scripts\preflight.py")
    $preflight = $LASTEXITCODE
} else { $preflight = 1 }

if ($fallos.Count -gt 0) {
    Escribe "`nFALTA POR RESOLVER:" "Yellow"
    $fallos | ForEach-Object { Escribe "  - $_" "Yellow" }
}
if ($preflight -eq 0 -and $fallos.Count -eq 0) {
    Escribe "`nListo. La máquina puede correr el skill.`n" "Green"
    exit 0
} else {
    Escribe "`nQuedan cosas pendientes: lee arriba y vuelve a correr este script.`n" "Yellow"
    exit 1
}
