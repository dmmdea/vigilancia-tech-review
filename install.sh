#!/usr/bin/env bash
# Instalador para macOS y Linux. Deja la máquina lista para correr el skill.
#
#   bash install.sh
#
# Instala lo que falte (Python 3.10+, ffmpeg, LibreOffice) con el gestor de
# paquetes de la máquina, luego las dependencias de Python, y termina
# corriendo preflight.py — la única verificación en la que hay que confiar,
# porque prueba los backends de verdad y no lo que diga este script.
#
# No toca lo que ya esté instalado y se puede volver a correr sin miedo.
set -uo pipefail
cd "$(dirname "$0")"
FALLOS=()
SIN_LIBREOFFICE="${SIN_LIBREOFFICE:-0}"

hay() { command -v "$1" >/dev/null 2>&1; }

py_ok() {
  for c in python3 python; do
    if hay "$c" && "$c" -c 'import sys; sys.exit(0 if sys.version_info[:2] >= (3,10) else 1)' 2>/dev/null; then
      PY="$c"; return 0
    fi
  done
  return 1
}

gestor() {
  if hay brew; then echo brew
  elif hay apt-get; then echo apt
  elif hay dnf; then echo dnf
  else echo ninguno; fi
}

instala() {  # instala <paquete-brew> <paquete-apt/dnf> <nombre legible>
  local g; g="$(gestor)"
  case "$g" in
    brew) echo "  instalando $3 ..."; brew install "$1" >/dev/null 2>&1 || FALLOS+=("$3") ;;
    apt)  echo "  instalando $3 ..."; sudo apt-get update -qq >/dev/null 2>&1
          sudo apt-get install -y -qq "$2" >/dev/null 2>&1 || FALLOS+=("$3") ;;
    dnf)  echo "  instalando $3 ..."; sudo dnf install -y -q "$2" >/dev/null 2>&1 || FALLOS+=("$3") ;;
    *)    FALLOS+=("$3 (no hay brew/apt/dnf: instálalo a mano)") ;;
  esac
}

echo ""
echo "== Dependencias del sistema =="

if py_ok; then
  echo "  ok  $($PY -c 'import sys; print("Python %d.%d.%d" % sys.version_info[:3])')"
else
  instala python@3.12 python3 "Python 3.12"
  py_ok || FALLOS+=("Python 3.10+ (abre una terminal nueva y repite)")
fi

if hay ffmpeg; then echo "  ok  ffmpeg"; else instala ffmpeg ffmpeg "ffmpeg"; fi

if [ "$SIN_LIBREOFFICE" = "1" ]; then
  echo "  -   LibreOffice omitido (SIN_LIBREOFFICE=1)"
elif hay soffice || hay libreoffice || [ -d "/Applications/LibreOffice.app" ]; then
  echo "  ok  backend para pasar .pptx/.docx a PDF"
else
  if [ "$(gestor)" = "brew" ]; then
    echo "  instalando LibreOffice ..."
    brew install --cask libreoffice >/dev/null 2>&1 || FALLOS+=("LibreOffice")
  else
    instala libreoffice libreoffice "LibreOffice"
  fi
fi

echo ""
echo "== Dependencias de Python =="
if py_ok; then
  "$PY" -m pip install --quiet --upgrade pip 2>/dev/null
  if "$PY" -m pip install --quiet -r requirements.txt; then
    echo "  ok  openpyxl, pypdf, pymupdf"
  else
    FALLOS+=("pip install -r requirements.txt")
  fi
else
  echo "  ! Python todavía no responde en esta terminal."
fi

echo ""
echo "== Verificación real (preflight) =="
PRE=1
if py_ok; then "$PY" scripts/preflight.py; PRE=$?; fi

if [ ${#FALLOS[@]} -gt 0 ]; then
  echo ""
  echo "FALTA POR RESOLVER:"
  for f in "${FALLOS[@]}"; do echo "  - $f"; done
fi

if [ "$PRE" -eq 0 ] && [ ${#FALLOS[@]} -eq 0 ]; then
  echo ""
  echo "Listo. La máquina puede correr el skill."
  exit 0
fi
echo ""
echo "Quedan cosas pendientes: lee arriba y vuelve a correr este script."
exit 1
