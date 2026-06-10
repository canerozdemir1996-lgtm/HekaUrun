#!/usr/bin/env bash
set -euo pipefail

APP_NAME="UrunYonetimMasasi_v3"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

CLEAN=0
WITH_LOCAL_SETTINGS=0
WITH_INDEX=0

for arg in "$@"; do
  case "$arg" in
    --clean)
      CLEAN=1
      ;;
    --with-local-settings)
      WITH_LOCAL_SETTINGS=1
      ;;
    --with-index)
      WITH_INDEX=1
      ;;
    -h|--help)
      cat <<'EOF'
Mac build:
  ./build_mac.sh --clean

Options:
  --clean                 Remove build/dist/spec before packaging.
  --with-local-settings   Copy local settings.json into the app bundle.
  --with-index            Copy local product_index.sqlite into the app bundle.

Output:
  dist/UrunYonetimMasasi_v3.app
EOF
      exit 0
      ;;
    *)
      echo "Unknown option: $arg" >&2
      exit 1
      ;;
  esac
done

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "This script must run on macOS. PyInstaller cannot create a macOS .app from Windows/Linux." >&2
  exit 1
fi

if [[ "$CLEAN" == "1" ]]; then
  rm -rf build dist "${APP_NAME}.spec"
fi

if [[ ! -d ".venv-mac" ]]; then
  python3 -m venv .venv-mac
fi

# shellcheck disable=SC1091
source ".venv-mac/bin/activate"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

ICON_ARG=()
if [[ -f "assets/app_icon.icns" ]]; then
  ICON_ARG=(--icon "assets/app_icon.icns")
elif [[ -f "assets/app_icon.png" ]]; then
  ICONSET="build/app_icon.iconset"
  rm -rf "$ICONSET"
  mkdir -p "$ICONSET"
  for size in 16 32 128 256 512; do
    sips -z "$size" "$size" "assets/app_icon.png" --out "$ICONSET/icon_${size}x${size}.png" >/dev/null
    double=$((size * 2))
    sips -z "$double" "$double" "assets/app_icon.png" --out "$ICONSET/icon_${size}x${size}@2x.png" >/dev/null
  done
  iconutil -c icns "$ICONSET" -o "assets/app_icon.icns"
  ICON_ARG=(--icon "assets/app_icon.icns")
fi

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  "${ICON_ARG[@]}" \
  "modern_app.py"

APP_BUNDLE="dist/${APP_NAME}.app"
MACOS_DIR="${APP_BUNDLE}/Contents/MacOS"

if [[ ! -d "$MACOS_DIR" ]]; then
  echo "Build failed: ${APP_BUNDLE} was not created." >&2
  exit 1
fi

rm -rf "${MACOS_DIR}/assets" "${MACOS_DIR}/shaders"
cp -R "assets" "${MACOS_DIR}/assets"
if [[ -d "shaders" ]]; then
  cp -R "shaders" "${MACOS_DIR}/shaders"
fi

if [[ "$WITH_LOCAL_SETTINGS" == "1" && -f "settings.json" ]]; then
  cp "settings.json" "${MACOS_DIR}/settings.json"
fi
if [[ "$WITH_INDEX" == "1" && -f "product_index.sqlite" ]]; then
  cp "product_index.sqlite" "${MACOS_DIR}/product_index.sqlite"
fi
if [[ -f "rename_log.jsonl" && "$WITH_LOCAL_SETTINGS" == "1" ]]; then
  cp "rename_log.jsonl" "${MACOS_DIR}/rename_log.jsonl"
fi

cat > "dist/BURADAKI_MAC_APP_CALISTIR.txt" <<EOF
Programi buradan calistir:
${APP_NAME}.app

Mac'e tasirken dist klasorundeki ${APP_NAME}.app paketini kopyala.
Varsayilan build yerel settings.json ve product_index.sqlite dosyalarini pakete koymaz.
Ayni ag/disk yollarini kullanan bir Mac icin gerekirse:
  ./build_mac.sh --clean --with-local-settings --with-index

Gatekeeper uygulamayi engellerse:
  Uygulamaya sag tikla > Open
veya Terminal:
  xattr -dr com.apple.quarantine "dist/${APP_NAME}.app"

SQL Server ODBC kullanilacaksa Mac'e Microsoft ODBC Driver for SQL Server kurulmalidir.
EOF

echo "Mac app ready: ${SCRIPT_DIR}/${APP_BUNDLE}"
