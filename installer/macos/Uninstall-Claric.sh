#!/bin/bash
#
# uninstall-claric.sh -- remove the Claric Word add-in registration installed
# by Install-Claric.sh.
#
# Deletes the developer-add-in manifest copies from Word's sideload folder
# (~/Library/Containers/com.microsoft.Word/Data/Documents/wef): any copy
# whose <Id> matches the installed manifest, plus any other Claric-published
# manifest (matched by <ProviderName>Claric</ProviderName>), so it also
# cleans up self-hosted manifests installed with --manifest. Then removes
# ~/Library/Application Support/ClaricAddin unless --keep-files is given.
#
# Usage:
#   ./Uninstall-Claric.sh
#   ./Uninstall-Claric.sh --manifest /path/to/manifest.xml --keep-files
set -euo pipefail

MANIFEST_PATH=''
KEEP_FILES=0

usage() {
  cat <<'EOF'
uninstall-claric.sh -- remove the Claric Word add-in registration installed
by Install-Claric.sh.

Usage:
  Uninstall-Claric.sh               # remove registration + ~/Library/Application Support/ClaricAddin
  Uninstall-Claric.sh --keep-files  # remove registration only

Options:
  --manifest <path>  Unregister this manifest's add-in id
  --keep-files       Keep manifest copy and launch document on disk
  -h, --help         Show this help
EOF
}

fail() {
  echo "[uninstall] error  : $*" >&2
  exit 1
}

while [ $# -gt 0 ]; do
  case "$1" in
    --manifest)  [ $# -ge 2 ] || fail "--manifest needs a path"; MANIFEST_PATH="$2"; shift 2 ;;
    --keep-files) KEEP_FILES=1; shift ;;
    -h|--help)   usage; exit 0 ;;
    *)           fail "Unknown option: $1 (see --help)" ;;
  esac
done

wef_dir="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
install_dir="$HOME/Library/Application Support/ClaricAddin"

if [ -z "$MANIFEST_PATH" ] && [ -f "$install_dir/manifest.xml" ]; then
  MANIFEST_PATH="$install_dir/manifest.xml"
fi

GUID_RE='[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
addin_id=''
if [ -n "$MANIFEST_PATH" ] && [ -f "$MANIFEST_PATH" ]; then
  addin_id="$(grep -Eo "<Id>[[:space:]]*${GUID_RE}[[:space:]]*</Id>" "$MANIFEST_PATH" 2>/dev/null \
    | head -n1 | sed -E 's|</?Id>||g' | tr -d '[:space:]' || true)"
fi

removed=0
if [ -d "$wef_dir" ]; then
  for entry in "$wef_dir"/*.xml; do
    [ -f "$entry" ] || continue
    name="$(basename "$entry")"
    id="$(grep -Eo "<Id>[[:space:]]*${GUID_RE}[[:space:]]*</Id>" "$entry" 2>/dev/null \
      | head -n1 | sed -E 's|</?Id>||g' | tr -d '[:space:]' || true)"
    if { [ -n "$addin_id" ] && [ "$id" = "$addin_id" ]; } \
        || { [ -n "$addin_id" ] && [ "${name#"${addin_id}".}" != "$name" ]; } \
        || grep -q '<ProviderName>Claric</ProviderName>' "$entry" 2>/dev/null; then
      rm -f "$entry"
      echo "[uninstall] wef     : removed $name"
      removed=$((removed + 1))
    fi
  done
  if [ "$removed" -eq 0 ]; then
    echo '[uninstall] wef     : no Claric manifests found -- nothing to do'
  fi
else
  echo '[uninstall] wef     : sideload folder absent -- nothing to do'
fi

if [ "$KEEP_FILES" -eq 1 ]; then
  echo "[uninstall] files   : kept ($install_dir)"
elif [ -d "$install_dir" ]; then
  rm -rf "$install_dir"
  echo "[uninstall] files   : removed $install_dir"
fi

echo 'Done. Restart Word if it is running so the change takes effect.'
