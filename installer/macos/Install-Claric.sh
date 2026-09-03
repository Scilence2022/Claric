#!/bin/bash
#
# install-claric.sh -- one-click sideload of the Claric Word add-in for
# Word on macOS.
#
# Registers a Claric manifest.xml (the static GitHub Pages build by default,
# or a self-hosted one via --manifest) as a Word "developer add-in":
#
#   1. Copies the manifest to a stable per-user location:
#        ~/Library/Application Support/ClaricAddin/manifest.xml
#   2. Copies it into Word's developer sideload folder:
#        ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/<Id>.manifest.xml
#      Word reads that folder at startup -- the same registration mechanism
#      Microsoft's office-addin-dev-settings tool uses on the Mac (there is
#      no per-user registry to write to here; this folder IS the developer
#      catalog).
#   3. Builds Claric-Launch.docx from the bundled Microsoft template whose
#      webextension references the add-in from the developer store, then
#      opens it -- Word resolves the reference and mounts the taskpane.
#
# Idempotent: re-running refreshes the manifest, registration, and launch
# document in place. No admin rights, no Node.js, no Homebrew -- only stock
# macOS tools (curl, zip, unzip).
#
# Usage:
#   ./Install-Claric.sh                          # download the static build manifest
#   ./Install-Claric.sh --manifest manifest.xml  # install a self-hosted manifest
#   curl -fsSL <raw-url>/Install-Claric.sh | bash
set -euo pipefail

MANIFEST_PATH=''
MANIFEST_URL='https://scilence2022.github.io/claric-addin/manifest.xml'
TEMPLATE_PATH=''
TEMPLATE_URL='https://raw.githubusercontent.com/Scilence2022/Claric/main/installer/windows/templates/WordDocumentWithTaskPane.docx'
NO_LAUNCH=0

usage() {
  cat <<'EOF'
install-claric.sh -- one-click sideload of the Claric Word add-in for
Word on macOS.

Usage:
  Install-Claric.sh                          # download the static build manifest
  Install-Claric.sh --manifest manifest.xml  # install a self-hosted manifest
  curl -fsSL <raw-url>/Install-Claric.sh | bash

Options:
  --manifest <path>     Install this local manifest.xml instead of downloading
  --manifest-url <url>  Download the manifest from this URL instead
  --template <path>     Use this docx as the launch-document template
  --no-launch           Register everything, but do not open Word
  -h, --help            Show this help
EOF
}

fail() {
  echo "[install] error    : $*" >&2
  exit 1
}

while [ $# -gt 0 ]; do
  case "$1" in
    --manifest)     [ $# -ge 2 ] || fail "--manifest needs a path"; MANIFEST_PATH="$2"; shift 2 ;;
    --manifest-url) [ $# -ge 2 ] || fail "--manifest-url needs a URL"; MANIFEST_URL="$2"; shift 2 ;;
    --template)     [ $# -ge 2 ] || fail "--template needs a path"; TEMPLATE_PATH="$2"; shift 2 ;;
    --no-launch)    NO_LAUNCH=1; shift ;;
    -h|--help)      usage; exit 0 ;;
    *)              fail "Unknown option: $1 (see --help)" ;;
  esac
done

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || true)"

word_app="/Applications/Microsoft Word.app"
wef_dir="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
install_dir="$HOME/Library/Application Support/ClaricAddin"
manifest_dest="$install_dir/manifest.xml"
launch_doc="$install_dir/Claric-Launch.docx"

if [ ! -d "$word_app" ] && [ ! -d "$HOME/Library/Containers/com.microsoft.Word" ]; then
  echo "[install] warning  : Microsoft Word not found on this Mac -- registering anyway"
fi

work="$(mktemp -d)"
cleanup() { rm -rf "$work" 2>/dev/null || true; }
trap cleanup EXIT

# --- 1. manifest -------------------------------------------------------------
mkdir -p "$install_dir"
if [ -n "$MANIFEST_PATH" ]; then
  if [ ! -f "$MANIFEST_PATH" ]; then fail "Manifest not found: $MANIFEST_PATH"; fi
  cp -f "$MANIFEST_PATH" "$manifest_dest"
  echo "[install] manifest  : $MANIFEST_PATH"
else
  echo "[install] manifest  : downloading $MANIFEST_URL"
  curl -fsSL "$MANIFEST_URL" -o "$manifest_dest" \
    || fail "Could not download $MANIFEST_URL -- pass a local file with --manifest"
fi

# GUID that names the add-in; Word keys everything on it.
GUID_RE='[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
read_manifest_id() {
  grep -Eo "<Id>[[:space:]]*${GUID_RE}[[:space:]]*</Id>" "$1" 2>/dev/null \
    | head -n1 | sed -E 's|</?Id>||g' | tr -d '[:space:]' || true
}

addin_id="$(read_manifest_id "$manifest_dest")"
[ -n "$addin_id" ] || fail "No <Id> GUID found in $manifest_dest -- not a valid Office add-in manifest."
addin_version="$(grep -Eo '<Version>[[:space:]]*[0-9]+(\.[0-9]+){0,3}[[:space:]]*</Version>' "$manifest_dest" \
  | head -n1 | sed -E 's|</?Version>||g' | tr -d '[:space:]')"
[ -n "$addin_version" ] || fail "No <Version> found in $manifest_dest."
echo "[install] add-in id : $addin_id  (version $addin_version)"

# --- 2. registration (Word developer sideload folder) -------------------------
mkdir -p "$wef_dir"
# Remove prior copies of this same add-in so re-runs (or a switch between the
# static and a self-hosted manifest) never leave a stale ghost behind.
for prior in "$wef_dir"/*.xml; do
  [ -f "$prior" ] || continue
  if [ "$(read_manifest_id "$prior")" = "$addin_id" ]; then
    rm -f "$prior"
    echo "[install] removed prior manifest $(basename "$prior")"
  fi
done
wef_manifest="$wef_dir/$addin_id.manifest.xml"
cp -f "$manifest_dest" "$wef_manifest"
echo "[install] register  : $wef_manifest"

# --- 3. launch document -------------------------------------------------------
template="$TEMPLATE_PATH"
if [ -z "$template" ] && [ -n "$script_dir" ] \
    && [ -f "$script_dir/../windows/templates/WordDocumentWithTaskPane.docx" ]; then
  template="$script_dir/../windows/templates/WordDocumentWithTaskPane.docx"
fi
if [ -z "$template" ]; then
  # Standalone run (e.g. `curl ... | bash`): fetch the template from the repository.
  echo "[install] template  : not found locally, downloading from repository"
  template="$work/WordDocumentWithTaskPane.docx"
  curl -fsSL "$TEMPLATE_URL" -o "$template" \
    || fail "Could not download $TEMPLATE_URL -- pass a local file with --template"
fi

# Build on a staging copy first so a launch document held open by Word cannot
# leave a half-written file behind.
staging="$work/Claric-Launch.docx"
cp -f "$template" "$staging"
pkg_dir="$work/pkg"
xml_file="word/webextensions/webextension.xml"
mkdir -p "$(dirname "$pkg_dir/$xml_file")"
# Same reference Word writes for a developer add-in; Microsoft's own
# sideloading tool generates this identical part for Word on the Mac.
printf '%s' "<?xml version=\"1.0\" encoding=\"utf-8\"?><we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" id=\"{$addin_id}\"><we:reference id=\"$addin_id\" version=\"$addin_version\" store=\"developer\" storeType=\"Registry\" /><we:alternateReferences /><we:properties></we:properties><we:bindings /></we:webextension>" \
  > "$pkg_dir/$xml_file"
(cd "$pkg_dir" && zip -q "$staging" "$xml_file") \
  || fail "Could not update webextension.xml inside the launch document."
unzip -p "$staging" "word/webextensions/webextension.xml" | grep -q "$addin_id" \
  || fail "Launch document verification failed."
mv -f "$staging" "$launch_doc"
echo "[install] launch doc: $launch_doc"

# --- 4. done -------------------------------------------------------------------
echo ''
echo 'Claric is installed for the current user.'
if [ "$NO_LAUNCH" -eq 1 ]; then
  echo "Open $launch_doc in Word whenever you need the taskpane."
else
  if [ ! -d "$word_app" ]; then
    echo "Microsoft Word.app not found -- open $launch_doc with Word manually."
  else
    echo 'Opening Word with the Claric taskpane...'
    if pgrep -xq "Microsoft Word"; then
      echo '[install] note      : Word is running -- fully quit it (Cmd+Q) and reopen'
      echo '            '"$launch_doc"' if the taskpane does not appear.'
    fi
    open -a "Microsoft Word" "$launch_doc"
  fi
fi
