#requires -Version 5.1
<#
.SYNOPSIS
  One-click sideload of the Claric Word add-in for Word on Windows.

.DESCRIPTION
  Registers a Claric manifest.xml (the static GitHub Pages build by default,
  or a self-hosted one via -ManifestPath) as a Word "developer add-in":

    1. Copies the manifest to a stable per-user location:
       %LOCALAPPDATA%\ClaricAddin\manifest.xml
    2. Writes HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\<AddInId>
       (string) = that path -- the same registry mechanism the official
       office-addin-dev-settings tool uses for "npm start" sideloads.
    3. Builds Claric-Launch.docx from the bundled Microsoft template whose
       webextension references the add-in from the "developer" (Registry)
       store, then opens it -- Word resolves the reference and mounts the
       taskpane.

  Why not "Insert > Get Add-ins > Upload My Add-in"? That entry is gone from
  current consumer Microsoft 365 builds (the "Manage My Add-ins" link only
  opens the web portal), and the trusted-catalog route requires a real UNC
  network share. The developer registry works everywhere, per user, and
  needs no admin rights.

  Idempotent: re-running refreshes the manifest, registry value, and launch
  document in place.

.PARAMETER ManifestPath
  Local manifest.xml to install (e.g. one generated for your own server).
  Default: download the static build from GitHub Pages.

.PARAMETER ManifestUrl
  Where to download the manifest when -ManifestPath is not given.

.PARAMETER TemplatePath
  WordDocumentWithTaskPane.docx used to build the launch document.
  Default: templates\ next to this script, downloaded from the repository
  when the script is run standalone.

.PARAMETER NoLaunch
  Register everything but do not open the launch document in Word.

.EXAMPLE
  .\Install-Claric.ps1
  .\Install-Claric.ps1 -ManifestPath ..\..\manifest.xml
#>
[CmdletBinding()]
param(
  [string]$ManifestPath = '',
  [string]$ManifestUrl = 'https://scilence2022.github.io/claric-addin/manifest.xml',
  [string]$TemplatePath = '',
  [string]$TemplateUrl = 'https://raw.githubusercontent.com/Scilence2022/Claric/main/installer/windows/templates/WordDocumentWithTaskPane.docx',
  [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

$installDir = Join-Path $env:LOCALAPPDATA 'ClaricAddin'
$manifestDest = Join-Path $installDir 'manifest.xml'
$launchDoc = Join-Path $installDir 'Claric-Launch.docx'

function Read-ManifestInfo {
  param([Parameter(Mandatory)] [string]$Path)
  $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
  $id = if ($raw -match '<Id>\s*([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s*</Id>') { $Matches[1] } else { $null }
  $ver = if ($raw -match '<Version>\s*(\d+(?:\.\d+){0,3})\s*</Version>') { $Matches[1] } else { $null }
  if (-not $id) { throw "No <Id> GUID found in $Path -- not a valid Office add-in manifest." }
  if (-not $ver) { throw "No <Version> found in $Path." }
  [pscustomobject]@{ Id = $id; Version = $ver }
}

# --- 1. manifest ------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
if ($ManifestPath) {
  if (-not (Test-Path -LiteralPath $ManifestPath)) { throw "Manifest not found: $ManifestPath" }
  Copy-Item -LiteralPath $ManifestPath $manifestDest -Force
  Write-Host "[install] manifest  : $ManifestPath"
} else {
  Write-Host "[install] manifest  : downloading $ManifestUrl"
  Invoke-WebRequest -Uri $ManifestUrl -OutFile $manifestDest -UseBasicParsing
}
$info = Read-ManifestInfo -Path $manifestDest
Write-Host ("[install] add-in id : {0}  (version {1})" -f $info.Id, $info.Version)

# --- 2. registry (Word developer add-in) -------------------------------------
$devKey = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
New-Item -Path $devKey -Force | Out-Null
New-ItemProperty -Path $devKey -Name $info.Id -Value $manifestDest -PropertyType String -Force | Out-Null
Write-Host "[install] registry  : $devKey\<$($info.Id)> = $manifestDest"

# --- 3. launch document ------------------------------------------------------
if (-not $TemplatePath -and $PSScriptRoot) {
  $local = Join-Path $PSScriptRoot 'templates\WordDocumentWithTaskPane.docx'
  if (Test-Path -LiteralPath $local) { $TemplatePath = $local }
}
if (-not $TemplatePath) {
  # Standalone run (e.g. `irm ... | iex`): fetch the template next to the script source.
  Write-Host "[install] template  : not found locally, downloading from repository"
  $tmp = Join-Path $env:TEMP 'Claric-WordDocumentWithTaskPane.docx'
  Invoke-WebRequest -Uri $TemplateUrl -OutFile $tmp -UseBasicParsing
  $TemplatePath = $tmp
}

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
# Build on a temp copy first so a launch document held open by Word cannot
# leave a half-written file behind.
$staging = Join-Path $env:TEMP ('Claric-Launch-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.docx')
Copy-Item -LiteralPath $TemplatePath $staging -Force
$webExtXml = '<?xml version="1.0" encoding="utf-8"?><we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{' + $info.Id + '}"><we:reference id="' + $info.Id + '" version="' + $info.Version + '" store="developer" storeType="Registry" /><we:alternateReferences /><we:properties></we:properties><we:bindings /></we:webextension>'
$zip = [System.IO.Compression.ZipFile]::Open($staging, 'Update')
try {
  $entry = $zip.GetEntry('word/webextensions/webextension.xml')
  if ($null -eq $entry) { throw 'word/webextensions/webextension.xml missing from template docx' }
  $entry.Delete()
  $new = $zip.CreateEntry('word/webextensions/webextension.xml')
  $writer = New-Object System.IO.StreamWriter($new.Open(), (New-Object System.Text.UTF8Encoding($false)))
  try { $writer.Write($webExtXml) } finally { $writer.Dispose() }
} finally { $zip.Dispose() }

$launchDocFinal = $launchDoc
try {
  [System.IO.File]::Copy($staging, $launchDocFinal, $true)
  Remove-Item -LiteralPath $staging -Force
} catch [System.IO.IOException] {
  # Word is holding Claric-Launch.docx open -- park the fresh copy next to it.
  $launchDocFinal = Join-Path $installDir ('Claric-Launch-{0:yyyyMMdd-HHmmss}.docx' -f (Get-Date))
  Move-Item -LiteralPath $staging -Destination $launchDocFinal
  Write-Host "[install] note      : $launchDoc is open in Word -- created $launchDocFinal instead"
}
Write-Host "[install] launch doc: $launchDocFinal"

# --- 4. done ------------------------------------------------------------------
Write-Host ''
Write-Host 'Claric is installed for the current user.'
if ($NoLaunch) {
  Write-Host "Open $launchDocFinal in Word whenever you need the taskpane."
} else {
  Write-Host 'Opening Word with the Claric taskpane...'
  Start-Process -FilePath $launchDocFinal
  Write-Host "If the taskpane does not appear, close Word fully and reopen $launchDocFinal."
}
