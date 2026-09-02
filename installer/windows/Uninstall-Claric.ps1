#requires -Version 5.1
<#
.SYNOPSIS
  Remove the Claric Word add-in registration installed by Install-Claric.ps1.

.DESCRIPTION
  Deletes the developer-add-in registry value(s) under
  HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer that point at
  %LOCALAPPDATA%\ClaricAddin\manifest.xml (any add-in id, so it also cleans
  up self-hosted manifests installed with -ManifestPath), then removes the
  %LOCALAPPDATA%\ClaricAddin folder unless -KeepFiles is given.

.PARAMETER ManifestPath
  Manifest whose <Id> should be unregistered, in case it was installed from
  elsewhere. Default: the copy inside %LOCALAPPDATA%\ClaricAddin.

.PARAMETER KeepFiles
  Remove only the registry registration; keep manifest and launch document.

.EXAMPLE
  .\Uninstall-Claric.ps1
  .\Uninstall-Claric.ps1 -ManifestPath C:\path\to\manifest.xml -KeepFiles
#>
[CmdletBinding()]
param(
  [string]$ManifestPath = '',
  [switch]$KeepFiles
)

$ErrorActionPreference = 'Stop'
$devKey = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$installDir = Join-Path $env:LOCALAPPDATA 'ClaricAddin'

$ids = @()
if (-not $ManifestPath) {
  $default = Join-Path $installDir 'manifest.xml'
  if (Test-Path -LiteralPath $default) { $ManifestPath = $default }
}
if ($ManifestPath -and (Test-Path -LiteralPath $ManifestPath)) {
  $raw = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8
  if ($raw -match '<Id>\s*([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s*</Id>') {
    $ids += $Matches[1]
  }
}

if (Test-Path $devKey) {
  $removed = 0
  $key = Get-Item $devKey
  foreach ($name in @($key.GetValueNames())) {
    $data = [string]$key.GetValue($name)
    if (($ids -contains $name) -or ($data -like '*\ClaricAddin\manifest.xml')) {
      Remove-ItemProperty -Path $devKey -Name $name
      Write-Host "[uninstall] registry : removed $name ($data)"
      $removed++
    }
  }
  if ($removed -eq 0) { Write-Host '[uninstall] registry : no Claric values found -- nothing to do' }
} else {
  Write-Host '[uninstall] registry : developer key absent -- nothing to do'
}

if (-not $KeepFiles) {
  # Only delete the folder when no other registration still points into it
  # (e.g. a static install alongside a self-hosted one).
  $stillReferenced = $false
  if (Test-Path $devKey) {
    $key = Get-Item $devKey
    foreach ($name in @($key.GetValueNames())) {
      if (([string]$key.GetValue($name)) -like '*\ClaricAddin\manifest.xml') { $stillReferenced = $true }
    }
  }
  if ($stillReferenced) {
    Write-Host "[uninstall] files    : kept -- other registrations still point into $installDir"
  } elseif (Test-Path $installDir) {
    # Best-effort: a launch document held open by Word cannot be deleted;
    # the registry cleanup above already did the functional part.
    Remove-Item -LiteralPath $installDir -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $installDir) {
      Write-Host "[uninstall] files    : some files in $installDir are locked (Word?) -- close Word, then re-run or delete the folder manually"
    } else {
      Write-Host "[uninstall] files    : removed $installDir"
    }
  }
} else {
  Write-Host "[uninstall] files    : kept ($installDir)"
}

Write-Host 'Done. Restart Word if it is running so the change takes effect.'
