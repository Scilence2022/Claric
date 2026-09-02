# Windows one-click installer

Installs the Claric Word add-in for the **current user** of a Windows PC ?
no admin rights, no Node.js, no Docker. It works on consumer Microsoft 365
builds where the classic "Upload My Add-in" entry is no longer offered.

## Install

### Static build (GitHub Pages) ? recommended for everyday use

Run in PowerShell:

```powershell
irm https://raw.githubusercontent.com/Scilence2022/Claric/main/installer/windows/Install-Claric.ps1 | iex
```

If `raw.githubusercontent.com` is unreliable from your network, download this
folder (`Install-Claric.ps1` + `Uninstall-Claric.ps1` + `templates/`) and run:

```powershell
.\Install-Claric.ps1
```

Word opens `Claric-Launch.docx` with the Claric taskpane mounted. Later,
whenever you need the taskpane, reopen that document (it stays in
`%LOCALAPPDATA%\ClaricAddin\`), or find Claric under
**Insert ? Get Add-ins ? My Add-ins**.

### Self-hosted build (Docker / npm)

After starting your server and generating its manifest, point the installer
at your own `manifest.xml`:

```powershell
.\Install-Claric.ps1 -ManifestPath C:\path\to\Claric\manifest.xml
```

or, from a repository checkout:

```
npm run install:windows
```

(`npm run install:windows` always registers the `manifest.xml` at the repo
root ? switch it to your server first with `npm run manifest:local`.)

## Uninstall

```powershell
.\Uninstall-Claric.ps1            # registry + %LOCALAPPDATA%\ClaricAddin
.\Uninstall-Claric.ps1 -KeepFiles # registry only
```

or `npm run uninstall:windows` from a checkout. Restart Word afterwards.

## How it works

The docx carries only a *reference* to the add-in, so the installer sets up
the two halves Word needs on the machine:

1. **Manifest** ? copied to `%LOCALAPPDATA%\ClaricAddin\manifest.xml`.
2. **Registration** ? a string value named after the manifest `<Id>` under
   `HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer` pointing at that
   file. This is the same mechanism Microsoft's `office-addin-dev-settings`
   uses for `npm start` sideloads.
3. **Launch document** ? `templates/WordDocumentWithTaskPane.docx` (bundled
   from that same Microsoft package, MIT) gets its
   `word/webextensions/webextension.xml` rewritten to reference the add-in
   id/version from the developer (Registry) store, saved as
   `Claric-Launch.docx`.

Opening that document makes Word resolve the reference through the registry
and mount the taskpane. The registration is per-user and persists, so the
document works on every subsequent launch; on a *different* PC you must run
the installer again (a docx alone does not carry the add-in).

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Taskpane does not appear when the launch document opens | Close Word completely and reopen `Claric-Launch.docx`. |
| "This add-in can't run..." / blank pane on a self-hosted build | The server certificate isn't trusted ? see the README section *Trust the Certificate on Windows*. |
| Re-installed a manifest with a different GUID (e.g. switched static ? local-server route) | Both registrations can coexist; `Uninstall-Claric.ps1` removes every registration pointing into `%LOCALAPPDATA%\ClaricAddin`. Re-run the installer afterwards. |
| `irm \| iex` fails behind a blocked network | Download the folder instead (see above); the script only needs outbound HTTPS to `scilence2022.github.io` (manifest) and `raw.githubusercontent.com` (template, when run standalone). |
| PowerShell execution policy blocks local scripts | Run with `powershell -ExecutionPolicy Bypass -File .\Install-Claric.ps1`. |
