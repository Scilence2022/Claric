# macOS one-click installer

Installs the Claric Word add-in for the **current user** of a Mac -- no admin
rights, no Node.js, no Docker, no Homebrew. It uses Word's developer sideload
folder, which is available on every Word for Mac build (Microsoft's own
`office-addin-dev-settings` tool registers add-ins the same way on macOS).

## Install

### Static build (GitHub Pages) -- recommended for everyday use

Run in Terminal:

```bash
curl -fsSL https://raw.githubusercontent.com/Scilence2022/Claric/main/installer/macos/Install-Claric.sh | bash
```

If `raw.githubusercontent.com` is unreliable from your network, download this
folder and run:

```bash
bash Install-Claric.sh
```

Word opens `Claric-Launch.docx` with the Claric taskpane mounted. Later,
whenever you need the taskpane, reopen that document (it stays in
`~/Library/Application Support/ClaricAddin/`), or find Claric under
**Home ? Add-ins ? Claric** (also listed under **Insert ? Add-ins**).

> The first time an add-in's code domain loads, Word may ask you to trust its
> HTTPS certificate -- accept once and the pane mounts.

### Self-hosted build (Docker / npm)

After starting your server and generating its manifest, point the installer
at your own `manifest.xml`:

```bash
./Install-Claric.sh --manifest /path/to/Claric/manifest.xml
```

or, from a repository checkout:

```
npm run install:macos
```

(`npm run install:macos` always registers the `manifest.xml` at the repo
root -- switch it to your server first with `npm run manifest:local`. It
registers without launching Word, like `npm run install:windows` does.)

### Options

| Option | Meaning |
|--------|---------|
| `--manifest <path>` | Install this local manifest instead of downloading the static build's |
| `--manifest-url <url>` | Download the manifest from this URL instead |
| `--template <path>` | Use this docx as the launch-document template (default: the bundled Microsoft template shared with the Windows installer, downloaded from the repository when the script runs standalone) |
| `--no-launch` | Register everything but do not open the launch document in Word |

## Uninstall

```bash
./Uninstall-Claric.sh               # wef registration + ~/Library/Application Support/ClaricAddin
./Uninstall-Claric.sh --keep-files  # wef registration only
```

or `npm run uninstall:macos` from a checkout. Restart Word afterwards.

## How it works

Word for Mac has no per-user registry; its developer-add-in catalog is a
folder inside Word's sandbox container. The docx carries only a *reference*
to the add-in, so the installer sets up the two halves Word needs:

1. **Manifest** -- copied to `~/Library/Application Support/ClaricAddin/manifest.xml`
   (the stable per-user copy) and into
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/<Id>.manifest.xml`.
   Word reads the `wef` folder at startup and offers everything it finds under
   the developer add-ins -- the same registration `office-addin-dev-settings`
   performs for `npm start` sideloads on macOS.
2. **Launch document** -- the same `WordDocumentWithTaskPane.docx` template the
   Windows installer uses (bundled from that Microsoft package, MIT) gets its
   `word/webextensions/webextension.xml` rewritten to reference the add-in
   id/version from the developer store, saved as `Claric-Launch.docx`.

Opening that document makes Word resolve the reference through the `wef`
registration and mount the taskpane. The registration is per-user and
persists, so the document works on every subsequent launch; on a *different*
Mac you must run the installer again (a docx alone does not carry the add-in).

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Taskpane does not appear when the launch document opens | Fully quit Word (Cmd+Q) and reopen `Claric-Launch.docx` -- Word scans the `wef` folder at startup. |
| "This add-in can't run..." / blank pane on a self-hosted build | The server certificate isn't trusted -- see the main README section *Trust the Certificate* for macOS (Keychain import). |
| Re-installed a manifest with a different GUID (e.g. switched static ? local-server route) | Both registrations can coexist; `Uninstall-Claric.sh` removes every Claric-published manifest from the `wef` folder. Re-run the installer afterwards. |
| `curl \| bash` fails behind a blocked network | Download the folder instead (see above); the script only needs outbound HTTPS to `scilence2022.github.io` (manifest) and `raw.githubusercontent.com` (template, when run standalone). |
| Prefer doing it by hand | Word for Mac scans `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` at startup -- create the folder if missing, drop the manifest in, then restart Word per [Microsoft's Mac sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac). |
