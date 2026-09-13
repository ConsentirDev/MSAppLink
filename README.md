# SharePoint Office Launcher for Chrome

Small Manifest V3 enterprise extension that converts approved SharePoint Server document navigations into Microsoft Office desktop URI launches (`ms-word:`, `ms-excel:`, `ms-powerpoint:`, `ms-visio:`).

## Why this exists

The target migration removes Office Online Server/WOPI. SharePoint Server can still be configured to prefer client applications, but Chrome may download a document when a user follows a native SharePoint shared link. This extension provides a narrowly scoped, reversible client-side compatibility shim for managed Chrome devices.

## Safety model

- HTTPS only.
- Exact allow-list of SharePoint origins compiled into the extension manifest.
- Optional path-prefix allow-list at runtime.
- Office document types only; PDF, ZIP, images and executables are never intercepted.
- Never guesses a document location from a filename alone.
- Unknown/unresolvable SharePoint links fail open to native SharePoint/Chrome behavior.
- No telemetry, remote code, credentials or document content collection.

## Build

1. Copy `config/targets.example.json` to `config/targets.json`.
2. Replace the example origin and path prefix with the actual SharePoint SE Workspace URL scope.
3. Run:

```bash
npm test
npm run build:dev
```

Load `dist/` as an unpacked extension for the first POC, then open the extension **Options** page and enter the real SharePoint origin/path. The dev build intentionally has broad HTTPS host permission so the target can be changed without rebuilding; runtime logic still limits action to the configured origin. It uses `webNavigation` and can race Chrome's normal navigation, so it is only for functional validation.

For the managed deployment build, put the real origin/path into `config/targets.json` first. The enterprise manifest is then compiled with only those exact host permissions:

```bash
npm run build
```

The enterprise build also requests `webRequest` + `webRequestBlocking`. Chrome retains `webRequestBlocking` for policy-installed Manifest V3 extensions. The extension cancels a recognised top-level document request before asking Chrome to launch the Office protocol.

## What it recognises

- Direct Office document URLs under an approved SharePoint origin/path.
- SharePoint download links containing a concrete `SourceUrl`, `FileRef`, `RootFolder`, `id`, or similar same-origin path.
- `Doc.aspx?sourcedoc={GUID}&file=...` native share links are handled separately by a same-origin resolver script. It attempts `/_api/web/GetFileById(...)` and only redirects if SharePoint returns a concrete `ServerRelativeUrl`.

The `GetFileById` REST behavior must be proven against the actual SharePoint 2019/SE farm. If unsupported, the resolver deliberately does nothing and Chrome/SharePoint continues normally. Direct-file and `SourceUrl` handling remain available.

## Office URI mapping

- Word: `ms-word:ofe|u|https://.../file.docx`
- Excel: `ms-excel:ofe|u|https://.../file.xlsx`
- PowerPoint: `ms-powerpoint:ofe|u|https://.../file.pptx`
- Visio: `ms-visio:ofe|u|https://.../file.vsdx`

Set `openMode` to `view` to use `ofv` instead of `ofe`.

## Chrome policy

`policy/chrome-policy.sample.json` contains examples for force installation and `AutoLaunchProtocolsFromOrigins`. Replace placeholders with the final extension ID/update URL and the real SharePoint origin.

`policy/managed-extension-policy.sample.json` shows the intended extension-managed settings. The exact Intune/Chrome ADMX delivery shape should be aligned with how the environment currently deploys Chrome third-party extension policy.

## Initial acceptance tests

1. Direct `.docx`, `.xlsx` and `.pptx` links from Chrome open the desktop application.
2. Native SharePoint “Share” link copied into Chrome opens the desktop app if the farm exposes a resolvable concrete URL.
3. The same share link from Outlook desktop with Chrome as default browser opens the desktop app.
4. Links outside the configured SharePoint origin/path are untouched.
5. PDF/ZIP/image links are untouched.
6. A document with no matching desktop Office protocol fails visibly rather than looping.
7. Chrome policy removes or minimizes the external-protocol confirmation prompt on managed devices.
8. Disabling the managed `enabled` value immediately returns behavior to native Chrome/SharePoint.

## POC limitation to prove first

The most environment-specific path is native `Doc.aspx?sourcedoc={GUID}` resolution on SharePoint Server. Microsoft documents Office URI schemes, but documentation around `GetFileById` in on-prem CSOM/REST is inconsistent. Treat that resolver as a testable optimization, not a production assumption.
