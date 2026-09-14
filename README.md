# SharePoint Office Launcher for Chrome

Small Manifest V3 enterprise extension that converts approved SharePoint Server document navigations into Microsoft Office desktop URI launches (`ms-word:`, `ms-excel:`, `ms-powerpoint:`, `ms-visio:`).

## Why this exists

The target migration removes Office Online Server/WOPI. SharePoint Server can still be configured to prefer client applications, but Chrome may download a document when a user follows a native SharePoint shared link. This extension provides a narrowly scoped, reversible client-side compatibility shim for managed Chrome devices.

## Safety model

- HTTPS only.
- Exact allow-list of SharePoint origins compiled into the enterprise extension manifest.
- Optional path-prefix allow-list at runtime.
- Office document types only; PDF, ZIP, images and executables are never intercepted.
- Never guesses a document location from a filename alone.
- Unknown/unresolvable SharePoint links fail open to native SharePoint/Chrome behavior.
- No telemetry, remote code, credentials or document content collection.

## Build

Install Node.js 20+ and run:

```bash
npm test
```

### Local POC / unpacked Chrome extension

For developer-mode testing use **only**:

```bash
npm run build:dev
```

Load `dist-dev/` from `chrome://extensions` using **Load unpacked**.

The development build intentionally does **not** request `webRequestBlocking`. In Manifest V3 that permission is available only to extensions installed by enterprise policy, so Chrome will reject an unpacked MV3 extension that requests it. The dev build therefore uses the `webNavigation` fallback for functional validation.

### Enterprise/policy build

For managed deployment:

```bash
npm run build:enterprise
```

This creates `dist-enterprise/` with host permissions restricted to the configured GRDC Workspaces origins. It requests `webRequest` + `webRequestBlocking` so recognised top-level document requests can be cancelled before Office is launched.

**Do not load `dist-enterprise/` with Chrome's Load unpacked button.** `webRequestBlocking` in Manifest V3 is valid only when the extension is installed by enterprise policy/force-installation.

The current configured origins are:

- `https://workspaces.grdc.com.au`
- `https://workspaces.test.grdc.com.au`
- `https://workspaces.dev.grdc.com.au`

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

`policy/chrome-policy.sample.json` contains examples for force installation and `AutoLaunchProtocolsFromOrigins`. Replace placeholders with the final extension ID/update URL.

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
