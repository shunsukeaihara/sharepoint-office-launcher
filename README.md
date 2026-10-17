# SharePoint Office Launcher (Chrome / Edge extension)

English | [日本語](README_jp.md)

A Manifest V3 extension that opens SharePoint PowerPoint, Excel, and Word links directly in the **Office desktop apps** instead of the browser.

When a link like `https://contoso.sharepoint.com/:p:/r/sites/Sales/Shared%20Documents/Proposal.pptx?d=…&web=1` is opened in Chrome or Edge (for example, from Slack), the extension converts it to the actual file URL and launches the app with the [Office URI scheme](https://learn.microsoft.com/office/client-developer/office-uri-schemes) (`ms-powerpoint:` / `ms-excel:` / `ms-word:`).

- Acts **only on tenants that the user (or an administrator) has allowed**. The allowlist is empty by default; the settings page opens right after installation.
- Any link that is not allowed or cannot be converted is left alone and opens in the browser as usual.
- The UI is available in English (default) and Japanese, following the browser language.

- Privacy policy: [PRIVACY.md](PRIVACY.md)
- License: [MIT](LICENSE)

## Supported links

| Link | Office URI |
|---|---|
| `https://<host>/:p:/r/<path>.pptx?…` | `ms-powerpoint:ofe\|u\|https://<host>/<path>.pptx` |
| `https://<host>/:x:/r/<path>.xlsx?…` | `ms-excel:ofe\|u\|https://<host>/<path>.xlsx` |
| `https://<host>/:w:/r/<path>.docx?…` | `ms-word:ofe\|u\|https://<host>/<path>.docx` |
| `https://<host>/:p:/r/<site>/_layouts/15/Doc.aspx?sourcedoc={ID}&file=…` | The file ID is resolved to the file path via the same tenant's SharePoint REST API (`GetFileById`), then launched as above |

- `ofe` opens for editing and `ofv` opens read-only (configurable).
- Query strings and fragments are dropped; percent-encoding in the path is preserved.
- Not supported (opened in the browser): sharing links such as `/:p:/g/…` and `/:p:/s/…` that do not contain the file path, non-allowed hosts, `http:`, explicit ports, and URLs with credentials.

## Install

### From the store
Install from the Chrome Web Store or Edge Add-ons, then enter your tenant on the settings page.

### Manually (unpacked)
1. Download `sharepoint-office-launcher-<version>.zip` from [Releases](https://github.com/shunsukeaihara/sharepoint-office-launcher/releases) and extract it.
2. Open `chrome://extensions` (or `edge://extensions`) and turn on Developer mode.
3. Click **Load unpacked** and select the `extension` folder.

## Settings

| Setting | Description | Default |
|---|---|---|
| Enabled | Turns the extension on or off | On |
| Allowed tenants | One host per line. `contoso` means `contoso.sharepoint.com`. Add `contoso-my.sharepoint.com` for OneDrive | **Empty** |
| Open mode | `ofe` (edit) / `ofv` (read-only) | `ofe` |
| Launch automatically | When off, choose on the launch page | On |
| Close the tab after launching | For environments where the confirmation dialog is suppressed by policy | Off |

## Enterprise deployment

- **Force install**: add `<extension ID>;https://clients2.google.com/service/update2/crx` (Chrome) or `<extension ID>;https://edge.microsoft.com/extensionwebstorebase/v1/crx` (Edge) to `ExtensionInstallForcelist`.
- **Managed settings**: the keys in [`managed_schema.json`](extension/managed_schema.json) (`enabled`, `allowedHosts`, `openMode`, `autoLaunch`, `closeTabAfterLaunch`) can be set through policy. Managed values override user settings and are locked on the settings page.
  - Windows: `HKLM\SOFTWARE\Policies\Google\Chrome\3rdparty\extensions\<ID>\policy` (Edge: `...\Microsoft\Edge\3rdparty\...`). Lists are subkeys with numbered values (`allowedHosts\1`, `allowedHosts\2`, …).
  - macOS: configuration profile for the domain `com.google.Chrome.extensions.<ID>` (Edge: `com.microsoft.Edge.extensions.<ID>`).
- **Suppress the confirmation dialog** with `AutoLaunchProtocolsFromOrigins`. The Office URI is launched from the extension's own page, so the origin must be the **extension origin**, not SharePoint:

  ```json
  [
    { "protocol": "ms-powerpoint", "allowed_origins": ["chrome-extension://<extension ID>"] },
    { "protocol": "ms-excel",      "allowed_origins": ["chrome-extension://<extension ID>"] },
    { "protocol": "ms-word",       "allowed_origins": ["chrome-extension://<extension ID>"] }
  ]
  ```

  - Use `chrome-extension://` in Edge as well. Do not add a path or trailing slash.
  - Chrome and Edge are configured separately, and the extension ID differs between distribution channels (unpacked, Chrome Web Store, Edge Add-ons).
  - Never use `"*"` in `allowed_origins`.

See the Japanese [README_jp.md](README_jp.md) for detailed steps (GPO / registry / MDM examples, self-hosted CRX, verification checklist).

## Privacy and security

- No data is collected or sent to the developer or any third party.
- For Doc.aspx links only, the launch page sends a read-only GET request to the same allowed tenant (`/_api/web/GetFileById('<ID>')?$select=ServerRelativeUrl`) using the browser's existing sign-in. The returned path must be inside the linked site and passes the same validation as `/:x:/r/` links.
- Exact host matching against the allowlist; `https:` only; rejects ports, credentials, `|` (the Office URI delimiter), encoded slashes, dot segments, and control characters.
- Minimal permissions: `webNavigation`, non-blocking `webRequest`, `storage`, and `https://*.sharepoint.com/*`. No content scripts, no `tabs`, no `<all_urls>`, no `web_accessible_resources`.
- Loop prevention: converted URLs never match the intercepted pattern; duplicate events are collapsed; "Open in browser" registers a one-time bypass; returning with the Back button does not relaunch.

## Limitations

- The browser asks for confirmation before launching the app unless suppressed by policy.
- Doc.aspx links require being signed in to SharePoint in the browser; blocking third-party cookies or using incognito may cause the lookup to fail (use "Open in browser" in that case).
- Only `*.sharepoint.com` is supported (not sovereign clouds).
- The Office desktop apps must be installed and signed in with an account for the tenant.

## Development

```bash
npm test
```

```bash
npm i -D playwright
```

```bash
npx playwright install chromium
```

```bash
npm run test:e2e
```

```bash
npm run package
```

- Unit tests cover URL conversion and validation, loop prevention, the manifest, and the consistency of the English and Japanese message catalogs.
- The E2E suite loads the extension into Playwright's Chromium in both English and Japanese UIs, against a local fake SharePoint over HTTPS. It never contacts a real tenant.
- Store listing texts are in [`store-listing/`](store-listing/), and images are in `store-assets/<lang>/`. Regenerate the images with `node scripts/render-store-assets.mjs`.

This extension is an unofficial tool and is not affiliated with Microsoft. Microsoft, SharePoint, PowerPoint, Excel, Word, and Office are trademarks of the Microsoft group of companies.
