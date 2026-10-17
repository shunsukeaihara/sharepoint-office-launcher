# Store listing (English)

Copy each block into the Chrome Web Store / Edge Add-ons dashboard. The store does not render Markdown, so the description is plain text.

## Description

```text
Open SharePoint PowerPoint, Excel, and Word links from Slack, email, and other apps directly in the Office desktop apps instead of the browser.

Clicking a SharePoint sharing link normally opens Office for the web in your browser, and you have to choose "Open in Desktop App" again to edit it there. With this extension, links from the tenants you register open straight in the desktop app.

■ Features
・Detects SharePoint PowerPoint (/:p:/r/), Excel (/:x:/r/), and Word (/:w:/r/) links and opens them in PowerPoint / Excel / Word
・Also supports ".../_layouts/15/Doc.aspx?sourcedoc=..." links (the file location is looked up in SharePoint before opening)
・Choose between edit mode and read-only mode
・The launch page also offers "Open read-only", "Open in browser", and "Copy file URL"
・Acts only on the tenants you register; links to any other site are never touched
・Works with OneDrive (<tenant>-my.sharepoint.com) as well
・Available in English and Japanese (follows your browser language)

■ How to use
1. The settings page opens right after installation
2. Enter your SharePoint tenant name under "Allowed tenants" (e.g. contoso or contoso.sharepoint.com) and save
3. Click a SharePoint Office file link in Slack or elsewhere, and the desktop app launches
Note: The first time, your browser may ask whether to open the app. Choose "Open".
Note: To open Doc.aspx links, you need to be signed in to SharePoint in the browser.

■ For organizations
・Administrators can distribute settings such as allowed tenants through policy (managed storage)
・The browser policy AutoLaunchProtocolsFromOrigins can suppress the confirmation dialog when launching

■ Privacy
・No data is collected. The extension never communicates with the developer or any third party
・Only when opening a Doc.aspx link, it asks the same (registered) SharePoint tenant for the file location. That information is neither stored nor sent anywhere else
・Page contents and cookies are never read
・Only your settings (allowed tenants, open mode, etc.) are stored

■ Requirements
・Google Chrome or Microsoft Edge (version 110 or later)
・Microsoft Office desktop apps (Windows / macOS) installed and signed in with an account for the target tenant

■ Unsupported links
・Sharing links in the /:p:/g/ or /:p:/s/ format (for example, "Anyone with the link") do not contain the file location and cannot be converted. They open in the browser as usual

This extension is an unofficial tool and is not affiliated with Microsoft. Microsoft, SharePoint, PowerPoint, Excel, Word, and Office are trademarks of the Microsoft group of companies.
```

## Single purpose

```text
Open PowerPoint / Excel / Word file links on SharePoint tenants registered by the user directly in the Office desktop apps instead of the browser.
```

## Permission justification: webNavigation

```text
Used to detect SharePoint Office file links passed to the default browser from apps such as Slack, before the page starts loading. The supported formats are:
・https://<tenant>.sharepoint.com/:p:/r/... (likewise :x: and :w:)
・https://<tenant>.sharepoint.com/:p:/r/.../_layouts/15/Doc.aspx?sourcedoc=...

chrome.webNavigation.onBeforeNavigate is filtered to URLs whose host ends with ".sharepoint.com" and whose path starts with "/:". Only top-level navigations in a tab are handled; subframes and prerendering are ignored. If the URL belongs to a tenant the user registered, the tab is switched to the extension's launch page, which opens the file in the Office desktop app (PowerPoint / Excel / Word).

Browsing history is never stored or transmitted, and the extension does not touch any navigation that does not match these conditions.
```

## Permission justification: webRequest

```text
Used to also detect SharePoint file links that are reached through a redirect (such as HTTP 302). webNavigation does not report URLs reached in the middle of a redirect chain, so webRequest fills that gap.

chrome.webRequest.onBeforeRequest is filtered to requests of type main_frame whose URL matches "https://*.sharepoint.com/:*". It is used for observation only: no blocking, no cancelling or rewriting of requests, no reading or modifying of headers, and no access to request or response bodies. The URL is simply passed to the same check used for webNavigation, with duplicates removed so that each navigation is handled once.
```

## Permission justification: storage

```text
Used to store the extension's settings and short-lived data needed for it to work.

・chrome.storage.sync: settings entered by the user on the options page (host names of the SharePoint tenants to handle, open mode [edit / read-only], whether to launch automatically, etc.).
・chrome.storage.session: when the user chooses "Open in browser" on the launch page, a temporary exclusion entry is kept for up to 60 seconds so the same link is not intercepted again. It is cleared when the browser closes.
・chrome.storage.managed: reads settings (such as allowed tenants) distributed by the organization's administrator through policy.

Only settings are stored. Browsing history, file contents, file locations obtained from SharePoint, and personal information are never stored, and nothing is sent outside the browser.
```

## Permission justification: host permission

```text
The host permission "https://*.sharepoint.com/*" is required for two purposes:

1. To observe top-level requests to SharePoint with webRequest (Chrome requires host permission for a host in order to receive webRequest events for it).

2. To find the actual file location for links in the form "/:p:/r/.../_layouts/15/Doc.aspx?sourcedoc={file ID}". These URLs do not include the file's folder path, so the launch page sends a GET request, using the browser's existing sign-in, to the SharePoint REST API of the same tenant and site as the link (/_api/web/GetFileById('<file ID>')?$select=ServerRelativeUrl). Only the file's server path is retrieved; it is used solely to launch the Office app and is neither stored nor sent anywhere other than SharePoint. No data on SharePoint is modified.

Each organization's SharePoint tenant has its own host name, "<tenant>.sharepoint.com", and the user's tenant cannot be known at publishing time, so *.sharepoint.com is requested. Detection and lookups are performed only for tenants the user registered on the options page (or that an administrator specified through policy). For any other host the extension does nothing and the link opens in the browser as usual.

This permission is not used to inject scripts into pages or to read or modify page contents or cookies (no content scripts are used). No permission for sites other than SharePoint is requested.
```

## Are you using remote code?

```text
No. All code is included in the package; no external scripts are loaded and nothing is executed with eval. The SharePoint API lookup only returns JSON data (a file path), which is never executed as code.
```

## Privacy practices tab

- Data usage: select "none" for every category of collected data. The file location obtained from SharePoint is neither sent to the developer nor stored on the device, so it is not "collected".
- Privacy policy URL: https://github.com/shunsukeaihara/sharepoint-office-launcher/blob/main/PRIVACY.md
- Support / homepage URL: https://github.com/shunsukeaihara/sharepoint-office-launcher

## Images (`store-assets/en/`)

| File | Use |
|---|---|
| `screenshot-1-launch.png` (1280×800) | Screenshot (launch page) |
| `screenshot-2-options.png` (1280×800) | Screenshot (settings page) |
| `promo-small-440x280.png` | Small promo tile |
| `promo-marquee-1400x560.png` | Marquee promo tile (Chrome Web Store) / large promo tile (Edge Add-ons) |
