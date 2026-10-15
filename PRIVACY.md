# プライバシーポリシー / Privacy Policy

SharePoint Office Launcher（以下「本拡張機能」）の個人情報の取り扱いについて説明します。

## 収集する情報

本拡張機能は、個人情報・閲覧履歴・ファイルの内容・認証情報を含め、**いかなるデータも収集・送信しません**。開発者や第三者のサーバーとの通信は一切行いません。

## ブラウザ内に保存する情報

本拡張機能は、動作に必要な次の情報をブラウザの拡張機能用ストレージにのみ保存します。

| 保存先 | 内容 | 保存期間 |
|---|---|---|
| `chrome.storage.sync` | ユーザーが登録した設定（対象とする SharePoint テナントのホスト名、開き方、自動起動の有無など） | ユーザーが変更・削除するまで（ブラウザの同期機能が有効な場合は、ブラウザ提供元のアカウントに同期されます） |
| `chrome.storage.session` | 「ブラウザで開く」を選んだリンクを再び横取りしないための一時情報（タブ ID と URL） | 最長 60 秒。ブラウザ終了時に消去 |

組織の管理者がポリシーで配布した設定（`chrome.storage.managed`）は読み取るだけで、変更や送信はしません。

## 権限の利用目的

- `webNavigation` / `webRequest` / `https://*.sharepoint.com/*`：ユーザーが登録した SharePoint テナントの Office ファイルリンクへのナビゲーションを検出するためだけに使用します。観測のみで、通信の内容やページの内容・Cookie は読み取りません。
- `storage`：上記の設定と一時情報を保存するために使用します。

## 第三者への提供

データを収集しないため、第三者に提供・販売することはありません。

## 変更

本ポリシーを変更する場合は、このページを更新します。

## お問い合わせ

https://github.com/shunsukeaihara/sharepoint-office-launcher/issues

---

## English summary

SharePoint Office Launcher does **not collect, transmit, or sell any data**. It makes no network requests to the developer or any third party. Settings you enter (allowed SharePoint tenant host names, open mode, etc.) are stored only in the browser's extension storage (`chrome.storage.sync`), and a short-lived bypass entry (tab ID and URL, kept up to 60 seconds) is stored in `chrome.storage.session`. The `webNavigation`, `webRequest`, and `https://*.sharepoint.com/*` permissions are used solely to detect navigations to Office file links on the SharePoint tenants you register; request bodies, page contents, and cookies are never read.
