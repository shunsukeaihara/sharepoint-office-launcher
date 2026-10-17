# SharePoint Office Launcher（Chrome / Edge 拡張機能）

[English](README.md) | 日本語

Slack などから既定ブラウザ（Chrome / Edge）に渡された SharePoint の
`/:p:/r/...pptx`・`/:x:/r/...xlsx`・`/:w:/r/...docx` リンク（`/:p:/r/…/_layouts/15/Doc.aspx?sourcedoc=…` 形式を含む）を検出し、
実ファイル URL に変換して **Office URI スキーム**（`ms-powerpoint:` / `ms-excel:` / `ms-word:`）で
デスクトップ版 Office を直接起動する Manifest V3 拡張機能です。

- 対象は **ユーザー（または組織ポリシー）が許可リストに登録したテナントのみ**。既定の許可リストは空で、インストール直後に設定画面が開きます。登録するまでは一切介入しません。
- 変換できないリンク・許可外のリンクは一切触らず、通常どおりブラウザで開きます。

- ソースコード: https://github.com/shunsukeaihara/sharepoint-office-launcher
- プライバシーポリシー: [PRIVACY.md](PRIVACY.md)
- ライセンス: [MIT](LICENSE)

---

## 目次

1. [動作の仕組み](#1-動作の仕組み)
2. [同梱物](#2-同梱物)
3. [手動導入（Chrome / Edge）](#3-手動導入chrome--edge)
4. [初期設定（許可テナント）](#4-初期設定許可テナント)
5. [組織配布](#5-組織配布)
6. [AutoLaunchProtocolsFromOrigins（確認ダイアログの抑止）](#6-autolaunchprotocolsfromorigins確認ダイアログの抑止)
7. [ループ防止](#7-ループ防止)
8. [セキュリティ設計](#8-セキュリティ設計)
9. [fallback（自動起動できないとき）](#9-fallback自動起動できないとき)
10. [ブラウザ・環境の制限事項](#10-ブラウザ環境の制限事項)
11. [検証手順（macOS / Windows）](#11-検証手順macos--windows)
12. [テスト](#12-テスト)
13. [トラブルシューティング](#13-トラブルシューティング)
14. [変更履歴・更新時の注意](#14-変更履歴更新時の注意)

---

## 1. 動作の仕組み

```
Slack でリンクをクリック
   │  https://contoso.sharepoint.com/:p:/r/sites/Sales/Shared%20Documents/提案.pptx?d=w…&csf=1&web=1&e=…
   ▼
既定ブラウザ（Chrome / Edge）でナビゲーション開始
   │  background.js（Service Worker）が webNavigation / webRequest でトップレベル遷移を検知
   │  ・ホストが許可リストにあるか ・/:p|x|w:/r/ 形式か ・拡張子が種別と一致するか を検証
   ▼
タブを拡張機能内ページ launch.html に差し替え（SharePoint の読み込みは中断）
   │  launch.html が設定を読み直して再検証し、Office URI を組み立てる
   │  ms-powerpoint:ofe|u|https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/提案.pptx
   ▼
location.href = Office URI  →  ブラウザの外部プロトコル処理 →  PowerPoint が起動
```

`/:p:/r/…/_layouts/15/Doc.aspx?sourcedoc={ファイルID}&file=…` 形式のリンクは、URL にファイルのフォルダパスが含まれていません。
そのため launch.html が、同じテナント・同じサイトの SharePoint REST API
`/_api/web/GetFileById('<ファイルID>')?$select=ServerRelativeUrl` に、ブラウザのサインイン状態（Cookie）で問い合わせて実ファイルの場所を取得してから起動します（[1.1](#11-docaspx-形式のリンク)）。

### 変換規則

| 入力リンク | Office URI |
|---|---|
| `https://<host>/:p:/r/<path>.pptx?…` | `ms-powerpoint:ofe\|u\|https://<host>/<path>.pptx` |
| `https://<host>/:x:/r/<path>.xlsx?…` | `ms-excel:ofe\|u\|https://<host>/<path>.xlsx` |
| `https://<host>/:w:/r/<path>.docx?…` | `ms-word:ofe\|u\|https://<host>/<path>.docx` |

- `/:<種別>:/r` を取り除き、クエリ（`d=`, `csf=`, `web=`, `e=` など）とフラグメントは破棄します。
- パスのパーセントエンコード（`%20`、日本語ファイル名など）はそのまま維持します。
- `ofe` = 編集モード、`ofv` = 読み取り専用。設定で切り替え可能です。
- 対応拡張子: PowerPoint `pptx pptm ppt ppsx potx` / Excel `xlsx xlsm xlsb xls` / Word `docx docm doc dotx`

### 1.1 Doc.aspx 形式のリンク

例:

```
https://contoso.sharepoint.com/:p:/r/sites/pj-sample/_layouts/15/Doc.aspx?sourcedoc=%7BAA3B8D18-237C-446C-9D6E-F98E2E9A8FBA%7D&file=%E8%B3%87%E6%96%99.pptx&action=edit
```

1. `/:p:/r` の後ろからサイトのパス（`/sites/pj-sample`）を、`sourcedoc` からファイル ID（GUID）を取り出します。`file` パラメータがあれば、拡張子が種別と一致することも確認します。
2. launch.html が `https://contoso.sharepoint.com/sites/pj-sample/_api/web/GetFileById('aa3b8d18-…')?$select=ServerRelativeUrl` に GET で問い合わせます（Cookie 付き、リダイレクトは追わない、タイムアウト 10 秒）。
3. 返ってきた `ServerRelativeUrl`（例 `/sites/pj-sample/Shared Documents/資料.pptx`）が、リンクのサイト配下にあることを確認します。そのうえで `/:p:/r/` 形式と同じ検証（許可テナント・拡張子・禁止文字・ドットセグメント）を通してから Office URI を組み立てます。
4. 問い合わせに失敗した場合（未サインイン・アクセス権なし・ファイルなし・タイムアウトなど）は、理由と「ブラウザで開く」ボタンを表示します。

- 問い合わせ先は、リンクと同じ許可済みテナントの SharePoint だけです。取得した情報はどこにも送信・保存しません。
- `/:p:/r/` が付かない `…/_layouts/15/Doc.aspx` のリンク（SharePoint の画面内でファイルをクリックしたときなど）には介入しません。

### 対象外（ブラウザでそのまま開く）

- `/:p:/g/…`、`/:p:/s/…` などの **共有リンク（パスを含まない形式）**。実ファイル URL をクライアント側で復元できないため。
- 許可リスト外のホスト、`http:`、ポート指定、認証情報付き URL
- 種別と拡張子が一致しないもの（例 `/:p:/r/…/a.xlsx`）、PDF（`/:b:/`）など

---

## 2. 同梱物

```
sharepoint-office-launcher-1.3.0/
├── README.md               … README（英語）
├── README_jp.md            … 本書（日本語）
├── package.json            … テスト / パッケージ用スクリプト定義
├── extension/              … ★ブラウザに読み込むフォルダ（manifest.json があるフォルダ）
│   ├── manifest.json
│   ├── background.js       … ナビゲーション検知・タブ差し替え
│   ├── launch.html / launch.js … Office URI 起動ページ（fallback ボタン付き）
│   ├── options.html / options.js … 設定画面
│   ├── managed_schema.json … 組織ポリシーで設定を配布するためのスキーマ
│   ├── lib/convert.js      … URL 変換・検証（純粋関数）
│   ├── lib/guard.js        … 二重処理・ループ防止（純粋関数）
│   ├── lib/config.js       … 設定の解決（ポリシー > ユーザー設定 > 既定値）
│   ├── lib/resolve.js      … Doc.aspx 形式のファイル ID から実ファイルの場所を取得
│   ├── lib/i18n.js         … 表示文字列の多言語化（chrome.i18n）
│   ├── _locales/{en,ja}/messages.json … 英語（既定）・日本語のメッセージ
│   ├── ui.css
│   └── icons/
├── test/                   … node:test による単体テスト、test/e2e/ に Playwright E2E
├── assets/icon.svg        … アイコンの元データ
├── store-assets/{en,ja}/ … ストア掲載用画像（スクリーンショット 1280x800 ×2、プロモタイル 440x280 / 1400x560）
├── store-listing/{en,ja}.md … ストア掲載用の説明文・権限の説明
└── scripts/
    ├── package.sh          … 配布 zip 作成
    ├── render-icons.mjs    … assets/icon.svg からアイコン PNG を生成（Playwright 使用）
    └── render-store-assets.mjs … 実際の画面を撮影してストア掲載用画像を生成（Playwright 使用）
```

別途 `sharepoint-office-launcher-1.3.0-store.zip`（`manifest.json` がルートにある zip）を
Chrome Web Store / Edge Add-ons へのアップロード用に用意しています（`npm run package` で再生成可能）。

> 本書の `contoso` は Microsoft のドキュメントで使われる架空のテナント名です。実際の自社テナント名に読み替えてください。

### 要求権限

| 権限 | 用途 |
|---|---|
| `webNavigation` | トップレベル遷移の URL を検知 |
| `webRequest` + `https://*.sharepoint.com/*` | リダイレクタ経由（302）で SharePoint に到達したケースの検知。**観測のみ**（ブロッキングなし）。ホスト権限は、Doc.aspx 形式のリンクのファイルの場所を同じテナントの SharePoint API に問い合わせる際にも使います |
| `storage` | 設定・「ブラウザで開く」用の一時バイパス情報 |

`tabs` 権限、`<all_urls>`、コンテンツスクリプト、`web_accessible_resources` は使用しません。
通信するのは、Doc.aspx 形式のリンクについて、リンク先と同じ許可済み SharePoint テナントにファイルの場所を問い合わせるときだけです。それ以外の外部への通信は一切行いません。

---

## 3. 手動導入（Chrome / Edge）

### Chrome

1. zip を展開する（展開先フォルダは削除・移動しないこと。移動すると拡張機能が無効になります）
2. `chrome://extensions` を開く
3. 右上の **デベロッパー モード** をオン
4. **パッケージ化されていない拡張機能を読み込む** → 展開した `extension` フォルダを選択
5. 初回インストール時に設定画面が自動で開くので、[4. 初期設定](#4-初期設定許可テナント) を行う

### Edge

1. zip を展開する
2. `edge://extensions` を開く
3. 左下（または左ペイン）の **開発者モード** をオン
4. **展開して読み込み** → `extension` フォルダを選択
5. 設定画面で許可テナントを登録

> **注意**
> - 手動導入（unpacked）の拡張 ID は、読み込んだフォルダのパスから決まります。PC ごと・フォルダごとに ID が変わる可能性があり、後述の `AutoLaunchProtocolsFromOrigins` に書く ID と一致させる必要があります。ID は設定画面最下部、または `chrome://extensions` / `edge://extensions` で確認できます。
> - Chrome はデベロッパーモードの拡張機能について起動時に警告を出すことがあります。恒常運用は [5. 組織配布](#5-組織配布) を推奨します。
> - ID を固定したい場合は、`manifest.json` に `"key"`（公開鍵）を追加すると unpacked でも ID が固定されます（Web Store 配布時はストア側の ID になるため `key` は不要）。

---

## 4. 初期設定（許可テナント）

`chrome://extensions`（`edge://extensions`）→ 本拡張の **詳細** → **拡張機能のオプション**。

| 項目 | 説明 | 既定値 |
|---|---|---|
| 有効にする | 全体のオン／オフ | オン |
| 許可するテナント | 1 行 1 ホスト。`contoso` と書くと `contoso.sharepoint.com`。OneDrive（個人用サイト）も対象なら `contoso-my.sharepoint.com` を追加 | **空（何もしない）** |
| 開き方 | `ofe`（編集）/ `ofv`（読み取り専用） | `ofe` |
| 自動でアプリを起動する | オフにすると launch.html でボタン選択 | オン |
| 起動後 4 秒でタブを閉じる | 確認ダイアログを出さない環境（ポリシー適用済み）向け | オフ |

- 表示言語はブラウザの表示言語に合わせて英語・日本語が自動で切り替わります（それ以外の言語では英語）。
- ホストは **完全一致**で照合します。`*.sharepoint.com` のようなワイルドカードは受け付けません。
- 組織ポリシー（`chrome.storage.managed`）で設定された項目は、設定画面でグレーアウトされ変更できません。

---

## 5. 組織配布

### 5.0 一般公開（Chrome Web Store / Edge Add-ons）

本拡張は特定テナントを埋め込んでいないため、そのまま一般公開できます。インストールしたユーザーは初回に開く設定画面で自分のテナントを登録します。

- `sharepoint-office-launcher-1.3.0-store.zip` をそれぞれのストアに提出し、公開範囲を「公開（Public）」にします。
- 審査用の文言（説明文・単一用途・権限の説明・リモートコード・プライバシー）は、言語別に [store-listing/ja.md](store-listing/ja.md)・[store-listing/en.md](store-listing/en.md) にまとめてあります。
- 掲載用画像は `store-assets/ja/`・`store-assets/en/` にあります（各言語の UI で撮影）。

  | ファイル | 用途 |
  |---|---|
  | `screenshot-1-launch.png`（1280×800） | スクリーンショット（起動画面） |
  | `screenshot-2-options.png`（1280×800） | スクリーンショット（設定画面） |
  | `promo-small-440x280.png` | Chrome Web Store 小プロモタイル（必須）／ Edge Add-ons 小プロモタイル |
  | `promo-marquee-1400x560.png` | Chrome Web Store マーキー（任意）／ Edge Add-ons 大プロモタイル |

  ストアのアイコンには `extension/icons/icon128.png` を使います。画面を変更したら `node scripts/render-store-assets.mjs` で両言語分を作り直せます。
- **多言語の掲載**：拡張機能の名前と説明は `_locales` で多言語化しているため、ストアのダッシュボードで言語ごとに説明文・画像を登録できます。Chrome Web Store では「ストアの掲載情報」の言語選択で English / 日本語 を切り替えて、それぞれ入力します。Edge Add-ons では「Store listings」で言語を追加します。既定の言語は英語（`default_locale: en`）です。
- 企業はストア版の ID を使い、以下 5.1〜5.3 の方法で強制インストールとテナント設定の一括配布ができます。

配布方法は大きく 2 通りです。

| 方式 | 長所 | 注意 |
|---|---|---|
| A. ストア経由（Chrome Web Store 非公開/限定公開、Edge Add-ons 非公開） | 自動更新・ID 固定・審査あり | 公開手続きが必要 |
| B. 自己ホスト CRX + update manifest | ストア不要 | Windows はドメイン参加 / Chrome Enterprise Core（CBCM）・Intune 等の管理下、macOS は MDM 管理下でないと強制インストールが効かない |

### 5.1 Chrome

#### A. Chrome Web Store（推奨）

1. `sharepoint-office-launcher-1.3.0-store.zip` を Chrome Web Store デベロッパーダッシュボードにアップロード
2. 公開設定を「非公開（Private: 自組織の Google Workspace ドメインのみ）」または「限定公開（Unlisted）」にする
3. 発行された **拡張 ID**（32 文字）を控える
4. ポリシーで強制インストール

**Google 管理コンソール**（Chrome Enterprise Core / Workspace）:
デバイス → Chrome → アプリと拡張機能 → ユーザーとブラウザ → ID で追加 → 「自動インストール」。
同画面の「ポリシー（JSON）」欄で拡張機能の設定（後述 5.3）も配布できます。

**Windows（GPO / レジストリ）**:

```
HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist
  "1" (REG_SZ) = "<拡張ID>;https://clients2.google.com/service/update2/crx"
```

GPO の場合は Chrome の ADMX テンプレートを導入し、
「Google Chrome → 拡張機能 → 自動インストールするアプリと拡張機能のリストの設定」に同じ値を入れます。

**macOS（構成プロファイル / MDM）**: ドメイン `com.google.Chrome`

```xml
<key>ExtensionInstallForcelist</key>
<array>
  <string>&lt;拡張ID&gt;;https://clients2.google.com/service/update2/crx</string>
</array>
```

#### B. 自己ホスト

1. `chrome://extensions` →「拡張機能をパッケージ化」で `extension` フォルダから `.crx` と `.pem`（秘密鍵）を作成。**`.pem` は厳重に保管**（更新時に同じ鍵が必要で、ID は鍵から決まる）
2. `.crx` と update manifest（XML）を社内 HTTPS サーバーに配置

```xml
<?xml version='1.0' encoding='UTF-8'?>
<gupdate xmlns='http://www.google.com/update2/response' protocol='2.0'>
  <app appid='<拡張ID>'>
    <updatecheck codebase='https://intra.example.com/ext/sp-office-launcher-1.3.0.crx' version='1.3.0' />
  </app>
</gupdate>
```

3. `ExtensionInstallForcelist` に `<拡張ID>;https://intra.example.com/ext/updates.xml` を設定
4. 必要に応じて `ExtensionInstallSources` にホスト URL を追加

### 5.2 Edge

#### A. Edge Add-ons（推奨）

Microsoft Partner Center で `…-store.zip` を提出し、可視性を「非公開（Hidden）」にして ID を取得。

**Windows（GPO / Intune / レジストリ）**:

```
HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist
  "1" (REG_SZ) = "<拡張ID>;https://edge.microsoft.com/extensionwebstorebase/v1/crx"
```

Intune では「構成プロファイル → 設定カタログ → Microsoft Edge → 拡張機能 → Control which extensions are installed silently」。

**macOS**: ドメイン `com.microsoft.Edge` に同じキー `ExtensionInstallForcelist` を構成プロファイルで配布。

#### B. 自己ホスト

Chrome と同様に `.crx` と update manifest を用意し、`ExtensionInstallForcelist` に update manifest URL を指定します。
Edge で作成した CRX と Chrome で作成した CRX は同じ `.pem` を使えば同じ ID になります。

### 5.3 拡張機能の設定をポリシーで配布（managed storage）

`managed_schema.json` に従い、許可テナント等を組織から一括配布できます。ポリシーで設定した項目はユーザー設定より優先され、ユーザーは変更できません。

設定例（JSON 表現）:

```json
{
  "enabled": true,
  "allowedHosts": ["contoso.sharepoint.com", "contoso-my.sharepoint.com"],
  "openMode": "ofe",
  "autoLaunch": true,
  "closeTabAfterLaunch": false
}
```

| ブラウザ / OS | 配置先 |
|---|---|
| Chrome / Windows | `HKLM\SOFTWARE\Policies\Google\Chrome\3rdparty\extensions\<拡張ID>\policy` |
| Edge / Windows | `HKLM\SOFTWARE\Policies\Microsoft\Edge\3rdparty\extensions\<拡張ID>\policy` |
| Chrome / macOS | 構成プロファイルのドメイン `com.google.Chrome.extensions.<拡張ID>` |
| Edge / macOS | 構成プロファイルのドメイン `com.microsoft.Edge.extensions.<拡張ID>` |
| Chrome（クラウド管理） | Google 管理コンソールの拡張機能詳細「ポリシー（JSON）」 |

Windows レジストリでの配列は、サブキーに連番の値を作ります:

```
...\3rdparty\extensions\<拡張ID>\policy
    enabled      REG_DWORD 1
    openMode     REG_SZ    ofe
    autoLaunch   REG_DWORD 1
...\3rdparty\extensions\<拡張ID>\policy\allowedHosts
    1            REG_SZ    contoso.sharepoint.com
    2            REG_SZ    contoso-my.sharepoint.com
```

反映確認: `chrome://policy`（`edge://policy`）で「ポリシーを再読み込み」→ 拡張機能の設定画面で項目がグレーアウトされていること。

---

## 6. AutoLaunchProtocolsFromOrigins（確認ダイアログの抑止）

既定では、Office URI を起動するたびにブラウザが
「Microsoft PowerPoint を開きますか？」という確認ダイアログを表示します。
組織で抑止したい場合は `AutoLaunchProtocolsFromOrigins` ポリシーを使います。

### 6.1 正しい origin は「拡張機能の origin」

本拡張では **Office URI へ遷移するのは拡張機能内ページ `launch.html`** です。
したがって外部プロトコル起動の「起動元 origin」は SharePoint ではなく

```
chrome-extension://<拡張ID>
```

になります（**Edge でもスキームは `chrome-extension://`** です。`edge-extension://` ではありません）。

`https://contoso.sharepoint.com` を指定しても、本拡張経由の起動ダイアログは抑止されません。
（SharePoint を許可すると、SharePoint の Web 画面自身の「デスクトップアプリで開く」が無確認になります。必要であれば別途判断のうえ追加してください。）

### 6.2 設定値

値は JSON 配列です。`protocol` は **コロンなし** のスキーム名です。

```json
[
  { "protocol": "ms-powerpoint", "allowed_origins": ["chrome-extension://<拡張ID>"] },
  { "protocol": "ms-excel",      "allowed_origins": ["chrome-extension://<拡張ID>"] },
  { "protocol": "ms-word",       "allowed_origins": ["chrome-extension://<拡張ID>"] }
]
```

**Windows（レジストリ / GPO）**: 値は 1 行の JSON 文字列（REG_SZ）

```
HKLM\SOFTWARE\Policies\Google\Chrome\AutoLaunchProtocolsFromOrigins    (Chrome)
HKLM\SOFTWARE\Policies\Microsoft\Edge\AutoLaunchProtocolsFromOrigins   (Edge)
  REG_SZ = [{"protocol":"ms-powerpoint","allowed_origins":["chrome-extension://<拡張ID>"]},{"protocol":"ms-excel","allowed_origins":["chrome-extension://<拡張ID>"]},{"protocol":"ms-word","allowed_origins":["chrome-extension://<拡張ID>"]}]
```

**macOS（構成プロファイル）**: ドメイン `com.google.Chrome` / `com.microsoft.Edge`

```xml
<key>AutoLaunchProtocolsFromOrigins</key>
<array>
  <dict>
    <key>protocol</key><string>ms-powerpoint</string>
    <key>allowed_origins</key><array><string>chrome-extension://&lt;拡張ID&gt;</string></array>
  </dict>
  <dict>
    <key>protocol</key><string>ms-excel</string>
    <key>allowed_origins</key><array><string>chrome-extension://&lt;拡張ID&gt;</string></array>
  </dict>
  <dict>
    <key>protocol</key><string>ms-word</string>
    <key>allowed_origins</key><array><string>chrome-extension://&lt;拡張ID&gt;</string></array>
  </dict>
</array>
```

### 6.3 注意点

- **Chrome と Edge は別ポリシー**です。両方使う場合はそれぞれ設定してください（レジストリパス・plist ドメインが異なる）。
- **拡張 ID は配布方法ごとに異なります。** unpacked / Chrome Web Store / Edge Add-ons / 自己ホスト CRX で ID が変わるため、実際に配布する ID を書くこと。Chrome 版と Edge 版でストアが違えば ID も違います。
- `allowed_origins` に `"*"` を書くと **すべてのサイトから無確認で Office を起動できる**状態になります。使用しないでください。
- origin にはパスや末尾スラッシュを付けないでください（`chrome-extension://<ID>/launch.html` ではなく `chrome-extension://<ID>`）。
- 設定は「このサイトでは常に許可」チェックボックスと同等で、`URLBlocklist` で `ms-powerpoint:*` 等がブロックされていれば起動されません。
- ブラウザのポリシーは **管理下デバイス**で有効になります。macOS では MDM の構成プロファイル、Windows では GPO / Intune / HKLM レジストリでの配布が確実です。
  （macOS の `defaults write` は検証用途では反映される場合がありますが、本番配布には使わないでください。）
- 適用状況は `chrome://policy` / `edge://policy` で、ステータスが「OK」になっていることを確認します。JSON の書式誤りは「エラー」と表示されます。
- 本拡張の起動経路（拡張機能ページからのトップレベル遷移）でポリシーが期待通り効くかは、**導入前に必ず対象のブラウザ・バージョンで実機確認**してください（[11. 検証手順](#11-検証手順macos--windows) の T7）。ポリシーが効かない場合でも、ダイアログで「開く」を選べば動作します。

---

## 7. ループ防止

| 仕組み | 内容 |
|---|---|
| 構造的に再一致しない | 変換後の実ファイル URL（`/sites/…/x.pptx`）は `/:p:/r/` 形式ではないため、Office が「ブラウザで開く」へ戻しても拡張機能は介入しません |
| 二重処理防止 | 同じタブ・同じ URL のイベントを 5 秒以内に 1 回だけ処理（webNavigation と webRequest の重複発火、再送対策） |
| 「ブラウザで開く」のバイパス | launch.html の「ブラウザで開く」は、そのタブ・その URL に限り 60 秒有効・1 回限りのバイパスを登録してから遷移。再度横取りされません |
| 戻るボタン対策 | launch.html は起動済みを履歴状態に記録し、戻る操作で再表示されたときは自動起動せずボタン表示のみ |
| サブフレーム・プリレンダー除外 | トップレベル（メインフレーム）の通常遷移のみ対象 |

---

## 8. セキュリティ設計

- **許可テナントの完全一致**: `^[a-z0-9-]+\.sharepoint\.com$` 形式かつ許可リストに完全一致するホストのみ。`contoso.sharepoint.com.evil.example` などのサフィックス偽装は拒否。
- **既定で無効相当**: 許可リストは既定で空。ユーザーが登録するまで何もしません。
- **URL 検証**: `https:` のみ、ポート指定・認証情報付き URL を拒否。
- **Office URI インジェクション対策**: パスに `|`（Office URI のコマンド区切り）、エンコード済み `%7C`、エンコード済み `/`・`\`、空白、制御文字が含まれる場合は拒否。不正なパーセントエンコード、空セグメントも拒否。
- **種別と拡張子の整合**: `/:p:/` なら PowerPoint 系拡張子のみ等。合わなければ介入しない。
- **クエリ・フラグメント破棄**: 共有トークン等を Office URI に渡さない。
- **二重検証**: Service Worker で判定後、launch.html 側でも設定を読み直して再検証（launch.html の `src` パラメータを信用しない）。
- **最小権限**: `tabs` / `<all_urls>` / コンテンツスクリプト不使用。`webRequest` は非ブロッキングの観測のみ。`launch.html` は `web_accessible_resources` に含めないため Web ページから呼び出せません。
- **通信は許可テナントへの問い合わせのみ**: Doc.aspx 形式のファイルの場所を取得するときだけ、リンクと同じ許可済みテナントの SharePoint API（`GetFileById`、GET、読み取りのみ）に問い合わせます。それ以外にはどこにも通信しません。
- **API 応答の検証**: 応答の `ServerRelativeUrl` は、絶対パスであること、リンクのサイト配下であること、ドット・空セグメントがないことを確認します。そのうえで `/:x:/r/` 形式と同じ検証を通します。未サインイン時のログイン画面へのリダイレクトは追いません。
- **ファイル ID の検証**: `sourcedoc` は GUID 形式のみ受け付けます（API の URL に任意の文字列が入らないようにするため）。
- **CSP**: 拡張機能ページは `script-src 'self'` のみ。

---

## 9. fallback（自動起動できないとき）

launch.html には常に以下のボタンを表示します（ボタン操作はユーザー操作扱いになるため、自動起動がブロックされた場合も起動できます）。

| ボタン | 動作 |
|---|---|
| ○○ で開く（編集） | `ofe` で再起動 |
| 読み取り専用で開く | `ofv` で起動 |
| ブラウザで開く | 元のリンクをブラウザ（Office for the web）で開く。ループしないようバイパス登録 |
| ファイル URL をコピー | 実ファイル URL をクリップボードへ（Office の「開く」に貼り付け可能） |

Office が未インストール・プロトコル未登録の場合、ブラウザは何も起こさないか、エラーを出します。その場合は「ブラウザで開く」を使ってください。

---

## 10. ブラウザ・環境の制限事項

- **確認ダイアログ**: ポリシー未設定時、Office URI 起動のたびにブラウザの確認ダイアログが出ます（「常に許可」チェックが出るかはブラウザ／起動元によって異なります）。抑止は [6.](#6-autolaunchprotocolsfromorigins確認ダイアログの抑止) を参照。
- **ユーザー操作なしの連続起動**: Chromium は、ユーザー操作を伴わない外部プロトコル起動が短時間に連続するとブロックすることがあります。その場合は launch.html のボタンを押してください。
- **既定ブラウザ**: Slack デスクトップアプリは OS の既定ブラウザでリンクを開きます。既定ブラウザが拡張機能を入れた Chrome / Edge であること、かつリンクが開くプロファイルに拡張機能が入っていることが必要です（複数プロファイル利用時は「最後に使ったプロファイル」で開かれます）。
- **シークレット / InPrivate**: 既定では動作しません（拡張機能の「シークレットモードでの実行を許可」をオンにすれば動作）。
- **Doc.aspx 形式はブラウザでのサインインが必要**: ファイルの場所を SharePoint API に問い合わせるため、ブラウザで対象テナントにサインインしている必要があります。サードパーティ Cookie をブロックする設定やシークレットモードでは問い合わせが失敗する場合があります（その場合は「ブラウザで開く」で開けます）。
- **共有リンク非対応**: `/:p:/g/`・`/:p:/s/`（「リンクを知っているユーザー」等の共有リンク）は実パスを含まないため変換できず、ブラウザで開きます。
- **Office アカウント**: Office デスクトップアプリが対象テナントのアカウントでサインインしている必要があります。未サインインならアプリ側でサインインを求められます。
- **Edge の Office 連携**: Edge には Office ファイルを Edge 内で開く機能や SharePoint の「デスクトップアプリで開く」連携があります。本拡張は `/:p:/r/` の遷移を先に差し替えるため通常は干渉しませんが、組織ポリシーでの動作は実機で確認してください。
- **ソブリンクラウド**: `sharepoint.us`、`sharepoint.cn`、`sharepoint.de` 等は未対応（`*.sharepoint.com` のみ）。
- **Service Worker 起動遅延**: 拡張機能の Service Worker がスリープ中だと、起動までの一瞬 SharePoint の読み込みが始まることがありますが、その後 launch.html に差し替わります。
- **Office for Mac**: Office URI スキームは Office for Mac（Microsoft 365 版 / 2019 以降）で対応しています。古いバージョンでは起動できない場合があります。

---

## 11. 検証手順（macOS / Windows）

### 事前準備

1. Office デスクトップ版（PowerPoint / Excel / Word）をインストールし、対象テナントのアカウントでサインイン
2. Chrome または Edge を **既定のブラウザ** に設定
   - macOS: システム設定 → デスクトップと Dock → デフォルトの Web ブラウザ
   - Windows: 設定 → アプリ → 既定のアプリ → Chrome / Edge →「既定に設定」
3. 拡張機能を導入し、許可テナントを登録（[3.](#3-手動導入chrome--edge)・[4.](#4-初期設定許可テナント)）
4. SharePoint 上に検証用の `.pptx` `.xlsx` `.docx` を用意し、各ファイルの「…」→「リンクのコピー」で
   `/:p:/r/` 形式のリンクを取得（「リンクの設定」で「既存のアクセス権を持つユーザー」を選ぶと `/r/` 形式になります）

### 「Slack から開く」の再現

Slack の代わりに、OS に既定ブラウザで URL を開かせて再現できます（URL は引用符で囲む）。

macOS（ターミナル）:

```bash
open "https://contoso.sharepoint.com/:p:/r/sites/Test/Shared%20Documents/test.pptx?d=w123&csf=1&web=1"
```

Windows（PowerShell）:

```powershell
Start-Process "https://contoso.sharepoint.com/:p:/r/sites/Test/Shared%20Documents/test.pptx?d=w123&csf=1&web=1"
```

最終確認は、実際に Slack（デスクトップアプリ）に同じリンクを投稿してクリックしてください。

### 確認項目

| # | 操作 | 期待結果 |
|---|---|---|
| T1 | `/:p:/r/…pptx` を開く | タブが「PowerPoint で開いています…」になり、PowerPoint が起動して該当ファイルを開く |
| T2 | `/:x:/r/…xlsx`、`/:w:/r/…docx` を開く | Excel / Word が起動 |
| T3 | 日本語・空白を含むファイル名 | 正しいファイルが開く |
| T4 | 許可リスト外テナントのリンク | 通常どおりブラウザで SharePoint が開く（介入なし） |
| T5 | `/:p:/g/…` 共有リンク | 通常どおりブラウザで開く |
| T6 | launch.html の「ブラウザで開く」 | Office for the web で開き、再度 launch.html に戻らない（ループしない） |
| T7 | ポリシー `AutoLaunchProtocolsFromOrigins` 適用後に T1 | 確認ダイアログなしで起動。`chrome://policy` / `edge://policy` でステータス OK |
| T8 | T6 の後にブラウザの「戻る」 | launch.html がボタン表示のみで、自動起動しない |
| T9 | 設定「有効にする」をオフ | 介入しない |
| T10 | ポリシーで allowedHosts を配布 | 設定画面の該当項目がグレーアウトされ、配布値で動作 |
| T11 | Chrome と Edge の両方で T1〜T8 | 同様に動作 |
| T12 | `/:p:/r/…/_layouts/15/Doc.aspx?sourcedoc=…` 形式のリンク（ブラウザでサインイン済み） | 「ファイルを確認しています…」の後、PowerPoint が起動して該当ファイルを開く |
| T13 | T12 をシークレット / InPrivate（未サインイン）で開く | 「サインインしていない」と表示され、「ブラウザで開く」で開ける |
| T14 | T12 でアクセス権のないファイル | 「アクセス権がない」と表示される |

### ログの確認

`chrome://extensions`（`edge://extensions`）→ 本拡張の「Service Worker」リンク（`service worker を検証`）で DevTools を開くと、
エラー時に `[sp-office-launcher]` で始まるログが出ます。

### プロトコル登録の確認（Office 側の問題切り分け）

拡張機能を経由せず、Office URI を直接開いて Office 側が正しく応答するか確認できます。

macOS:

```bash
open "ms-powerpoint:ofe|u|https://contoso.sharepoint.com/sites/Test/Shared%20Documents/test.pptx"
```

Windows（PowerShell）:

```powershell
Start-Process "ms-powerpoint:ofe|u|https://contoso.sharepoint.com/sites/Test/Shared%20Documents/test.pptx"
```

---

## 12. テスト

Node.js 20 以上で、URL 変換・検証、ループ防止ロジック、manifest の整合性、多言語メッセージの整合性（英語と日本語でキー・プレースホルダーが一致し、画面から参照するキーがすべて存在すること）を検証します。

```bash
npm test
```

ブラウザ実機での E2E（Playwright の Chromium に拡張機能を読み込み、横取り・素通り・Doc.aspx 形式の解決〔Cookie 付きの API 問い合わせ・未サインイン・403〕・「ブラウザで開く」・戻る対策・SW 再起動後の動作を確認。`contoso.sharepoint.com` はローカルの偽 SharePoint〔自己署名証明書の HTTPS。生成に openssl を使用〕に向け、他の `*.sharepoint.com` は名前解決させないため、実テナントには通信しません）:

```bash
npm i -D playwright
```

```bash
npx playwright install chromium
```

```bash
npm run test:e2e
```

E2E は英語・日本語の両方の UI で実行し、画面の文言を `_locales` のメッセージと照合します（`E2E_LANG=en` / `E2E_LANG=ja` で片方だけ実行することもできます）。macOS の Chromium は `--lang` ではなく OS の言語設定に従うため、macOS では Chromium を直接起動して `-AppleLanguages` で言語を指定し、CDP で接続しています（`test/e2e/launch-chromium.mjs`）。

Office アプリの実起動はヘッドレスでは確認できないため、[11. 検証手順](#11-検証手順macos--windows) で手動確認してください。

配布 zip の再作成:

```bash
npm run package
```

---

## 13. トラブルシューティング

| 症状 | 確認点 |
|---|---|
| 何も起きずブラウザで SharePoint が開く | 許可リストにホストが入っているか / リンクが `/r/` 形式か / 拡張機能が有効か / リンクが開いたプロファイルに拡張機能が入っているか |
| launch.html は出るがアプリが起動しない | 確認ダイアログが背面に出ていないか / Office がインストール・プロトコル登録済みか（11. の直接起動で確認）/ ボタンで再試行 |
| 毎回確認ダイアログが出る | `AutoLaunchProtocolsFromOrigins` の origin が `chrome-extension://<実際の拡張ID>` か / `chrome://policy` でエラーになっていないか / Chrome と Edge を取り違えていないか |
| Office で「ファイルが見つからない」 | Office のサインインアカウントが対象テナントか / ファイルへのアクセス権があるか |
| 設定画面で変更できない | 組織ポリシーで管理されている項目です |

---

## 14. 変更履歴・更新時の注意

### 1.3.0

- 英語に対応しました（既定は英語。ブラウザの表示言語が日本語なら日本語で表示）。拡張機能の名前・説明、起動ページ、設定画面、エラーメッセージを多言語化しています。
- 英語版・日本語版のストア掲載用の説明文（`store-listing/`）と画像（`store-assets/en/`・`store-assets/ja/`）を用意しました。
- `short_name` をストア推奨の 12 文字以内（`SPO Launcher`）にしました。
- 組織ポリシー用スキーマ（managed_schema.json）の項目名を英語にしました。
- README を英語（`README.md`）に変更し、日本語版を `README_jp.md` に移しました。
- プライバシーポリシー（`PRIVACY.md`）に英語の全文を加え、英語・日本語の両方を 1 ファイルに記載しました。

### 1.2.0

- `/:p|x|w:/r/…/_layouts/15/Doc.aspx?sourcedoc={ID}&file=…` 形式のリンクに対応しました。ファイル ID から実ファイルの場所を、同じテナントの SharePoint API で取得します。
- これに伴い、プライバシーポリシーとストアの権限説明を更新しました（許可テナントの SharePoint にのみ問い合わせる）。

### 1.1.0

- 許可テナントの既定値を廃止しました（既定は空。ユーザーまたは組織ポリシーで登録）。
- 更新後に許可テナントが未設定の場合、設定画面を自動で開いて登録を促します。
- エラー時に操作ボタンが表示されてしまう表示不具合を修正しました。
- アイコンを作り直し、ストア掲載用画像を追加しました。

### ストアで公開済みの版から更新するとき

- **バージョン番号は必ず上げてください。** ストアは公開中の版以下の番号を受け付けず、配布済みのブラウザも番号が上がったときだけ自動更新します。
- **同じストアのアイテムに新しいパッケージをアップロードしてください。** 新しいアイテムとして登録すると拡張 ID が変わり、`ExtensionInstallForcelist`・`AutoLaunchProtocolsFromOrigins`・managed storage の設定をすべて書き換える必要があります。非公開や限定公開から一般公開に切り替える場合も、同じアイテムの公開範囲を変更すれば ID は変わりません。
- **既定のテナントに頼っていた利用者への対応**：以前の版で許可テナントの既定値を使っていた場合、1.1.0 に更新すると、テナントを自分で登録していない利用者では何も起きなくなります（更新時に設定画面は開きます）。社内展開では、更新の前に managed storage で `allowedHosts` を配布しておくのが確実です（[5.3](#53-拡張機能の設定をポリシーで配布managed-storage)）。
