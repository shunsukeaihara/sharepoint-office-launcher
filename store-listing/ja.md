# ストア掲載情報（日本語）

各ブロックを Chrome Web Store / Edge Add-ons のダッシュボードに貼り付けてください。ストアの説明欄ではマークダウンが表示されないため、説明文はプレーンテキストです。

## 説明文

```text
Slack やメールから届いた SharePoint の PowerPoint・Excel・Word のリンクを、ブラウザ版ではなく Office デスクトップアプリで直接開く拡張機能です。

SharePoint の共有リンクをクリックすると、通常はブラウザで Office for the web が開きます。デスクトップアプリで編集したい場合は「アプリで開く」を選び直す必要がありました。この拡張機能を入れると、登録したテナントのリンクを開いたときに自動でデスクトップアプリが起動します。

■ 主な機能
・SharePoint の PowerPoint（/:p:/r/）、Excel（/:x:/r/）、Word（/:w:/r/）形式のリンクを検出し、PowerPoint / Excel / Word で直接開きます
・「…/_layouts/15/Doc.aspx?sourcedoc=…」形式のリンクにも対応しています（ファイルの場所を SharePoint に確認してから開きます）
・編集モードと読み取り専用モードを選べます
・起動ページから「読み取り専用で開く」「ブラウザで開く」「ファイル URL をコピー」も選べます
・自分で登録したテナントのリンクにだけ反応します。登録していないサイトのリンクには一切関与しません
・OneDrive（<テナント名>-my.sharepoint.com）にも対応しています
・日本語と英語に対応しています（ブラウザの表示言語に合わせて切り替わります）

■ 使い方
1. インストールすると設定画面が開きます
2. 「許可するテナント」に、自社の SharePoint のテナント名（例：contoso、または contoso.sharepoint.com）を入力して保存します
3. Slack などで SharePoint の Office ファイルのリンクをクリックすると、デスクトップアプリが起動します
※ 初回は、ブラウザが「アプリを開きますか？」と確認する場合があります。「開く」を選んでください。
※ Doc.aspx 形式のリンクを開くには、ブラウザで SharePoint にサインインしている必要があります。

■ 組織での利用
・管理者は、対象テナントなどの設定をポリシー（managed storage）でまとめて配布できます
・ブラウザのポリシー AutoLaunchProtocolsFromOrigins を使うと、起動時の確認ダイアログを省略できます

■ プライバシー
・データの収集は一切行いません。開発者や第三者のサーバーとは通信しません
・Doc.aspx 形式のリンクを開くときだけ、リンクと同じ（登録済みの）SharePoint テナントにファイルの場所を問い合わせます。取得した情報は保存も外部送信もしません
・ページの内容や Cookie は読み取りません
・保存するのは設定値（対象テナント・開き方など）のみです

■ 動作要件
・Google Chrome または Microsoft Edge（バージョン 110 以降）
・Microsoft Office デスクトップアプリ（Windows / macOS）がインストールされ、対象テナントのアカウントでサインイン済みであること

■ 対応していないリンク
・「リンクを知っているユーザー」などで作られた /:p:/g/ や /:p:/s/ 形式の共有リンクは、ファイルの場所を特定できないため変換できません。これらは通常どおりブラウザで開きます

本拡張機能は Microsoft とは関係のない非公式のツールです。Microsoft、SharePoint、PowerPoint、Excel、Word、Office は Microsoft グループの商標です。
```

## 単一用途

```text
ユーザーが登録した SharePoint テナントの PowerPoint / Excel / Word ファイルへのリンクを、ブラウザで開く代わりに Office デスクトップアプリで直接開くこと。
```

## 権限の説明：webNavigation

```text
Slack などから既定のブラウザに渡された SharePoint の Office ファイルリンクを、ページの読み込みが始まる前に検出するために使用します。対象は次の形式です。
・https://<テナント>.sharepoint.com/:p:/r/…（:x: / :w: も同様）
・https://<テナント>.sharepoint.com/:p:/r/…/_layouts/15/Doc.aspx?sourcedoc=…

chrome.webNavigation.onBeforeNavigate を、ホスト名が「.sharepoint.com」で終わり、パスが「/:」で始まる URL だけに絞って監視しています。対象はタブのトップレベルの遷移のみで、サブフレームやプリレンダリングは無視します。検出した URL がユーザーの登録したテナントのものであれば、そのタブを拡張機能内の起動ページに切り替え、Office デスクトップアプリ（PowerPoint / Excel / Word）で開きます。

閲覧履歴を保存・送信することはなく、上記の条件に合わない遷移には一切関与しません。
```

## 権限の説明：webRequest

```text
リダイレクト（HTTP 302 など）を経由して SharePoint のファイルリンクに到達した場合も検出するために使用します。webNavigation はリダイレクトの途中の URL を通知しないため、その分を webRequest で補っています。

chrome.webRequest.onBeforeRequest を、種類が main_frame で URL が「https://*.sharepoint.com/:*」に一致するリクエストだけに絞って監視しています。使い方は観測のみです。blocking、リクエストのブロック、書き換え、ヘッダーの読み取りや変更、通信内容の取得はいずれも行いません。webNavigation と同じ判定処理に URL を渡すだけで、二重に処理しないよう重複を除いています。
```

## 権限の説明：storage

```text
拡張機能の設定と、動作に必要な一時情報を保存するために使用します。

・chrome.storage.sync：ユーザーが設定画面で登録した内容（対象とする SharePoint テナントのホスト名、開き方〔編集／読み取り専用〕、自動起動の有無など）を保存します。
・chrome.storage.session：起動ページで「ブラウザで開く」が選ばれたとき、同じリンクを再び横取りしないための一時的な除外情報を、最長 60 秒だけ保持します。ブラウザを閉じると消去されます。
・chrome.storage.managed：組織の管理者がポリシーで配布した設定（対象テナントなど）を読み取ります。

保存するのは設定値のみで、閲覧履歴、ファイルの内容、SharePoint から取得したファイルの場所、個人情報は保存しません。外部への送信も行いません。
```

## 権限の説明：ホスト権限

```text
ホスト権限「https://*.sharepoint.com/*」は、次の 2 つの目的に必要です。

1. webRequest で、SharePoint 宛てのトップレベルのリクエストを観測するため（Chrome では、webRequest のイベントを受け取るにはそのホストへの権限が必要です）。

2. 「/:p:/r/…/_layouts/15/Doc.aspx?sourcedoc={ファイルID}」形式のリンクについて、実ファイルの場所を取得するため。この形式の URL にはファイルのフォルダパスが含まれていないため、起動ページから、リンクと同じテナント・同じサイトの SharePoint REST API（/_api/web/GetFileById('<ファイルID>')?$select=ServerRelativeUrl）に、ブラウザのサインイン状態のまま GET で問い合わせます。取得するのはファイルのサーバー上のパスだけで、Office アプリの起動にのみ使い、保存や SharePoint 以外への送信はしません。SharePoint 上のデータを変更することもありません。

SharePoint のテナントは組織ごとに「<テナント名>.sharepoint.com」という別々のホスト名になっていて、公開時点では利用者のテナントを特定できないため、*.sharepoint.com を指定しています。実際に検出・問い合わせの対象とするのは、ユーザーが設定画面で登録したテナント（または管理者がポリシーで指定したテナント）だけです。登録のないホストについては何もせず、通常どおりブラウザで開きます。

この権限を使ってページにスクリプトを挿入したり、ページの内容や Cookie を読み取ったり変更したりすることはありません（コンテンツスクリプトは使用していません）。SharePoint 以外のサイトへの権限は要求していません。
```

## リモートコードの使用

```text
いいえ。すべてのコードはパッケージに含まれており、外部のスクリプトを読み込んだり eval で実行したりはしません。SharePoint API への問い合わせで受け取るのは JSON のデータ（ファイルのパス）のみで、コードとして実行することはありません。
```

## 「プライバシーへの取り組み」タブ

- データ使用の申告：収集するデータの項目はすべて「なし」。SharePoint から取得するファイルの場所は、開発者に送信もせず端末に保存もしないため、「収集」にはあたりません。
- プライバシーポリシーの URL：https://github.com/shunsukeaihara/sharepoint-office-launcher/blob/main/PRIVACY.md
- サポート／ホームページの URL：https://github.com/shunsukeaihara/sharepoint-office-launcher

## 画像（`store-assets/ja/`）

| ファイル | 用途 |
|---|---|
| `screenshot-1-launch.png`（1280×800） | スクリーンショット（起動画面） |
| `screenshot-2-options.png`（1280×800） | スクリーンショット（設定画面） |
| `promo-small-440x280.png` | 小プロモタイル |
| `promo-marquee-1400x560.png` | マーキー（Chrome Web Store）／大プロモタイル（Edge Add-ons） |
