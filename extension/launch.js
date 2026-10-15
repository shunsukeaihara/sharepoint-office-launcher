import { linkFromServerRelativeUrl, parseDocAspxLink, parseSharePointLink, toOfficeUri } from './lib/convert.js';
import { getConfig } from './lib/config.js';
import { addBypass, navKey } from './lib/guard.js';
import { resolveServerRelativeUrl } from './lib/resolve.js';

const BYPASS_KEY = 'bypass';
const CLOSE_DELAY_MS = 4000;
const $ = (id) => document.getElementById(id);

const REASONS = {
  'host-not-allowed': '許可されていないテナントのリンクです。',
  'not-r-link': '対応していない形式のリンクです。',
  'not-doc-aspx': '対応していない形式のリンクです。',
  'extension-mismatch': 'リンク種別とファイル拡張子が一致しません。',
  'outside-site': 'ファイルの場所がリンク先のサイト外を指しているため開けません。',
  'not-signed-in': 'SharePoint にサインインしていないため、ファイルの場所を確認できませんでした。ブラウザで開いてサインインしてから、もう一度リンクを開いてください。',
  forbidden: 'このファイルへのアクセス権がないため、ファイルの場所を確認できませんでした。',
  'not-found': 'ファイルが見つかりませんでした（移動または削除された可能性があります）。',
  timeout: 'SharePoint からの応答がタイムアウトしました。',
  network: 'SharePoint に接続できませんでした。',
  'http-error': 'SharePoint からエラーが返されました。',
  'bad-response': 'SharePoint から想定外の応答が返されました。',
};

main().catch((e) => showError(`起動に失敗しました: ${e?.message || e}`));

async function main() {
  const src = new URLSearchParams(location.search).get('src') || '';
  $('src').textContent = src;

  // background.js での判定を信用せず、ここでも設定に基づいて再検証する。
  const config = await getConfig();
  let r = parseSharePointLink(src, config.allowedHosts, config.openMode);

  if (!r.ok) {
    const doc = parseDocAspxLink(src, config.allowedHosts);
    if (!doc.ok) {
      $('details').hidden = false;
      showError(REASONS[r.reason] || `このリンクは処理できません（${r.reason}）。`);
      return;
    }
    // Doc.aspx 形式: ファイル ID から実ファイルの場所を SharePoint に問い合わせる。
    $('title').textContent = `${doc.app} のファイルを確認しています…`;
    if (doc.fileName) {
      $('file').textContent = doc.fileName;
      $('file').hidden = false;
    }
    $('message').textContent = 'SharePoint にファイルの場所を問い合わせています。';
    try {
      const sru = await resolveServerRelativeUrl(doc);
      r = linkFromServerRelativeUrl(doc, sru, config.openMode);
    } catch (e) {
      r = { ok: false, reason: e?.code || 'network' };
    }
    if (!r.ok) {
      $('details').hidden = false;
      showError(REASONS[r.reason] || `ファイルの場所を確認できませんでした（${r.reason}）。`, { src });
      return;
    }
  }

  const editUri = toOfficeUri(r.scheme, 'ofe', r.fileUrl);
  const viewUri = toOfficeUri(r.scheme, 'ofv', r.fileUrl);

  document.title = `${r.app} で開いています`;
  $('title').textContent = `${r.app} で開いています…`;
  $('message').textContent = 'ブラウザに「アプリを開きますか？」と表示された場合は「開く」を選んでください。';
  $('file').textContent = r.fileName;
  $('file').hidden = false;
  $('file-url').textContent = r.fileUrl;
  $('open-edit').textContent = `${r.app} で開く（編集）`;
  $('actions').hidden = false;
  $('details').hidden = false;

  $('open-edit').addEventListener('click', () => launch(editUri));
  $('open-view').addEventListener('click', () => launch(viewUri));
  bindOpenInBrowser(src);
  $('copy').addEventListener('click', async () => {
    await navigator.clipboard.writeText(r.fileUrl);
    $('copy').textContent = 'コピーしました';
  });

  // 戻るボタン等でこのページに戻った場合は自動起動しない（再起動ループ防止）。
  const alreadyLaunched = history.state && history.state.launched === true;
  if (!config.autoLaunch || alreadyLaunched) {
    $('title').textContent = `${r.app} で開く`;
    $('message').textContent = '下のボタンから開き方を選んでください。';
    return;
  }

  history.replaceState({ launched: true }, '');
  launch(r.officeUri);

  if (config.closeTabAfterLaunch) {
    setTimeout(async () => {
      const tab = await chrome.tabs.getCurrent();
      if (tab?.id != null) chrome.tabs.remove(tab.id);
    }, CLOSE_DELAY_MS);
  }
}

function launch(officeUri) {
  // 外部プロトコルへのトップレベル遷移。ページ自体は残る。
  location.href = officeUri;
}

function bindOpenInBrowser(src) {
  $('open-browser').addEventListener('click', () => openInBrowser(src), { once: true });
}

async function openInBrowser(src) {
  const tab = await chrome.tabs.getCurrent();
  if (tab?.id != null) {
    const stored = await chrome.storage.session.get(BYPASS_KEY);
    const next = addBypass(stored[BYPASS_KEY], navKey(tab.id, src), Date.now());
    await chrome.storage.session.set({ [BYPASS_KEY]: next });
  }
  location.href = src;
}

/** @param {{src?: string}} [opts] src を渡すと、許可済みリンクとして「ブラウザで開く」を表示する */
function showError(text, { src } = {}) {
  $('title').textContent = 'Office アプリで開けませんでした';
  $('message').hidden = true;
  $('error').textContent = text;
  $('error').hidden = false;
  if (src) {
    for (const id of ['open-edit', 'open-view', 'copy']) $(id).hidden = true;
    $('open-browser').classList.add('primary');
    $('actions').hidden = false;
    bindOpenInBrowser(src);
  }
}
