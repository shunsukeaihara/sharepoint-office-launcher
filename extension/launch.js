import { linkFromServerRelativeUrl, parseDocAspxLink, parseSharePointLink, toOfficeUri } from './lib/convert.js';
import { getConfig } from './lib/config.js';
import { addBypass, navKey } from './lib/guard.js';
import { localizePage, t } from './lib/i18n.js';
import { resolveServerRelativeUrl } from './lib/resolve.js';

const BYPASS_KEY = 'bypass';
const CLOSE_DELAY_MS = 4000;
const $ = (id) => document.getElementById(id);

// 失敗理由コード → メッセージキー（_locales/*/messages.json）
const REASONS = {
  'host-not-allowed': 'errHostNotAllowed',
  'not-r-link': 'errUnsupportedLink',
  'not-doc-aspx': 'errUnsupportedLink',
  'extension-mismatch': 'errExtensionMismatch',
  'outside-site': 'errOutsideSite',
  'not-signed-in': 'errNotSignedIn',
  forbidden: 'errForbidden',
  'not-found': 'errNotFound',
  timeout: 'errTimeout',
  network: 'errNetwork',
  'http-error': 'errHttp',
  'bad-response': 'errBadResponse',
};
const reasonText = (reason, fallbackKey) => (REASONS[reason] ? t(REASONS[reason]) : t(fallbackKey, [reason]));

localizePage();
main().catch((e) => showError(t('errStartup', [String(e?.message || e)])));

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
      showError(reasonText(r.reason, 'errUnknownLink'));
      return;
    }
    // Doc.aspx 形式: ファイル ID から実ファイルの場所を SharePoint に問い合わせる。
    $('title').textContent = t('launchResolving', [doc.app]);
    if (doc.fileName) {
      $('file').textContent = doc.fileName;
      $('file').hidden = false;
    }
    $('message').textContent = t('launchResolvingHint');
    try {
      const sru = await resolveServerRelativeUrl(doc);
      r = linkFromServerRelativeUrl(doc, sru, config.openMode);
    } catch (e) {
      r = { ok: false, reason: e?.code || 'network' };
    }
    if (!r.ok) {
      $('details').hidden = false;
      showError(reasonText(r.reason, 'errUnknownResolve'), { src });
      return;
    }
  }

  const editUri = toOfficeUri(r.scheme, 'ofe', r.fileUrl);
  const viewUri = toOfficeUri(r.scheme, 'ofv', r.fileUrl);

  document.title = t('launchDocTitle', [r.app]);
  $('title').textContent = t('launchOpening', [r.app]);
  $('message').textContent = t('launchDialogHint');
  $('file').textContent = r.fileName;
  $('file').hidden = false;
  $('file-url').textContent = r.fileUrl;
  $('open-edit').textContent = t('btnOpenEdit', [r.app]);
  $('actions').hidden = false;
  $('details').hidden = false;

  $('open-edit').addEventListener('click', () => launch(editUri));
  $('open-view').addEventListener('click', () => launch(viewUri));
  bindOpenInBrowser(src);
  $('copy').addEventListener('click', async () => {
    await navigator.clipboard.writeText(r.fileUrl);
    $('copy').textContent = t('copied');
  });

  // 戻るボタン等でこのページに戻った場合は自動起動しない（再起動ループ防止）。
  const alreadyLaunched = history.state && history.state.launched === true;
  if (!config.autoLaunch || alreadyLaunched) {
    $('title').textContent = t('launchChoose', [r.app]);
    $('message').textContent = t('launchChooseHint');
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
  $('title').textContent = t('errorTitle');
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
