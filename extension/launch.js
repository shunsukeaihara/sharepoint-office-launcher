import { parseSharePointLink } from './lib/convert.js';
import { getConfig } from './lib/config.js';
import { addBypass, navKey } from './lib/guard.js';

const BYPASS_KEY = 'bypass';
const CLOSE_DELAY_MS = 4000;
const $ = (id) => document.getElementById(id);

const REASONS = {
  'host-not-allowed': '許可されていないテナントのリンクです。',
  'not-r-link': '対応していない形式のリンクです（/:p:/r/ 形式のみ対応）。',
  'extension-mismatch': 'リンク種別とファイル拡張子が一致しません。',
};

main().catch((e) => showError(`起動に失敗しました: ${e?.message || e}`));

async function main() {
  const src = new URLSearchParams(location.search).get('src') || '';
  $('src').textContent = src;

  // background.js での判定を信用せず、ここでも設定に基づいて再検証する。
  const config = await getConfig();
  const r = parseSharePointLink(src, config.allowedHosts, config.openMode);
  if (!r.ok) {
    $('details').hidden = false;
    showError(REASONS[r.reason] || `このリンクは処理できません（${r.reason}）。`);
    return;
  }

  const editUri = parseSharePointLink(src, config.allowedHosts, 'ofe').officeUri;
  const viewUri = parseSharePointLink(src, config.allowedHosts, 'ofv').officeUri;

  document.title = `${r.app} で開いています`;
  $('title').textContent = `${r.app} で開いています…`;
  $('file').textContent = r.fileName;
  $('file').hidden = false;
  $('file-url').textContent = r.fileUrl;
  $('open-edit').textContent = `${r.app} で開く（編集）`;
  $('actions').hidden = false;
  $('details').hidden = false;

  $('open-edit').addEventListener('click', () => launch(editUri));
  $('open-view').addEventListener('click', () => launch(viewUri));
  $('open-browser').addEventListener('click', () => openInBrowser(src));
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

async function openInBrowser(src) {
  const tab = await chrome.tabs.getCurrent();
  if (tab?.id != null) {
    const stored = await chrome.storage.session.get(BYPASS_KEY);
    const next = addBypass(stored[BYPASS_KEY], navKey(tab.id, src), Date.now());
    await chrome.storage.session.set({ [BYPASS_KEY]: next });
  }
  location.href = src;
}

function showError(text) {
  $('title').textContent = 'Office アプリで開けませんでした';
  $('message').hidden = true;
  $('error').textContent = text;
  $('error').hidden = false;
}
