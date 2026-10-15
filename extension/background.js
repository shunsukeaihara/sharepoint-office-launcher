import { parseDocAspxLink, parseSharePointLink } from './lib/convert.js';
import { getConfig } from './lib/config.js';
import { RecentGuard, navKey, takeBypass } from './lib/guard.js';

const BYPASS_KEY = 'bypass';
const guard = new RecentGuard(5000);

// webNavigation は直接のナビゲーションを、webRequest はリダイレクタ経由（302）で
// 到達した場合も拾う。両方が発火しても RecentGuard で 1 回に畳む。
chrome.webNavigation.onBeforeNavigate.addListener(
  (d) => {
    if (d.frameType && d.frameType !== 'outermost_frame') return;
    if (d.documentLifecycle === 'prerender') return;
    handle(d.tabId, d.frameId, d.url);
  },
  { url: [{ hostSuffix: '.sharepoint.com', pathPrefix: '/:' }] },
);

chrome.webRequest.onBeforeRequest.addListener(
  (d) => {
    if (d.documentLifecycle === 'prerender') return;
    handle(d.tabId, d.frameId, d.url);
  },
  { urls: ['https://*.sharepoint.com/:*'], types: ['main_frame'] },
);

async function handle(tabId, frameId, url) {
  try {
    if (typeof tabId !== 'number' || tabId < 0 || frameId !== 0) return;

    const key = navKey(tabId, url);
    if (!guard.claim(key)) return;
    if (await consumeBypass(key)) return;

    const config = await getConfig();
    if (!config.enabled) return;

    // /:x:/r/<ファイルパス> 形式、または /:x:/r/…/_layouts/15/Doc.aspx?sourcedoc=… 形式のみ対象。
    // 後者のファイル場所の解決は launch.html で行う。
    const direct = parseSharePointLink(url, config.allowedHosts, config.openMode);
    if (!direct.ok && !parseDocAspxLink(url, config.allowedHosts).ok) return;

    const launchUrl = `${chrome.runtime.getURL('launch.html')}?src=${encodeURIComponent(url)}`;
    await chrome.tabs.update(tabId, { url: launchUrl });
  } catch (e) {
    console.warn('[sp-office-launcher] failed to handle navigation', e);
  }
}

async function consumeBypass(key) {
  const stored = await chrome.storage.session.get(BYPASS_KEY);
  const { hit, next } = takeBypass(stored[BYPASS_KEY], key, Date.now());
  await chrome.storage.session.set({ [BYPASS_KEY]: next });
  return hit;
}

// 新規インストール時、および更新後に許可テナントが未設定の場合は設定画面を開いて登録を促す。
chrome.runtime.onInstalled.addListener(async (details) => {
  if (details.reason !== 'install' && details.reason !== 'update') return;
  const config = await getConfig();
  if (config.allowedHosts.length === 0) chrome.runtime.openOptionsPage();
});
