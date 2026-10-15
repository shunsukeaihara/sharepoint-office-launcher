import { normalizeHosts, OPEN_MODES } from './convert.js';

export const DEFAULTS = Object.freeze({
  enabled: true,
  allowedHosts: [],
  openMode: 'ofe',
  autoLaunch: true,
  closeTabAfterLaunch: false,
});

/**
 * 管理ポリシー（chrome.storage.managed）> ユーザー設定（chrome.storage.sync）> 既定値 の順で解決する。
 * @returns {Promise<typeof DEFAULTS & {managedKeys: string[]}>}
 */
export async function getConfig() {
  const [sync, managed] = await Promise.all([
    chrome.storage.sync.get(DEFAULTS).catch(() => ({ ...DEFAULTS })),
    chrome.storage.managed.get(null).catch(() => ({})),
  ]);
  const managedKeys = Object.keys(managed || {}).filter((k) => k in DEFAULTS);
  const merged = { ...DEFAULTS, ...sync };
  for (const k of managedKeys) merged[k] = managed[k];

  return {
    enabled: merged.enabled !== false,
    allowedHosts: normalizeHosts(merged.allowedHosts),
    openMode: OPEN_MODES.includes(merged.openMode) ? merged.openMode : DEFAULTS.openMode,
    autoLaunch: merged.autoLaunch !== false,
    closeTabAfterLaunch: merged.closeTabAfterLaunch === true,
    managedKeys,
  };
}
