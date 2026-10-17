// 拡張機能を読み込んだ Playwright の Chromium を、指定した UI 言語で起動する。
// macOS の Chromium は --lang を無視して OS の言語設定（AppleLanguages）を使うため、
// macOS では Chromium を直接起動して「-AppleLanguages (xx)」を渡し、CDP で接続する
// （Playwright の起動 API はハイフン 1 つで始まる引数を受け付けないため）。
import { spawn } from 'node:child_process';
import { join } from 'node:path';
import { chromium } from 'playwright';

/**
 * @param {{ext: string, userDataDir: string, lang: string, args?: string[]}} opts
 * @returns {Promise<{ctx: import('playwright').BrowserContext, close: () => Promise<void>}>}
 */
export async function launchWithExtension({ ext, userDataDir, lang, args = [] }) {
  const common = [
    `--disable-extensions-except=${ext}`,
    `--load-extension=${ext}`,
    `--lang=${lang}`,
    ...args,
  ];
  if (process.platform !== 'darwin') {
    const ctx = await chromium.launchPersistentContext(userDataDir, { channel: 'chromium', headless: true, args: common });
    return { ctx, close: () => ctx.close() };
  }

  const proc = spawn(chromium.executablePath(), [
    '--headless=new',
    '--remote-debugging-port=0',
    `--user-data-dir=${join(userDataDir)}`,
    '--no-first-run',
    '--no-default-browser-check',
    ...common,
    // 値の「(xx)」は Chromium からは開く URL として扱われるため、他に URL を渡さない（headless は 1 つまで）
    '-AppleLanguages', `(${lang})`,
  ], { stdio: ['ignore', 'ignore', 'pipe'] });
  const wsEndpoint = await new Promise((resolve, reject) => {
    let buf = '';
    const timer = setTimeout(() => reject(new Error(`Chromium did not start:\n${buf}`)), 30000);
    proc.stderr.on('data', (d) => {
      buf += d;
      const m = /DevTools listening on (ws:\/\/\S+)/.exec(buf);
      if (m) {
        clearTimeout(timer);
        resolve(m[1]);
      }
    });
    proc.on('exit', (code) => reject(new Error(`Chromium exited (${code}):\n${buf}`)));
  });
  const browser = await chromium.connectOverCDP(wsEndpoint);
  const ctx = browser.contexts()[0];
  return {
    ctx,
    close: async () => {
      await browser.close().catch(() => {});
      proc.kill();
      await new Promise((r) => (proc.exitCode !== null ? r() : proc.once('exit', r)));
    },
  };
}
