// Playwright の Chromium に拡張機能を読み込んで動作を確認する E2E。
// *.sharepoint.com は名前解決させないため、実テナントへは一切通信しない。
// Office URI による実アプリ起動はヘッドレスでは確認できないため、手動検証（README 11 章）で行う。
//   npm i -D playwright && npx playwright install chromium && npm run test:e2e
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ext = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'extension');
const profile = mkdtempSync(join(tmpdir(), 'sp-office-launcher-'));
const ctx = await chromium.launchPersistentContext(profile, {
  channel: 'chromium', headless: true,
  args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`,
         '--host-resolver-rules=MAP *.sharepoint.com ~NOTFOUND'],
});
let [sw] = ctx.serviceWorkers();
if (!sw) sw = await ctx.waitForEvent('serviceworker');
await sw.evaluate(() => chrome.storage.sync.set({ allowedHosts: ['contoso.sharepoint.com'] })); // テスト用テナントを登録（SW の初期化完了待ちも兼ねる）
const errs = [];
const page = await ctx.newPage();
page.on('pageerror', (e) => errs.push(String(e)));
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const go = async (u, expectLaunch = true) => {
  await page.goto(u, { waitUntil: 'commit', timeout: 5000 }).catch(() => {});
  const target = `launch.html?src=${encodeURIComponent(u)}`;
  if (expectLaunch) await page.waitForURL((x) => x.href.endsWith(target), { timeout: 5000 }).catch(() => {});
  else await wait(2500);
  const url = page.url();
  // 期待した src の launch.html でなければ（直前ページのまま等）横取りされていない扱い
  return url.startsWith('chrome-extension://') && !url.endsWith(target) ? `stale:${url}` : url;
};
const isLaunch = (u) => u.startsWith('chrome-extension://') && u.includes('/launch.html');
const results = [];
const check = (name, cond, info = '') => { results.push(cond); console.log(`${cond ? 'PASS' : 'FAIL'} ${name} ${info}`); };

const ok = 'https://contoso.sharepoint.com/:p:/r/sites/Test/Shared%20Documents/%E6%8F%90%E6%A1%88.pptx?d=w1&csf=1&web=1';
check('T1 許可テナントの pptx を横取り', isLaunch(await go(ok)));
const text = (await page.locator('main').innerText({ timeout: 5000 })).replace(/\n+/g, ' / ');
check('T1 launch.html の表示', text.includes('PowerPoint で開いています') && text.includes('提案.pptx'), text.slice(0, 160));
check('T1 起動済みを履歴に記録', (await page.evaluate(() => history.state))?.launched === true);
check('T2 xlsx', isLaunch(await go('https://contoso.sharepoint.com/:x:/r/sites/a/Book.xlsx?web=1')));
check('T2 docx', isLaunch(await go('https://contoso.sharepoint.com/:w:/r/sites/a/Doc.docx?web=1')));
check('T4 許可外テナントは素通り', !isLaunch(await go('https://other.sharepoint.com/:p:/r/sites/a/b.pptx', false)));
check('T5 /g/ 共有リンクは素通り', !isLaunch(await go('https://contoso.sharepoint.com/:p:/g/sites/a/EaBc', false)));
check('拡張子不一致は素通り', !isLaunch(await go('https://contoso.sharepoint.com/:p:/r/sites/a/b.xlsx', false)));
await go(ok);
await page.waitForSelector('#open-browser:visible');
await wait(6000); // 二重処理防止の時間窓を過ぎてから押し、バイパス経路そのものを検証する
await page.evaluate(() => document.getElementById('open-browser').click()).catch(() => {});
await wait(2500);
check('T6 ブラウザで開く（再横取りしない）', !isLaunch(page.url()), page.url().slice(0, 80));
check('T6 バイパスは消費済み', JSON.stringify(await sw.evaluate(() => chrome.storage.session.get('bypass'))) === '{"bypass":{}}');
await page.goBack({ timeout: 5000 }).catch(() => {});
await wait(1500);
const backText = await page.locator('#title').innerText({ timeout: 5000 }).catch(() => '');
check('T8 戻るでは自動起動しない', isLaunch(page.url()) && backText === 'PowerPoint で開く', backText);
await page.goto('about:blank');
await wait(40000); // SW のアイドル停止（約 30 秒）後の再起動でも横取りできるか
check('SW アイドル後も横取り', isLaunch(await go('https://contoso.sharepoint.com/:x:/r/sites/a/Cold.xlsx')));
check('ページエラーなし', errs.length === 0, errs.join(' | '));
await ctx.close();
rmSync(profile, { recursive: true, force: true });
process.exit(results.every(Boolean) ? 0 : 1);
