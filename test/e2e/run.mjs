// Playwright の Chromium に拡張機能を読み込んで動作を確認する E2E。
// contoso.sharepoint.com はローカルの偽 SharePoint（HTTPS・自己署名証明書）に向け、
// それ以外の *.sharepoint.com は名前解決させないため、実テナントへは一切通信しない。
// Office URI による実アプリ起動はヘッドレスでは確認できないため、手動検証（README 11 章）で行う。
//   npm i -D playwright && npx playwright install chromium && npm run test:e2e
// 前提: openssl コマンド（自己署名証明書の生成に使用）
import { execFileSync } from 'node:child_process';
import { createServer } from 'node:https';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const ext = join(dirname(fileURLToPath(import.meta.url)), '..', '..', 'extension');
const work = mkdtempSync(join(tmpdir(), 'sp-office-launcher-'));
const HOST = 'contoso.sharepoint.com';
const ORIGIN = `https://${HOST}`;
const GUID_OK = 'aa3b8d18-237c-446c-9d6e-f98e2e9a8fba';
const GUID_FORBIDDEN = 'bb3b8d18-237c-446c-9d6e-f98e2e9a8fbb';
const SRU = '/sites/pj-sample/Shared Documents/営業/サンプル資料 v2.pptx';

// --- 偽 SharePoint ---
execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-days', '1', '-subj', `/CN=${HOST}`,
  '-keyout', join(work, 'key.pem'), '-out', join(work, 'cert.pem')], { stdio: 'ignore' });
const apiCalls = [];
const server = createServer({ key: readFileSync(join(work, 'key.pem')), cert: readFileSync(join(work, 'cert.pem')) }, (req, res) => {
  const m = /\/_api\/web\/GetFileById\('([0-9a-f-]+)'\)/.exec(decodeURIComponent(req.url));
  if (m) {
    apiCalls.push({ url: req.url, cookie: req.headers.cookie || '', accept: req.headers.accept || '' });
    if (!/FedAuth=ok/.test(req.headers.cookie || '')) {
      res.writeHead(302, { Location: 'https://login.microsoftonline.com/' });
      return res.end();
    }
    if (m[1] === GUID_FORBIDDEN) {
      res.writeHead(403);
      return res.end();
    }
    if (m[1] !== GUID_OK) {
      res.writeHead(404);
      return res.end();
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ ServerRelativeUrl: SRU }));
  }
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end('<title>SP</title>SP PAGE');
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const port = server.address().port;

const ctx = await chromium.launchPersistentContext(join(work, 'profile'), {
  channel: 'chromium',
  headless: true,
  args: [
    `--disable-extensions-except=${ext}`,
    `--load-extension=${ext}`,
    `--host-resolver-rules=MAP ${HOST} 127.0.0.1:${port}, MAP *.sharepoint.com ~NOTFOUND`,
    '--ignore-certificate-errors',
  ],
});
const signIn = () => ctx.addCookies([
  { name: 'FedAuth', value: 'ok', domain: HOST, path: '/', secure: true, httpOnly: true, sameSite: 'Lax' },
]);
let [sw] = ctx.serviceWorkers();
if (!sw) sw = await ctx.waitForEvent('serviceworker');
// テスト用テナントを登録（SW の初期化完了待ちも兼ねる）
await sw.evaluate((h) => chrome.storage.sync.set({ allowedHosts: [h] }), HOST);
await signIn();

const errs = [];
const page = await ctx.newPage();
page.on('pageerror', (e) => errs.push(String(e)));
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const go = async (u, expectLaunch = true) => {
  await page.goto(u, { waitUntil: 'commit', timeout: 5000 }).catch(() => {});
  const target = `launch.html?src=${encodeURIComponent(u)}`;
  // 拡張機能の tabs.update による遷移は page.url() への反映が遅れることがあるため、自前でポーリングする。
  if (expectLaunch) {
    for (let i = 0; i < 80 && !page.url().endsWith(target); i++) await wait(100);
  } else await wait(2500);
  const url = page.url();
  // 期待した src の launch.html でなければ（直前ページのまま等）横取りされていない扱い
  return url.startsWith('chrome-extension://') && !url.endsWith(target) ? `stale:${url}` : url;
};
const isLaunch = (u) => u.startsWith('chrome-extension://') && u.includes('/launch.html');
const text = (id) => page.locator(`#${id}`).innerText({ timeout: 5000 }).catch(() => '');
const visible = (id) => page.locator(`#${id}`).isVisible();
const results = [];
const check = (name, cond, info = '') => { results.push(cond); console.log(`${cond ? 'PASS' : 'FAIL'} ${name} ${info}`); };

const ok = `${ORIGIN}/:p:/r/sites/Test/Shared%20Documents/%E6%8F%90%E6%A1%88.pptx?d=w1&csf=1&web=1`;
check('T1 許可テナントの pptx を横取り', isLaunch(await go(ok)));
const mainText = (await page.locator('main').innerText({ timeout: 5000 })).replace(/\n+/g, ' / ');
check('T1 launch.html の表示', mainText.includes('PowerPoint で開いています') && mainText.includes('提案.pptx'), mainText.slice(0, 120));
check('T1 起動済みを履歴に記録', (await page.evaluate(() => history.state))?.launched === true);
check('T2 xlsx', isLaunch(await go(`${ORIGIN}/:x:/r/sites/a/Book.xlsx?web=1`)));
check('T2 docx', isLaunch(await go(`${ORIGIN}/:w:/r/sites/a/Doc.docx?web=1`)));
check('T4 許可外テナントは素通り', !isLaunch(await go('https://other.sharepoint.com/:p:/r/sites/a/b.pptx', false)));
check('T5 /g/ 共有リンクは素通り', !isLaunch(await go(`${ORIGIN}/:p:/g/sites/a/EaBc`, false)));
check('拡張子不一致は素通り', !isLaunch(await go(`${ORIGIN}/:p:/r/sites/a/b.xlsx`, false)));

// --- Doc.aspx 形式 ---
const docLink = (guid) =>
  `${ORIGIN}/:p:/r/sites/pj-sample/_layouts/15/Doc.aspx?sourcedoc=%7B${guid.toUpperCase()}%7D&file=%E3%82%B5%E3%83%B3%E3%83%97%E3%83%AB%E8%B3%87%E6%96%99%20v2.pptx&action=edit`;
apiCalls.length = 0;
check('D1 Doc.aspx 形式を横取り', isLaunch(await go(docLink(GUID_OK))));
await page.waitForFunction(() => document.getElementById('file-url').textContent !== '', null, { timeout: 5000 }).catch(() => {});
const resolved = await page.evaluate(() => document.getElementById('file-url').textContent);
check('D1 ファイル ID から実ファイル URL を解決', resolved === `${ORIGIN}${SRU.split('/').map(encodeURIComponent).join('/')}`, resolved);
check('D1 タイトル', (await text('title')).startsWith('PowerPoint で開いています'));
check('D1 API 呼び出しにサインイン Cookie が付く', apiCalls.length === 1 && /FedAuth=ok/.test(apiCalls[0].cookie), JSON.stringify(apiCalls[0] || {}));
check('D1 API は同じサイトの GetFileById・JSON 指定', apiCalls[0]?.url.startsWith('/sites/pj-sample/_api/web/GetFileById(') && /application\/json/.test(apiCalls[0]?.accept));

check('D2 権限なし（403）', isLaunch(await go(docLink(GUID_FORBIDDEN))));
await page.waitForSelector('#error:visible', { timeout: 5000 }).catch(() => {});
check('D2 アクセス権なしのメッセージ', (await text('error')).includes('アクセス権がない'), await text('error'));
check('D2 「ブラウザで開く」だけを表示', (await visible('open-browser')) && !(await visible('open-edit')) && !(await visible('copy')));

await ctx.clearCookies();
check('D3 未サインイン', isLaunch(await go(docLink(GUID_OK).replace('action=edit', 'action=view'))));
await page.waitForSelector('#error:visible', { timeout: 5000 }).catch(() => {});
check('D3 サインインを促すメッセージ', (await text('error')).includes('サインインしていない'), await text('error'));
await wait(6000); // 二重処理防止の時間窓を過ぎてから押す
const d3src = docLink(GUID_OK).replace('action=edit', 'action=view');
await page.evaluate(() => document.getElementById('open-browser').click()).catch(() => {});
await wait(2500);
check('D3 「ブラウザで開く」で元のリンクを開く（再横取りしない）', page.url() === d3src, page.url().slice(0, 80));
await signIn();

check('D4 許可外テナントの Doc.aspx は素通り', !isLaunch(await go(docLink(GUID_OK).replace('contoso', 'other'), false)));

// --- ループ防止 ---
await go(ok);
await page.waitForSelector('#open-browser:visible');
await wait(6000); // 二重処理防止の時間窓を過ぎてから押し、バイパス経路そのものを検証する
await page.evaluate(() => document.getElementById('open-browser').click()).catch(() => {});
await wait(2500);
check('T6 ブラウザで開く（再横取りしない）', page.url() === ok, page.url().slice(0, 80));
check('T6 バイパスは消費済み', JSON.stringify(await sw.evaluate(() => chrome.storage.session.get('bypass'))) === '{"bypass":{}}');
await page.goBack({ timeout: 5000 }).catch(() => {});
await wait(1500);
const backTitle = await text('title');
check('T8 戻るでは自動起動しない', isLaunch(page.url()) && backTitle === 'PowerPoint で開く', backTitle);

await page.goto('about:blank');
await wait(40000); // SW のアイドル停止（約 30 秒）後の再起動でも横取りできるか
const coldUrl = await go(`${ORIGIN}/:x:/r/sites/a/Cold.xlsx`);
check('SW アイドル後も横取り', isLaunch(coldUrl), coldUrl.slice(0, 100));
check('ページエラーなし', errs.length === 0, errs.join(' | '));

await ctx.close();
server.close();
rmSync(work, { recursive: true, force: true });
console.log(`${results.filter(Boolean).length}/${results.length} passed`);
process.exit(results.every(Boolean) ? 0 : 1);
