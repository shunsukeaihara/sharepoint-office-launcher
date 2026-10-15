// ストア掲載用画像を store-assets/ に生成する。
//   npm i -D playwright && npx playwright install chromium && node scripts/render-store-assets.mjs
// 実際の拡張機能ページ（launch.html / options.html）を Chromium に読み込んで撮影し、
// 説明文付きのスクリーンショット（1280x800）とプロモタイル（440x280 / 1400x560）を作る。
import { mkdirSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const ext = join(root, 'extension');
const out = join(root, 'store-assets');
mkdirSync(out, { recursive: true });

const SAMPLE = 'https://contoso.sharepoint.com/:p:/r/sites/Sales/Shared%20Documents/%E5%96%B6%E6%A5%AD%E6%8F%90%E6%A1%88_2026Q4.pptx?d=w0c1e&csf=1&web=1&e=Ab12Cd';
const iconSvg = readFileSync(join(root, 'assets', 'icon.svg'), 'utf8');
const FONT = `system-ui, -apple-system, "Segoe UI", "Hiragino Sans", "Yu Gothic UI", sans-serif`;

const profile = mkdtempSync(join(tmpdir(), 'sp-office-assets-'));
const ctx = await chromium.launchPersistentContext(profile, {
  channel: 'chromium',
  headless: true,
  args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`],
});
let [sw] = ctx.serviceWorkers();
if (!sw) sw = await ctx.waitForEvent('serviceworker');
await sw.evaluate(() => chrome.storage.sync.set({ allowedHosts: ['contoso.sharepoint.com', 'contoso-my.sharepoint.com'] }));
const extId = new URL(sw.url()).host;

// 拡張機能ページの実画面を撮影（1.5 倍で描画して文字を大きく見せる）
async function capturePage(path, { width, height, ready, prepare }) {
  const page = await ctx.newPage();
  await page.setViewportSize({ width, height });
  await page.goto(`chrome-extension://${extId}/${path}`);
  await page.waitForLoadState('load');
  await page.waitForFunction(ready, null, { timeout: 10000 }); // 設定の読み込み（初回は managed storage で約 1 秒）を待つ
  if (prepare) await prepare(page);
  await page.waitForTimeout(300);
  const buf = await page.screenshot();
  await page.close();
  return `data:image/png;base64,${buf.toString('base64')}`;
}

const launchShot = await capturePage(`launch.html?src=${encodeURIComponent(SAMPLE)}`, {
  width: 760, height: 560,
  ready: () => !document.getElementById('file').hidden,
  prepare: (p) => p.evaluate(() => { document.getElementById('details').open = true; }),
});
const optionsShot = await capturePage('options.html', {
  width: 760, height: 900,
  ready: () => document.getElementById('ext-origin').textContent !== '',
});

// 説明文付きの合成画像を描画
const render = await chromium.launch();
async function compose(file, width, height, html) {
  const page = await render.newPage({ viewport: { width, height }, deviceScaleFactor: 1 });
  await page.setContent(`<!doctype html><meta charset="utf-8"><style>
    *{box-sizing:border-box} html,body{margin:0;width:${width}px;height:${height}px;overflow:hidden}
    body{font-family:${FONT};color:#1c1917;background:#f5f0eb}
    .icon svg{display:block;width:100%;height:100%}
    .shot{border-radius:14px;box-shadow:0 18px 50px rgba(60,30,20,.18),0 2px 6px rgba(60,30,20,.08);border:1px solid #e3d9d2;display:block}
  </style>${html}`);
  await page.waitForTimeout(200);
  await page.screenshot({ path: join(out, file) });
  await page.close();
}

const icon = (size) => `<div class="icon" style="width:${size}px;height:${size}px">${iconSvg}</div>`;
const flow = `
  <div style="display:flex;align-items:center;gap:10px;font-size:15px;color:#57534e;flex-wrap:wrap">
    <span style="padding:6px 12px;border-radius:999px;background:#fff;border:1px solid #e3d9d2">Slack のリンク</span><span>→</span>
    <span style="padding:6px 12px;border-radius:999px;background:#fff;border:1px solid #e3d9d2">ブラウザ</span><span>→</span>
    <span style="padding:6px 12px;border-radius:999px;background:#b7472a;color:#fff">Office アプリ</span>
  </div>`;

await compose('screenshot-1-launch.png', 1280, 800, `
  <div style="display:flex;height:100%;align-items:center;padding:0 64px;gap:56px">
    <div style="flex:0 0 400px">
      ${icon(64)}
      <h1 style="font-size:36px;line-height:1.35;margin:28px 0 18px;word-break:keep-all">SharePoint のリンクを<br>Office アプリで<br>直接開く</h1>
      <p style="font-size:18px;line-height:1.7;color:#57534e;margin:0 0 28px">Slack などから届いた PowerPoint・Excel・Word のリンクを検出し、ブラウザを経由せずデスクトップアプリで開きます。</p>
      ${flow}
    </div>
    <img class="shot" src="${launchShot}" style="width:700px">
  </div>`);

await compose('screenshot-2-options.png', 1280, 800, `
  <div style="display:flex;height:100%;align-items:center;padding:0 64px;gap:56px">
    <div style="flex:0 0 400px">
      <h1 style="font-size:36px;line-height:1.35;margin:0 0 18px">許可したテナント<br>だけに反応します</h1>
      <ul style="font-size:18px;line-height:1.8;color:#57534e;margin:0;padding-left:1.2em">
        <li>対象の SharePoint テナントを自分で登録</li>
        <li>編集モード／読み取り専用を選択</li>
        <li>組織ポリシーで一括配布にも対応</li>
        <li>外部へのデータ送信なし</li>
      </ul>
    </div>
    <div style="width:700px;height:620px;overflow:hidden;border-radius:14px" class="shot">
      <img src="${optionsShot}" style="width:700px;display:block">
    </div>
  </div>`);

await compose('promo-small-440x280.png', 440, 280, `
  <div style="height:100%;padding:36px;display:flex;flex-direction:column;justify-content:center;background:linear-gradient(135deg,#b7472a,#8f3320);color:#fff">
    <div style="display:flex;align-items:center;gap:14px">
      <div style="width:56px;height:56px;border-radius:12px;background:#fff;padding:0">${iconSvg.replace('fill="#B7472A"', 'fill="#FFFFFF"').replace('stroke="#FFFFFF"', 'stroke="#B7472A"')}</div>
      <div style="font-size:24px;font-weight:700;line-height:1.2">SharePoint<br>Office Launcher</div>
    </div>
    <p style="font-size:17px;line-height:1.6;margin:22px 0 0;opacity:.95">SharePoint のリンクを<br>Office アプリで直接開く</p>
  </div>`);

await compose('promo-marquee-1400x560.png', 1400, 560, `
  <div style="height:100%;display:flex;align-items:center;gap:64px;padding:0 80px;background:linear-gradient(135deg,#b7472a,#8f3320);color:#fff">
    <div style="flex:1">
      <div style="display:flex;align-items:center;gap:18px">
        <div style="width:72px;height:72px;border-radius:16px;background:#fff">${iconSvg.replace('fill="#B7472A"', 'fill="#FFFFFF"').replace('stroke="#FFFFFF"', 'stroke="#B7472A"')}</div>
        <div style="font-size:30px;font-weight:700">SharePoint Office Launcher</div>
      </div>
      <h1 style="font-size:44px;line-height:1.35;margin:32px 0 16px">SharePoint のリンクを<br>Office アプリで直接開く</h1>
      <p style="font-size:20px;opacity:.9;margin:0">PowerPoint ・ Excel ・ Word ／ Chrome ・ Edge 対応</p>
    </div>
    <img class="shot" src="${launchShot}" style="width:560px;border:none">
  </div>`);

await render.close();
await ctx.close();
rmSync(profile, { recursive: true, force: true });
console.log('written to', out);
