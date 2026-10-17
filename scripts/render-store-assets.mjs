// ストア掲載用画像を store-assets/<言語>/ に生成する（en / ja）。
//   npm i -D playwright && npx playwright install chromium && node scripts/render-store-assets.mjs
// 実際の拡張機能ページ（launch.html / options.html）を各言語の UI で Chromium に読み込んで撮影し、
// 説明文付きのスクリーンショット（1280x800）とプロモタイル（440x280 / 1400x560）を作る。
import { mkdirSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';
import { launchWithExtension } from '../test/e2e/launch-chromium.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const ext = join(root, 'extension');
const iconSvg = readFileSync(join(root, 'assets', 'icon.svg'), 'utf8');
const iconInverted = iconSvg.replace('fill="#B7472A"', 'fill="#FFFFFF"').replace('stroke="#FFFFFF"', 'stroke="#B7472A"');
const FONT = `system-ui, -apple-system, "Segoe UI", "Hiragino Sans", "Yu Gothic UI", sans-serif`;
const sampleLink = (fileName) =>
  `https://contoso.sharepoint.com/:p:/r/sites/Sales/Shared%20Documents/${encodeURIComponent(fileName)}?d=w0c1e&csf=1&web=1&e=Ab12Cd`;

const COPY = {
  ja: {
    sample: sampleLink('営業提案_2026Q4.pptx'),
    s1Title: 'SharePoint のリンクを<br>Office アプリで<br>直接開く',
    s1Body: 'Slack などから届いた PowerPoint・Excel・Word のリンクを検出し、ブラウザを経由せずデスクトップアプリで開きます。',
    flow: ['Slack のリンク', 'ブラウザ', 'Office アプリ'],
    s2Title: '許可したテナント<br>だけに反応します',
    s2Points: ['対象の SharePoint テナントを自分で登録', '編集モード／読み取り専用を選択', '組織ポリシーで一括配布にも対応', '外部へのデータ送信なし'],
    tagline: 'SharePoint のリンクを<br>Office アプリで直接開く',
    support: 'PowerPoint ・ Excel ・ Word ／ Chrome ・ Edge 対応',
  },
  en: {
    sample: sampleLink('Sales Proposal 2026Q4.pptx'),
    s1Title: 'Open SharePoint links<br>directly in<br>Office apps',
    s1Body: 'Detects PowerPoint, Excel, and Word links shared in Slack and other apps, and opens them in the desktop app instead of the browser.',
    flow: ['Slack link', 'Browser', 'Office app'],
    s2Title: 'Works only on<br>tenants you allow',
    s2Points: ['Register your own SharePoint tenants', 'Choose edit or read-only mode', 'Deploy settings via enterprise policy', 'No data is sent to any third party'],
    tagline: 'Open SharePoint links<br>directly in Office apps',
    support: 'PowerPoint · Excel · Word  /  Chrome · Edge',
  },
};

// 拡張機能ページを指定言語で開いて撮影する
async function captureExtensionPages(lang, sample) {
  const work = mkdtempSync(join(tmpdir(), 'sp-office-assets-'));
  const { ctx, close } = await launchWithExtension({ ext, userDataDir: join(work, 'profile'), lang });
  try {
    const isOurs = (w) => w.url().startsWith('chrome-extension://') && w.url().endsWith('/background.js');
    const sw = ctx.serviceWorkers().find(isOurs) ?? await ctx.waitForEvent('serviceworker', { predicate: isOurs });
    await sw.evaluate(() => chrome.storage.sync.set({ allowedHosts: ['contoso.sharepoint.com', 'contoso-my.sharepoint.com'] }));
    const extId = new URL(sw.url()).host;

    const capture = async (path, { width, height, ready, prepare }) => {
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
    };

    const launchShot = await capture(`launch.html?src=${encodeURIComponent(sample)}`, {
      width: 760, height: 560,
      ready: () => !document.getElementById('file').hidden,
      prepare: (p) => p.evaluate(() => { document.getElementById('details').open = true; }),
    });
    const optionsShot = await capture('options.html', {
      width: 760, height: 900,
      ready: () => document.getElementById('ext-origin').textContent !== '',
    });
    return { launchShot, optionsShot };
  } finally {
    await close();
    rmSync(work, { recursive: true, force: true });
  }
}

const render = await chromium.launch();
async function compose(out, file, width, height, lang, html) {
  const page = await render.newPage({ viewport: { width, height }, deviceScaleFactor: 1 });
  await page.setContent(`<!doctype html><html lang="${lang}"><meta charset="utf-8"><style>
    *{box-sizing:border-box} html,body{margin:0;width:${width}px;height:${height}px;overflow:hidden}
    body{font-family:${FONT};color:#1c1917;background:#f5f0eb}
    .icon svg{display:block;width:100%;height:100%}
    .shot{border-radius:14px;box-shadow:0 18px 50px rgba(60,30,20,.18),0 2px 6px rgba(60,30,20,.08);border:1px solid #e3d9d2;display:block}
  </style>${html}`);
  await page.waitForTimeout(200);
  await page.screenshot({ path: join(out, file) });
  await page.close();
}

for (const [lang, c] of Object.entries(COPY)) {
  const out = join(root, 'store-assets', lang);
  mkdirSync(out, { recursive: true });
  const { launchShot, optionsShot } = await captureExtensionPages(lang, c.sample);

  const icon = (size) => `<div class="icon" style="width:${size}px;height:${size}px">${iconSvg}</div>`;
  const pill = (text, accent) =>
    `<span style="padding:6px 12px;border-radius:999px;${accent ? 'background:#b7472a;color:#fff' : 'background:#fff;border:1px solid #e3d9d2'}">${text}</span>`;
  const flow = `<div style="display:flex;align-items:center;gap:10px;font-size:15px;color:#57534e;flex-wrap:wrap">
    ${c.flow.map((f, i) => pill(f, i === c.flow.length - 1)).join('<span>→</span>')}</div>`;

  await compose(out, 'screenshot-1-launch.png', 1280, 800, lang, `
    <div style="display:flex;height:100%;align-items:center;padding:0 64px;gap:56px">
      <div style="flex:0 0 400px">
        ${icon(64)}
        <h1 style="font-size:36px;line-height:1.3;margin:28px 0 18px;word-break:keep-all">${c.s1Title}</h1>
        <p style="font-size:18px;line-height:1.6;color:#57534e;margin:0 0 28px">${c.s1Body}</p>
        ${flow}
      </div>
      <img class="shot" src="${launchShot}" style="width:700px">
    </div>`);

  await compose(out, 'screenshot-2-options.png', 1280, 800, lang, `
    <div style="display:flex;height:100%;align-items:center;padding:0 64px;gap:56px">
      <div style="flex:0 0 400px">
        <h1 style="font-size:36px;line-height:1.3;margin:0 0 18px">${c.s2Title}</h1>
        <ul style="font-size:18px;line-height:1.8;color:#57534e;margin:0;padding-left:1.2em">
          ${c.s2Points.map((p) => `<li>${p}</li>`).join('')}
        </ul>
      </div>
      <div style="width:700px;height:620px;overflow:hidden;border-radius:14px" class="shot">
        <img src="${optionsShot}" style="width:700px;display:block">
      </div>
    </div>`);

  await compose(out, 'promo-small-440x280.png', 440, 280, lang, `
    <div style="height:100%;padding:36px;display:flex;flex-direction:column;justify-content:center;background:linear-gradient(135deg,#b7472a,#8f3320);color:#fff">
      <div style="display:flex;align-items:center;gap:14px">
        <div style="width:56px;height:56px;border-radius:12px;background:#fff">${iconInverted}</div>
        <div style="font-size:24px;font-weight:700;line-height:1.2">SharePoint<br>Office Launcher</div>
      </div>
      <p style="font-size:17px;line-height:1.5;margin:22px 0 0;opacity:.95">${c.tagline}</p>
    </div>`);

  await compose(out, 'promo-marquee-1400x560.png', 1400, 560, lang, `
    <div style="height:100%;display:flex;align-items:center;gap:64px;padding:0 80px;background:linear-gradient(135deg,#b7472a,#8f3320);color:#fff">
      <div style="flex:1">
        <div style="display:flex;align-items:center;gap:18px">
          <div style="width:72px;height:72px;border-radius:16px;background:#fff">${iconInverted}</div>
          <div style="font-size:30px;font-weight:700">SharePoint Office Launcher</div>
        </div>
        <h1 style="font-size:44px;line-height:1.3;margin:32px 0 16px">${c.tagline}</h1>
        <p style="font-size:20px;opacity:.9;margin:0">${c.support}</p>
      </div>
      <img class="shot" src="${launchShot}" style="width:560px;border:none">
    </div>`);
  console.log('written to', out);
}
await render.close();
