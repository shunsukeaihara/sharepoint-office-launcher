// assets/icon.svg から拡張機能アイコン（PNG）を Chromium で描画する。
//   npm i -D playwright && npx playwright install chromium && node scripts/render-icons.mjs
// 128px はストアのガイドラインに従い、周囲 16px を透明にして中央 96px に描く。
// 16/32px は線が細く潰れないよう太めの線幅にする。
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const svg = readFileSync(join(root, 'assets', 'icon.svg'), 'utf8');
const SPECS = [
  { size: 16, pad: 0, stroke: 14 },
  { size: 32, pad: 0, stroke: 12 },
  { size: 48, pad: 0, stroke: 10 },
  { size: 128, pad: 16, stroke: 10 },
];

const browser = await chromium.launch();
const page = await browser.newPage();
for (const { size, pad, stroke } of SPECS) {
  const inner = size - pad * 2;
  const body = svg.replace('stroke-width="10"', `stroke-width="${stroke}"`)
    .replace('<svg ', `<svg width="${inner}" height="${inner}" `);
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(`<style>html,body{margin:0;background:transparent}svg{display:block;margin:${pad}px}</style>${body}`);
  await page.screenshot({ path: join(root, 'extension', 'icons', `icon${size}.png`), omitBackground: true });
}
await browser.close();
