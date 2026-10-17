import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', 'extension');
const read = (f) => readFileSync(join(root, f), 'utf8');
const manifest = JSON.parse(read('manifest.json'));
const locales = readdirSync(join(root, '_locales'), { withFileTypes: true }).filter((d) => d.isDirectory()).map((d) => d.name);
const catalogs = Object.fromEntries(locales.map((l) => [l, JSON.parse(read(`_locales/${l}/messages.json`))]));
const en = catalogs.en;

test('英語（既定）と日本語のカタログがある', () => {
  assert.equal(manifest.default_locale, 'en');
  assert.deepEqual(locales.sort(), ['en', 'ja']);
});

test('すべての言語でキーとプレースホルダーが一致する', () => {
  for (const [lang, cat] of Object.entries(catalogs)) {
    assert.deepEqual(Object.keys(cat).sort(), Object.keys(en).sort(), `keys differ: ${lang}`);
    for (const [key, entry] of Object.entries(cat)) {
      assert.ok(entry.message, `${lang}.${key} is empty`);
      const used = [...entry.message.matchAll(/\$([a-z_]+)\$/gi)].map((m) => m[1].toLowerCase()).sort();
      const declared = Object.keys(entry.placeholders || {}).map((k) => k.toLowerCase()).sort();
      assert.deepEqual(used, declared, `${lang}.${key} placeholders`);
      assert.deepEqual(declared, Object.keys(en[key].placeholders || {}).sort(), `${lang}.${key} placeholders differ from en`);
    }
  }
});

test('HTML・JS・manifest から参照するキーがすべて存在する', () => {
  const keys = new Set();
  for (const f of ['launch.html', 'options.html']) {
    for (const m of read(f).matchAll(/data-i18n="([^"]+)"/g)) keys.add(m[1]);
  }
  for (const f of ['launch.js', 'options.js']) {
    const src = read(f);
    for (const m of src.matchAll(/\bt\('([^']+)'/g)) keys.add(m[1]);
    for (const m of src.matchAll(/:\s*'(err[A-Za-z]+)'/g)) keys.add(m[1]); // REASONS の値
    for (const m of src.matchAll(/reasonText\([^,]+,\s*'([^']+)'\)/g)) keys.add(m[1]);
  }
  for (const v of [manifest.name, manifest.short_name, manifest.description]) {
    const m = /^__MSG_(\w+)__$/.exec(v);
    assert.ok(m, `manifest field not localized: ${v}`);
    keys.add(m[1]);
  }
  assert.ok(keys.size > 30);
  for (const k of keys) assert.ok(en[k], `missing message: ${k}`);
});

test('ストアの文字数制限に収まる', () => {
  for (const [lang, cat] of Object.entries(catalogs)) {
    assert.ok(cat.extName.message.length <= 75, `${lang} name`);
    assert.ok(cat.extShortName.message.length <= 12, `${lang} short_name`);
    assert.ok(cat.extDescription.message.length <= 132, `${lang} description`);
  }
});

test('ページの HTML に日本語の固定文言が残っていない', () => {
  for (const f of ['launch.html', 'options.html']) {
    assert.doesNotMatch(read(f), /[぀-ヿ一-鿿]/, f);
  }
});
