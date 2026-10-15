import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseSharePointLink, normalizeHost, normalizeHosts } from '../extension/lib/convert.js';

const ALLOWED = ['contoso.sharepoint.com', 'contoso-my.sharepoint.com'];
const BASE = 'https://contoso.sharepoint.com';

test('pptx: /:p:/r/ リンクを ms-powerpoint の Office URI に変換し、クエリを捨てる', () => {
  const r = parseSharePointLink(
    `${BASE}/:p:/r/sites/Sales/Shared%20Documents/%E6%8F%90%E6%A1%88.pptx?d=w123&csf=1&web=1&e=AbCd`,
    ALLOWED,
  );
  assert.equal(r.ok, true);
  assert.equal(r.app, 'PowerPoint');
  assert.equal(r.fileName, '提案.pptx');
  assert.equal(r.fileUrl, `${BASE}/sites/Sales/Shared%20Documents/%E6%8F%90%E6%A1%88.pptx`);
  assert.equal(r.officeUri, `ms-powerpoint:ofe|u|${BASE}/sites/Sales/Shared%20Documents/%E6%8F%90%E6%A1%88.pptx`);
});

test('xlsx / docx と閲覧モード', () => {
  const x = parseSharePointLink(`${BASE}/:x:/r/sites/a/Book.xlsx?web=1`, ALLOWED, 'ofv');
  assert.equal(x.officeUri, `ms-excel:ofv|u|${BASE}/sites/a/Book.xlsx`);
  const w = parseSharePointLink(`${BASE}/:w:/r/sites/a/Doc.DOCX#frag`, ALLOWED);
  assert.equal(w.officeUri, `ms-word:ofe|u|${BASE}/sites/a/Doc.DOCX`);
});

test('ホスト名は大文字小文字を区別しない', () => {
  const r = parseSharePointLink('https://CONTOSO.sharepoint.com/:P:/r/sites/a/b.pptx', ALLOWED);
  assert.equal(r.ok, true);
  assert.equal(r.fileUrl, `${BASE}/sites/a/b.pptx`);
});

test('OneDrive (-my) も許可リストにあれば変換する', () => {
  const r = parseSharePointLink(
    'https://contoso-my.sharepoint.com/:x:/r/personal/user_example_com/Documents/a.xlsx?d=1',
    ALLOWED,
  );
  assert.equal(r.ok, true);
});

const rejects = [
  ['許可外テナント', 'https://evil.sharepoint.com/:p:/r/sites/a/b.pptx', 'host-not-allowed'],
  ['サフィックス偽装', 'https://contoso.sharepoint.com.evil.test/:p:/r/sites/a/b.pptx', 'not-sharepoint'],
  ['http', 'http://contoso.sharepoint.com/:p:/r/sites/a/b.pptx', 'not-https'],
  ['資格情報付き', 'https://u:p@contoso.sharepoint.com/:p:/r/sites/a/b.pptx', 'credentials-in-url'],
  ['ポート指定', 'https://contoso.sharepoint.com:8443/:p:/r/sites/a/b.pptx', 'non-default-port'],
  ['共有リンク /g/', `${BASE}/:p:/g/sites/a/EaBcD`, 'not-r-link'],
  ['共有リンク /s/', `${BASE}/:p:/s/Sales/EaBcD`, 'not-r-link'],
  ['変換済み URL（ループしない）', `${BASE}/sites/a/b.pptx`, 'not-r-link'],
  ['PDF 種別', `${BASE}/:b:/r/sites/a/b.pdf`, 'unsupported-type'],
  ['種別と拡張子の不一致', `${BASE}/:p:/r/sites/a/b.xlsx`, 'extension-mismatch'],
  ['拡張子なし', `${BASE}/:w:/r/sites/a/docx`, 'extension-mismatch'],
  ['パイプ（URI 区切り）', `${BASE}/:p:/r/sites/a|ofe|u|x/b.pptx`, 'forbidden-char'],
  ['エンコード済みパイプ', `${BASE}/:p:/r/sites/a%7Cx/b.pptx`, 'forbidden-char'],
  ['エンコード済みスラッシュ', `${BASE}/:p:/r/sites/a%2F..%2Fb/c.pptx`, 'forbidden-char'],
  ['ドットセグメントで /:p:/r/ の外へ出る', `${BASE}/:p:/r/%2E%2E/%2E%2E/c.pptx`, 'not-r-link'],
  ['不正なパーセントエンコード', `${BASE}/:p:/r/sites/%E6%8F/c.pptx`, 'bad-encoding'],
  ['空セグメント', `${BASE}/:p:/r/sites//c.pptx`, 'bad-segment'],
  ['URL でない', 'not a url', 'invalid-url'],
];

for (const [name, url, reason] of rejects) {
  test(`拒否: ${name}`, () => {
    const r = parseSharePointLink(url, ALLOWED);
    assert.equal(r.ok, false);
    assert.equal(r.reason, reason);
  });
}

test('エンコード済みドットセグメントはブラウザと同じく解決された URL で扱う', () => {
  const r = parseSharePointLink(`${BASE}/:p:/r/sites/%2E%2E/c.pptx`, ALLOWED);
  assert.equal(r.fileUrl, `${BASE}/c.pptx`);
});

test('不正なモードは拒否', () => {
  assert.equal(parseSharePointLink(`${BASE}/:p:/r/a/b.pptx`, ALLOWED, 'x').reason, 'invalid-mode');
});

test('許可リストが空なら何も変換しない', () => {
  assert.equal(parseSharePointLink(`${BASE}/:p:/r/a/b.pptx`, []).reason, 'host-not-allowed');
});

test('normalizeHost', () => {
  assert.equal(normalizeHost('contoso'), 'contoso.sharepoint.com');
  assert.equal(normalizeHost(' https://Contoso-my.sharepoint.com/ '), 'contoso-my.sharepoint.com');
  assert.equal(normalizeHost('*.sharepoint.com'), null);
  assert.equal(normalizeHost('example.com'), null);
  assert.equal(normalizeHost('a.b.sharepoint.com'), null);
  assert.equal(normalizeHost(''), null);
  assert.equal(normalizeHost(42), null);
  assert.deepEqual(normalizeHosts(['contoso', 'contoso.sharepoint.com', 'bad host']), ['contoso.sharepoint.com']);
  assert.deepEqual(normalizeHosts('contoso'), []);
});

test('既定の許可テナントは空（ユーザーが設定するまで何もしない）', async () => {
  const { DEFAULTS } = await import('../extension/lib/config.js');
  assert.deepEqual(DEFAULTS.allowedHosts, []);
});
