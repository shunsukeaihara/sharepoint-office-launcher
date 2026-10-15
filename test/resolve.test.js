import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseDocAspxLink } from '../extension/lib/convert.js';
import { fileApiUrl, resolveServerRelativeUrl } from '../extension/lib/resolve.js';

const doc = parseDocAspxLink(
  'https://contoso.sharepoint.com/:p:/r/sites/pj%20sample/_layouts/15/Doc.aspx?sourcedoc=%7BAA3B8D18-237C-446C-9D6E-F98E2E9A8FBA%7D&file=a.pptx',
  ['contoso.sharepoint.com'],
);

const reply = (status, body, extra = {}) => async () =>
  new Response(typeof body === 'string' ? body : JSON.stringify(body), { status, ...extra });

test('API の URL は同じテナント・同じサイトの GetFileById', () => {
  assert.equal(
    fileApiUrl(doc),
    "https://contoso.sharepoint.com/sites/pj%20sample/_api/web/GetFileById('aa3b8d18-237c-446c-9d6e-f98e2e9a8fba')?$select=ServerRelativeUrl",
  );
});

test('Cookie 付き・リダイレクト非追従・JSON で問い合わせる', async () => {
  let seen;
  const sru = await resolveServerRelativeUrl(doc, {
    fetchImpl: async (url, init) => {
      seen = { url, init };
      return new Response(JSON.stringify({ ServerRelativeUrl: '/sites/pj sample/Docs/a.pptx' }), { status: 200 });
    },
  });
  assert.equal(sru, '/sites/pj sample/Docs/a.pptx');
  assert.equal(seen.url, fileApiUrl(doc));
  assert.equal(seen.init.credentials, 'include');
  assert.equal(seen.init.redirect, 'manual');
  assert.match(seen.init.headers.Accept, /application\/json/);
});

test('odata=verbose 形式の応答にも対応', async () => {
  const sru = await resolveServerRelativeUrl(doc, { fetchImpl: reply(200, { d: { ServerRelativeUrl: '/sites/x/a.pptx' } }) });
  assert.equal(sru, '/sites/x/a.pptx');
});

const errors = [
  ['未サインイン（302）', reply(302, '', { headers: { Location: 'https://login.microsoftonline.com/' } }), 'not-signed-in'],
  ['未サインイン（opaqueredirect）', async () => ({ type: 'opaqueredirect', status: 0, ok: false }), 'not-signed-in'],
  ['401', reply(401, ''), 'not-signed-in'],
  ['権限なし（403）', reply(403, ''), 'forbidden'],
  ['見つからない（404）', reply(404, ''), 'not-found'],
  ['サーバーエラー（500）', reply(500, ''), 'http-error'],
  ['JSON でない', reply(200, '<html>'), 'bad-response'],
  ['ServerRelativeUrl なし', reply(200, { Name: 'a.pptx' }), 'bad-response'],
  ['通信エラー', async () => { throw new TypeError('Failed to fetch'); }, 'network'],
  ['タイムアウト', async () => { throw new DOMException('timeout', 'TimeoutError'); }, 'timeout'],
];
for (const [name, fetchImpl, code] of errors) {
  test(`解決エラー: ${name}`, async () => {
    await assert.rejects(resolveServerRelativeUrl(doc, { fetchImpl }), (e) => e.code === code);
  });
}
