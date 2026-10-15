import { test } from 'node:test';
import assert from 'node:assert/strict';
import { linkFromServerRelativeUrl, parseDocAspxLink } from '../extension/lib/convert.js';

const ALLOWED = ['contoso.sharepoint.com'];
const BASE = 'https://contoso.sharepoint.com';
const GUID = 'AA3B8D18-237C-446C-9D6E-F98E2E9A8FBA';
const docLink = (path = '/sites/pj-sample', type = 'p', file = '%E3%82%B5%E3%83%B3%E3%83%97%E3%83%AB%E8%B3%87%E6%96%99%20v2.pptx') =>
  `${BASE}/:${type}:/r${path}/_layouts/15/Doc.aspx?sourcedoc=%7B${GUID}%7D&file=${file}&action=edit&mobileredirect=true`;

test('Doc.aspx 形式のリンクからサイトパスとファイル ID を取り出す', () => {
  const d = parseDocAspxLink(docLink(), ALLOWED);
  assert.equal(d.ok, true);
  assert.equal(d.type, 'p');
  assert.equal(d.app, 'PowerPoint');
  assert.equal(d.host, 'contoso.sharepoint.com');
  assert.equal(d.webPath, '/sites/pj-sample');
  assert.equal(d.webPathDecoded, '/sites/pj-sample');
  assert.equal(d.fileId, GUID.toLowerCase());
  assert.equal(d.fileName, 'サンプル資料 v2.pptx');
});

test('ルートサイト・サブサイト・大文字小文字・波括弧なし・file なし', () => {
  assert.equal(parseDocAspxLink(docLink(''), ALLOWED).webPath, '');
  assert.equal(parseDocAspxLink(docLink('/teams/a/sub%20site'), ALLOWED).webPathDecoded, '/teams/a/sub site');
  const upper = `${BASE}/:X:/r/sites/a/_LAYOUTS/15/DOC.ASPX?sourcedoc=${GUID}`;
  const d = parseDocAspxLink(upper, ALLOWED);
  assert.equal(d.ok, true);
  assert.equal(d.type, 'x');
  assert.equal(d.fileName, '');
});

const rejects = [
  ['許可外テナント', docLink().replace('contoso', 'other'), 'host-not-allowed'],
  ['/:x:/r/ でない Doc.aspx', `${BASE}/sites/a/_layouts/15/Doc.aspx?sourcedoc=%7B${GUID}%7D`, 'not-doc-aspx'],
  ['sourcedoc なし', `${BASE}/:p:/r/sites/a/_layouts/15/Doc.aspx?file=a.pptx`, 'bad-sourcedoc'],
  ['sourcedoc が GUID でない', `${BASE}/:p:/r/sites/a/_layouts/15/Doc.aspx?sourcedoc=%7Bx')%20or%201%7D`, 'bad-sourcedoc'],
  ['file の拡張子が種別と不一致', docLink('/sites/a', 'p', 'a.xlsx'), 'extension-mismatch'],
  ['file にパイプ', docLink('/sites/a', 'p', 'a%7Cb.pptx'), 'forbidden-char'],
  ['サイトパスにエンコード済みスラッシュ', docLink('/sites/a%2F..%2Fb'), 'forbidden-char'],
  ['サイトパスに不正なエンコード', docLink('/sites/%E6%8F'), 'bad-encoding'],
  ['別の _layouts ページ', `${BASE}/:p:/r/sites/a/_layouts/15/WopiFrame.aspx?sourcedoc=%7B${GUID}%7D`, 'not-doc-aspx'],
];
for (const [name, url, reason] of rejects) {
  test(`Doc.aspx 拒否: ${name}`, () => {
    const d = parseDocAspxLink(url, ALLOWED);
    assert.equal(d.ok, false);
    assert.equal(d.reason, reason);
  });
}

test('ServerRelativeUrl から Office URI を作る（日本語・空白をエンコード）', () => {
  const d = parseDocAspxLink(docLink(), ALLOWED);
  const r = linkFromServerRelativeUrl(d, '/sites/pj-sample/Shared Documents/営業/サンプル資料 v2.pptx', 'ofe');
  assert.equal(r.ok, true);
  assert.equal(r.fileName, 'サンプル資料 v2.pptx');
  assert.equal(
    r.fileUrl,
    `${BASE}/sites/pj-sample/Shared%20Documents/%E5%96%B6%E6%A5%AD/%E3%82%B5%E3%83%B3%E3%83%97%E3%83%AB%E8%B3%87%E6%96%99%20v2.pptx`,
  );
  assert.equal(r.officeUri, `ms-powerpoint:ofe|u|${r.fileUrl}`);
});

test('ServerRelativeUrl のサイト判定は大文字小文字を区別しない', () => {
  const d = parseDocAspxLink(docLink('/sites/PJ-Sample'), ALLOWED);
  assert.equal(linkFromServerRelativeUrl(d, '/sites/pj-sample/Docs/a.pptx').ok, true);
});

const badSru = [
  ['サイト外', '/sites/other/Docs/a.pptx', 'outside-site'],
  ['前方一致だけのサイト外', '/sites/pj-sample-evil/Docs/a.pptx', 'outside-site'],
  ['ドットセグメントでサイト外へ', '/sites/pj-sample/../other/a.pptx', 'bad-segment'],
  ['プロトコル相対', '//evil.example/a.pptx', 'bad-response'],
  ['相対パス', 'sites/pj-sample/a.pptx', 'bad-response'],
  ['文字列でない', 42, 'bad-response'],
  ['拡張子が種別と不一致', '/sites/pj-sample/Docs/a.docx', 'extension-mismatch'],
  ['パイプを含む', '/sites/pj-sample/Docs/a|ofe|u|x.pptx', 'forbidden-char'],
];
for (const [name, sru, reason] of badSru) {
  test(`ServerRelativeUrl 拒否: ${name}`, () => {
    const d = parseDocAspxLink(docLink(), ALLOWED);
    const r = linkFromServerRelativeUrl(d, sru);
    assert.equal(r.ok, false);
    assert.equal(r.reason, reason);
  });
}
