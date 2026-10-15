// SharePoint の「/:p:/r/...」形式リンクを実ファイル URL と Office URI に変換する純粋関数群。
// chrome.* API に依存しないため Node.js のテストからもそのまま import できる。

export const TYPE_MAP = Object.freeze({
  p: Object.freeze({ scheme: 'ms-powerpoint', app: 'PowerPoint', exts: ['pptx', 'pptm', 'ppt', 'ppsx', 'potx'] }),
  x: Object.freeze({ scheme: 'ms-excel', app: 'Excel', exts: ['xlsx', 'xlsm', 'xlsb', 'xls'] }),
  w: Object.freeze({ scheme: 'ms-word', app: 'Word', exts: ['docx', 'docm', 'doc', 'dotx'] }),
});

export const OPEN_MODES = Object.freeze(['ofe', 'ofv']);

const HOST_RE = /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.sharepoint\.com$/;
const LINK_PATH_RE = /^\/:([a-z]):\/r(\/.+)$/i;
const DOC_ASPX_PATH_RE = /^\/:([a-z]):\/r((?:\/[^/]+)*?)\/_layouts\/15\/doc\.aspx$/i;
const GUID_RE = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i;
// Office URI のコマンド区切り「|」、空白、制御文字、バックスラッシュは URL 中に許可しない。
const FORBIDDEN_RAW_RE = /[|\\\s\u0000-\u001f\u007f]/;

/**
 * 設定値のホスト名（または "dcxxx" のようなテナント名）を正規化する。
 * 不正な値は null。
 */
export function normalizeHost(value) {
  if (typeof value !== 'string') return null;
  let h = value.trim().toLowerCase();
  h = h.replace(/^https:\/\//, '').replace(/\/+$/, '');
  if (h === '') return null;
  if (!h.includes('.')) h = `${h}.sharepoint.com`;
  return HOST_RE.test(h) ? h : null;
}

export function normalizeHosts(values) {
  if (!Array.isArray(values)) return [];
  const out = new Set();
  for (const v of values) {
    const h = normalizeHost(v);
    if (h) out.add(h);
  }
  return [...out];
}

/**
 * @param {string} rawUrl ブラウザに渡された URL
 * @param {string[]} allowedHosts 正規化済みの許可ホスト名
 * @param {'ofe'|'ofv'} [mode] ofe=編集で開く, ofv=閲覧で開く
 * @returns {{ok:true, type:string, app:string, fileName:string, fileUrl:string, officeUri:string} | {ok:false, reason:string}}
 */
export function parseSharePointLink(rawUrl, allowedHosts, mode = 'ofe') {
  if (!OPEN_MODES.includes(mode)) return fail('invalid-mode');

  const checked = checkUrl(rawUrl, allowedHosts);
  if (!checked.ok) return checked;
  const { url, host } = checked;

  const m = LINK_PATH_RE.exec(url.pathname);
  if (!m) return fail('not-r-link');
  const type = m[1].toLowerCase();
  const def = TYPE_MAP[type];
  if (!def) return fail('unsupported-type');

  const path = m[2];
  if (FORBIDDEN_RAW_RE.test(path)) return fail('forbidden-char');

  const segments = path.split('/').slice(1);
  let decoded;
  try {
    decoded = segments.map((s) => decodeURIComponent(s));
  } catch {
    return fail('bad-encoding');
  }
  for (const s of decoded) {
    if (s === '' || s === '.' || s === '..') return fail('bad-segment');
    if (/[/\\|\u0000-\u001f\u007f]/.test(s)) return fail('forbidden-char');
  }

  const fileName = decoded[decoded.length - 1];
  const dot = fileName.lastIndexOf('.');
  const ext = dot > 0 ? fileName.slice(dot + 1).toLowerCase() : '';
  if (!def.exts.includes(ext)) return fail('extension-mismatch');

  // クエリ（?d=...&csf=1&web=1&e=...）とフラグメントは破棄する。
  const fileUrl = `https://${host}${path}`;
  const officeUri = toOfficeUri(def.scheme, mode, fileUrl);
  return { ok: true, type, app: def.app, scheme: def.scheme, fileName, fileUrl, officeUri };
}

export function toOfficeUri(scheme, mode, fileUrl) {
  return `${scheme}:${mode}|u|${fileUrl}`;
}

/**
 * 「/:p:/r/sites/xxx/_layouts/15/Doc.aspx?sourcedoc={GUID}&file=a.pptx」形式のリンクを解析する。
 * この形式はファイルのフォルダパスを含まないため、ファイル ID から実ファイル URL を
 * 別途解決する必要がある（resolve.js）。
 * @returns {{ok:true, type:string, app:string, scheme:string, host:string, webPath:string, webPathDecoded:string, fileId:string, fileName:string} | {ok:false, reason:string}}
 */
export function parseDocAspxLink(rawUrl, allowedHosts) {
  const checked = checkUrl(rawUrl, allowedHosts);
  if (!checked.ok) return checked;
  const { url, host } = checked;

  const m = DOC_ASPX_PATH_RE.exec(url.pathname);
  if (!m) return fail('not-doc-aspx');
  const type = m[1].toLowerCase();
  const def = TYPE_MAP[type];
  if (!def) return fail('unsupported-type');

  // サイトのパス（"" = ルートサイト, "/sites/xxx", "/teams/xxx/sub" など）
  const webPath = m[2] || '';
  if (FORBIDDEN_RAW_RE.test(webPath)) return fail('forbidden-char');
  let webSegments;
  try {
    webSegments = webPath.split('/').slice(1).map((seg) => decodeURIComponent(seg));
  } catch {
    return fail('bad-encoding');
  }
  for (const seg of webSegments) {
    if (seg === '' || seg === '.' || seg === '..') return fail('bad-segment');
    if (/[/\\|\u0000-\u001f\u007f]/.test(seg)) return fail('forbidden-char');
  }
  const webPathDecoded = webSegments.map((seg) => `/${seg}`).join('');

  const sourcedoc = url.searchParams.get('sourcedoc') || '';
  const idMatch = GUID_RE.exec(sourcedoc.trim());
  if (!idMatch) return fail('bad-sourcedoc');
  const fileId = idMatch[1].toLowerCase();

  // file パラメータは表示用。あれば拡張子が種別と一致することを確認する（最終確認は解決後に行う）。
  const fileName = url.searchParams.get('file') || '';
  if (fileName) {
    if (/[/\\|\u0000-\u001f\u007f]/.test(fileName)) return fail('forbidden-char');
    const dot = fileName.lastIndexOf('.');
    const ext = dot > 0 ? fileName.slice(dot + 1).toLowerCase() : '';
    if (!def.exts.includes(ext)) return fail('extension-mismatch');
  }

  return { ok: true, type, app: def.app, scheme: def.scheme, host, webPath, webPathDecoded, fileId, fileName };
}

/**
 * SharePoint API が返した ServerRelativeUrl（デコード済みのパス）から、
 * /:x:/r/ 形式と同じ検証を通した変換結果を作る。
 */
export function linkFromServerRelativeUrl(doc, serverRelativeUrl, mode = 'ofe') {
  if (typeof serverRelativeUrl !== 'string') return fail('bad-response');
  const sru = serverRelativeUrl;
  if (!sru.startsWith('/') || sru.startsWith('//')) return fail('bad-response');
  // URL の正規化でサイト外へ出られないよう、エンコード前にドット・空セグメントを拒否する。
  for (const seg of sru.split('/').slice(1)) {
    if (seg === '' || seg === '.' || seg === '..') return fail('bad-segment');
  }
  // 解決結果は、リンクが指していたサイトの配下でなければならない。
  if (doc.webPathDecoded && !sru.toLowerCase().startsWith(`${doc.webPathDecoded.toLowerCase()}/`)) {
    return fail('outside-site');
  }
  const encoded = sru.split('/').map((seg) => encodeURIComponent(seg)).join('/');
  return parseSharePointLink(`https://${doc.host}/:${doc.type}:/r${encoded}`, [doc.host], mode);
}

function checkUrl(rawUrl, allowedHosts) {
  let url;
  try {
    url = new URL(rawUrl);
  } catch {
    return fail('invalid-url');
  }

  if (url.protocol !== 'https:') return fail('not-https');
  if (url.username || url.password) return fail('credentials-in-url');
  if (url.port !== '') return fail('non-default-port');

  const host = url.hostname.toLowerCase();
  if (!HOST_RE.test(host)) return fail('not-sharepoint');
  if (!Array.isArray(allowedHosts) || !allowedHosts.includes(host)) return fail('host-not-allowed');
  return { ok: true, url, host };
}

function fail(reason) {
  return { ok: false, reason };
}
