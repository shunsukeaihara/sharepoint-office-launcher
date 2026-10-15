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
  const officeUri = `${def.scheme}:${mode}|u|${fileUrl}`;
  return { ok: true, type, app: def.app, scheme: def.scheme, fileName, fileUrl, officeUri };
}

function fail(reason) {
  return { ok: false, reason };
}
