// Doc.aspx 形式のリンクについて、ファイル ID から実ファイルの場所を SharePoint REST API で取得する。
// 通信先はリンクと同じ（許可済みの）SharePoint テナントのみ。ブラウザのサインイン状態（Cookie）を使う。

export class ResolveError extends Error {
  constructor(code) {
    super(code);
    this.name = 'ResolveError';
    this.code = code;
  }
}

/** @param {{host:string, webPath:string, fileId:string}} doc parseDocAspxLink の結果 */
export function fileApiUrl(doc) {
  return `https://${doc.host}${doc.webPath}/_api/web/GetFileById('${doc.fileId}')?$select=ServerRelativeUrl`;
}

/**
 * @returns {Promise<string>} ServerRelativeUrl（例: "/sites/xxx/Shared Documents/a.pptx"）
 * @throws {ResolveError} code: not-signed-in | forbidden | not-found | http-error | timeout | network | bad-response
 */
export async function resolveServerRelativeUrl(doc, { fetchImpl = fetch, timeoutMs = 10000 } = {}) {
  let res;
  try {
    res = await fetchImpl(fileApiUrl(doc), {
      method: 'GET',
      credentials: 'include',
      redirect: 'manual', // 未サインイン時のログイン画面へのリダイレクトは追わない
      cache: 'no-store',
      headers: { Accept: 'application/json;odata=nometadata' },
      signal: AbortSignal.timeout(timeoutMs),
    });
  } catch (e) {
    throw new ResolveError(e?.name === 'TimeoutError' ? 'timeout' : 'network');
  }

  if (res.type === 'opaqueredirect' || (res.status >= 300 && res.status < 400) || res.status === 401) {
    throw new ResolveError('not-signed-in');
  }
  if (res.status === 403) throw new ResolveError('forbidden');
  if (res.status === 404) throw new ResolveError('not-found');
  if (!res.ok) throw new ResolveError('http-error');

  let body;
  try {
    body = await res.json();
  } catch {
    throw new ResolveError('bad-response');
  }
  const sru = body?.ServerRelativeUrl ?? body?.d?.ServerRelativeUrl;
  if (typeof sru !== 'string' || sru === '') throw new ResolveError('bad-response');
  return sru;
}
