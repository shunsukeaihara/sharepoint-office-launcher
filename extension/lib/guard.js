// ループ・二重処理防止のための純粋ロジック。

/** 同一タブ・同一 URL のイベントを一定時間内に 1 回だけ通す。 */
export class RecentGuard {
  constructor(windowMs = 5000, now = () => Date.now()) {
    this.windowMs = windowMs;
    this.now = now;
    this.seen = new Map();
  }

  /** 初回なら true を返して記録する。windowMs 以内の再来は false。 */
  claim(key) {
    const t = this.now();
    for (const [k, at] of this.seen) {
      if (t - at >= this.windowMs) this.seen.delete(k);
    }
    if (this.seen.has(key)) return false;
    this.seen.set(key, t);
    return true;
  }
}

export function navKey(tabId, url) {
  return `${tabId}|${url}`;
}

/**
 * 「ブラウザで開く」用のバイパス表から key を 1 回だけ取り出す。
 * 期限切れのエントリも同時に掃除する。
 * @returns {{hit:boolean, next:Record<string, number>}}
 */
export function takeBypass(map, key, now) {
  const next = {};
  let hit = false;
  for (const [k, exp] of Object.entries(map || {})) {
    if (typeof exp !== 'number' || exp <= now) continue;
    if (k === key) {
      hit = true;
      continue;
    }
    next[k] = exp;
  }
  return { hit, next };
}

export function addBypass(map, key, now, ttlMs = 60000) {
  const { next } = takeBypass(map, key, now);
  next[key] = now + ttlMs;
  return next;
}
