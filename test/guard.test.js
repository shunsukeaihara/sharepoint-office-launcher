import { test } from 'node:test';
import assert from 'node:assert/strict';
import { RecentGuard, navKey, takeBypass, addBypass } from '../extension/lib/guard.js';

test('RecentGuard は同一キーを時間窓内で 1 回だけ通す', () => {
  let now = 1000;
  const g = new RecentGuard(5000, () => now);
  const k = navKey(1, 'https://x');
  assert.equal(g.claim(k), true);
  assert.equal(g.claim(k), false);
  assert.equal(g.claim(navKey(2, 'https://x')), true);
  now += 5000;
  assert.equal(g.claim(k), true);
});

test('バイパスは 1 回だけ消費され、期限切れは掃除される', () => {
  const now = 10_000;
  let map = addBypass({}, 'a', now, 60_000);
  map = addBypass(map, 'b', now - 70_000, 60_000); // 期限切れ
  const first = takeBypass(map, 'a', now);
  assert.equal(first.hit, true);
  assert.deepEqual(first.next, {});
  assert.equal(takeBypass(first.next, 'a', now).hit, false);
});

test('takeBypass は不正な入力に耐える', () => {
  assert.deepEqual(takeBypass(undefined, 'a', 0), { hit: false, next: {} });
  assert.deepEqual(takeBypass({ a: 'x' }, 'a', 0), { hit: false, next: {} });
});
