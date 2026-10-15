import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', 'extension');
const manifest = JSON.parse(readFileSync(join(root, 'manifest.json'), 'utf8'));

test('manifest は MV3 で、参照ファイルがすべて存在する', () => {
  assert.equal(manifest.manifest_version, 3);
  const files = [
    manifest.background.service_worker,
    manifest.options_ui.page,
    manifest.storage.managed_schema,
    ...Object.values(manifest.icons),
    'launch.html',
    'launch.js',
  ];
  for (const f of files) assert.ok(existsSync(join(root, f)), `missing: ${f}`);
  JSON.parse(readFileSync(join(root, manifest.storage.managed_schema), 'utf8'));
});

test('権限は最小限（tabs / <all_urls> / web_accessible_resources なし）', () => {
  assert.deepEqual([...manifest.permissions].sort(), ['storage', 'webNavigation', 'webRequest']);
  assert.deepEqual(manifest.host_permissions, ['https://*.sharepoint.com/*']);
  assert.equal(manifest.web_accessible_resources, undefined);
  assert.equal(manifest.content_scripts, undefined);
});
