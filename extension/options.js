import { normalizeHost } from './lib/convert.js';
import { getConfig } from './lib/config.js';

const $ = (id) => document.getElementById(id);
const FIELDS = ['enabled', 'allowedHosts', 'openMode', 'autoLaunch', 'closeTabAfterLaunch'];

async function load() {
  const config = await getConfig();
  $('enabled').checked = config.enabled;
  $('allowedHosts').value = config.allowedHosts.join('\n');
  $('openMode').value = config.openMode;
  $('autoLaunch').checked = config.autoLaunch;
  $('closeTabAfterLaunch').checked = config.closeTabAfterLaunch;

  for (const k of config.managedKeys) $(k).disabled = true;
  $('managed').hidden = config.managedKeys.length === 0;
  $('save').disabled = FIELDS.every((k) => config.managedKeys.includes(k));
  $('ext-origin').textContent = `chrome-extension://${chrome.runtime.id}`;
}

async function save() {
  const lines = $('allowedHosts').value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const invalid = lines.filter((s) => !normalizeHost(s));
  if (invalid.length) {
    $('status').textContent = `不正なホスト名: ${invalid.join(', ')}（*.sharepoint.com のみ指定できます）`;
    return;
  }
  await chrome.storage.sync.set({
    enabled: $('enabled').checked,
    allowedHosts: [...new Set(lines.map(normalizeHost))],
    openMode: $('openMode').value,
    autoLaunch: $('autoLaunch').checked,
    closeTabAfterLaunch: $('closeTabAfterLaunch').checked,
  });
  $('status').textContent = '保存しました。';
  await load();
}

$('save').addEventListener('click', () => save().catch((e) => { $('status').textContent = String(e); }));
load();
