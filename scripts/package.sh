#!/bin/sh
# 配布用 zip を dist/ に作成する。
#   - sharepoint-office-launcher-<ver>.zip       : 納品物一式（README・拡張機能・テスト）
#   - sharepoint-office-launcher-<ver>-store.zip : Chrome Web Store / Edge Add-ons アップロード用（manifest がルート）
set -eu
cd "$(dirname "$0")/.."
VER=$(node -p "JSON.parse(require('fs').readFileSync('extension/manifest.json','utf8')).version")
NAME=sharepoint-office-launcher-$VER
rm -rf dist && mkdir -p dist/stage/$NAME
cp -R README.md LICENSE PRIVACY.md package.json assets store-assets extension test scripts dist/stage/$NAME/
(cd dist/stage && zip -qrX ../$NAME.zip $NAME -x '*.DS_Store')
(cd extension && zip -qrX ../dist/$NAME-store.zip . -x '*.DS_Store')
rm -rf dist/stage
ls -l dist
