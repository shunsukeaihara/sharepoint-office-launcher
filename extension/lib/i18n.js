// chrome.i18n（_locales/<言語>/messages.json）による表示文字列の切り替え。

/** メッセージを取得する。見つからない場合はキーを返す（未翻訳の検出用）。 */
export function t(key, substitutions) {
  return chrome.i18n.getMessage(key, substitutions) || key;
}

/** data-i18n 属性を持つ要素の textContent を、ブラウザの表示言語のメッセージに置き換える。 */
export function localizePage(root = document) {
  document.documentElement.lang = chrome.i18n.getUILanguage();
  for (const el of root.querySelectorAll('[data-i18n]')) el.textContent = t(el.dataset.i18n);
}
