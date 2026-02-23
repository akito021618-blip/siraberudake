'use strict';

// ===== データ管理 =====
const STORAGE_KEY = 'webtest_data_v1';
const STORAGE_META_KEY = 'webtest_meta_v1';

let data = []; // [{ question: string, answer: string }]

/** localStorage からキャッシュ読み込み */
function loadFromCache() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    data = JSON.parse(raw);
    return data.length > 0;
  } catch (e) {
    return false;
  }
}

/** localStorage へキャッシュ保存 */
function saveToCache(rows, fileName) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
    localStorage.setItem(STORAGE_META_KEY, JSON.stringify({
      fileName: fileName || '不明',
      count: rows.length,
      savedAt: new Date().toLocaleString('ja-JP'),
    }));
  } catch (e) {
    // ストレージ容量超過などは無視
  }
}

/** キャッシュ削除 */
function clearCache() {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(STORAGE_META_KEY);
}

/** キャッシュのメタ情報取得 */
function getCacheMeta() {
  try {
    const raw = localStorage.getItem(STORAGE_META_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

// ===== Excel パース =====

/**
 * ArrayBuffer を受け取り、先頭シートを解析してデータをセット
 * @param {ArrayBuffer} buffer
 * @param {string} fileName
 */
function parseExcel(buffer, fileName) {
  const workbook = XLSX.read(buffer, { type: 'array' });

  const parsed = [];
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    for (const row of rows) {
      const question = String(row[0] ?? '').trim();
      const answer = String(row[1] ?? '').trim();
      if (question && answer) {
        parsed.push({ question, answer });
      }
    }
  }

  if (parsed.length === 0) {
    setStatus('データが見つかりませんでした。A列に問題文、B列に答えが必要です。');
    return;
  }

  data = parsed;
  saveToCache(data, fileName);
  setStatus(`${data.length}件 読み込みました（${workbook.SheetNames.length}シート）`);
  hideUpload();
  updateDataInfo();
}

// ===== 検索 =====

let debounceTimer = null;

function onInput() {
  const val = searchInput.value;
  clearBtn.style.display = val ? 'flex' : 'none';

  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => renderResults(val), 60);
}

function renderResults(query) {
  const q = query.trim();
  if (!q) {
    resultsEl.innerHTML = '';
    if (data.length > 0) {
      setStatus(`${data.length}件 読み込み済み — キーワードを入力してください`);
    }
    return;
  }

  const lower = q.toLowerCase();
  const matches = data.filter(item =>
    item.question.toLowerCase().includes(lower) ||
    item.answer.toLowerCase().includes(lower)
  );

  if (matches.length === 0) {
    resultsEl.innerHTML = '<div class="no-result">見つかりませんでした</div>';
    setStatus('');
    return;
  }

  setStatus(`${matches.length}件 見つかりました`);
  resultsEl.innerHTML = matches.map(item => `
    <div class="result-card" role="listitem">
      <div class="result-question">${escapeHtml(item.question).replace(makeHighlightRE(q), '<mark class="hl">$&</mark>')}</div>
      <div class="result-answer">${escapeHtml(item.answer).replace(makeHighlightRE(q), '<mark class="hl">$&</mark>')}</div>
    </div>
  `).join('');
}

function makeHighlightRE(query) {
  return new RegExp(escapeRegExp(query), 'gi');
}

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ===== UI ヘルパー =====

const statusEl = document.getElementById('status');
const resultsEl = document.getElementById('results');
const uploadSection = document.getElementById('uploadSection');
const searchInput = document.getElementById('searchInput');
const clearBtn = document.getElementById('clearBtn');

function setStatus(msg) {
  statusEl.textContent = msg;
}

function showUpload() {
  uploadSection.style.display = 'block';
}

function hideUpload() {
  uploadSection.style.display = 'none';
}

function updateDataInfo() {
  const meta = getCacheMeta();
  const el = document.getElementById('dataInfo');
  if (!el) return;
  if (meta) {
    el.textContent = `${meta.fileName}（${meta.count}件）— ${meta.savedAt}`;
  } else {
    el.textContent = 'データなし';
  }
}

// ===== 設定パネル =====

const settingsBtn = document.getElementById('settingsBtn');
const settingsPanel = document.getElementById('settingsPanel');
const overlay = document.getElementById('overlay');

function openSettings() {
  updateDataInfo();
  settingsPanel.style.display = 'block';
  overlay.style.display = 'block';
}

function closeSettings() {
  settingsPanel.style.display = 'none';
  overlay.style.display = 'none';
}

settingsBtn.addEventListener('click', openSettings);
document.getElementById('closeSettingsBtn').addEventListener('click', closeSettings);
overlay.addEventListener('click', closeSettings);

document.getElementById('reloadBtn').addEventListener('click', () => {
  closeSettings();
  triggerFileUpload();
});

document.getElementById('clearCacheBtn').addEventListener('click', () => {
  if (!confirm('保存されたデータを削除しますか？')) return;
  clearCache();
  data = [];
  resultsEl.innerHTML = '';
  setStatus('データを削除しました');
  showUpload();
  closeSettings();
});

// ===== ファイル入力 =====

const fileInput = document.getElementById('fileInput');

function triggerFileUpload() {
  fileInput.value = '';
  fileInput.click();
}

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  setStatus('読み込み中...');
  const reader = new FileReader();
  reader.onload = (ev) => parseExcel(ev.target.result, file.name);
  reader.onerror = () => setStatus('ファイルの読み込みに失敗しました');
  reader.readAsArrayBuffer(file);
});

// ===== 検索入力 =====

searchInput.addEventListener('input', onInput);

clearBtn.addEventListener('click', () => {
  searchInput.value = '';
  clearBtn.style.display = 'none';
  renderResults('');
  searchInput.focus();
});

// ===== 起動時 =====

(function init() {
  if (loadFromCache()) {
    const meta = getCacheMeta();
    setStatus(
      meta
        ? `${meta.count}件 読み込み済み — キーワードを入力してください`
        : `${data.length}件 読み込み済み`
    );
  } else {
    setStatus('');
    showUpload();
  }
})();
