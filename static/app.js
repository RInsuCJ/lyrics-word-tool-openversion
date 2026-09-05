/* ═══════════════════════════════════════════════════════════════
   app.js — 歌詞 Word 工具 前端邏輯  (with i18n)
   ═══════════════════════════════════════════════════════════════ */

'use strict';

// ── i18n ──────────────────────────────────────────────────────────────────────
const TRANSLATIONS = {
  zh: {
    brandName          : '歌詞 Word 工具',
    brandSub           : '自動排版 · A4 即時預覽 · 一鍵生成',
    searchLabel        : '🔍 搜尋歌曲',
    searchPlaceholder  : '歌名 / 歌手名…',
    searchBtn          : '搜尋',
    resultsLabel       : '搜尋結果',
    resultsEmpty       : '請輸入關鍵字開始搜尋',
    lyricsPreviewLabel : '歌詞預覽',
    lyricsNotLoaded    : '（尚未載入）',
    previewLabel       : '📄 A4 版面預覽',
    previewInitStatus  : '請先搜尋並載入歌詞',
    placeholderText    : '載入歌詞後<br>版面預覽將顯示在這裡',
    songInfoLabel      : '🎤 已載入歌曲',
    songInfoEmpty      : '（尚未選擇）',
    fontLabel          : '🔤 字體大小',
    titleFontLabel     : '標題',
    bodyFontLabel      : '歌詞',
    layoutLabel        : '📐 版面配置',
    autoLayout         : '↺  自動計算最佳版面',
    statsLabel         : '📊 版面資訊',
    statLines          : '歌詞行數',
    statPages          : '預計頁數',
    statFill           : '頁面利用率',
    generateBtn        : '⬇  生成並下載 Word 檔',
    prevPage           : '‹',
    nextPage           : '›',
    // pinyin
    pinyinLabel        : '📖 拼音 / 注音',
    pinyinOff          : '關閉',
    pinyinOn           : '開啟✔',
    pinyinLangLabel    : '語言',
    pinyinLangCant     : '粵語 (Jyutping)',
    pinyinLangMand     : '國語 (Pinyin)',
    pinyinLangJpn      : '日本語 (Romaji)',
    pinyinLangKor      : '한국어 (Korean)',
    pinyinSwapOff      : '🔄 拼音↑ 原文↓',
    pinyinSwapOn       : '🔄 原文↑ 拼音↓ ✔',
    // mobile tabs
    tabSearch           : '搜尋',
    tabPreview          : '預覽',
    tabSettings         : '設定',
    // dynamic (use {var} placeholders)
    pageLabel          : '第 {cur} 頁 / 共 {total} 頁',
    singleColToggle    : '單欄  （點擊切換雙欄）',
    doubleColToggle    : '雙欄  （點擊切換單欄）',
    singleColShort     : '單欄',
    doubleColShort     : '雙欄',
    layoutInfoText     : '{col} · {pt} pt',
    perfectLayout      : '✅ 版面完美，一頁搞定！',
    goodLayout         : '✅ 一頁完成，利用率 {pct}%',
    smallFont          : '⚠️ 一頁完成，利用率偏低 {pct}%',
    twoPages           : '📄 共 {pages} 頁，第 2 頁佔 {p2}%',
    generating         : '⏳ 生成中…',
    generateSuccess    : '✅ 下載成功！',
    loadingSearch      : '搜尋中…',
    loadingFetch       : '載入歌詞中…',
    processing         : '處理中…',
    noResults          : '找不到相關歌曲，請嘗試其他關鍵字',
    loaded             : '✓ 已載入《{title}》，共 {lines} 行',
    autoLayoutDone     : '↺ 自動版面已重新計算',
    downloadSuccess    : '🎉 Word 文檔已生成並下載',
    noQuery            : '請輸入搜尋關鍵字',
    noTitle            : '缺少歌曲資訊',
    searchError        : '搜尋失敗：{msg}',
    fetchError         : '載入失敗：{msg}',
    generateError      : '生成失敗：{msg}',
    serverError        : '伺服器錯誤',
    // history
    historyLabel       : '📋 歷史記錄',
    historyEmpty       : '尚無記錄',
    // custom mode
    modeSearch         : '🔍 搜尋',
    modeCustom         : '✏️ 自訂',
    customTitleLabel   : '🎵 歌曲資訊',
    customTitlePlaceholder : '歌名（必填）',
    customArtistPlaceholder: '歌手（選填）',
    customLyricsLabel  : '📝 歌詞內容',
    customLyricsPlaceholder: '在此貼上或輸入歌詞…',
    customLoadBtn      : '📄 載入歌詞',
    customNoTitle      : '請輸入歌名',
    customNoLyrics     : '請輸入歌詞內容',
    customLoaded       : '✓ 已載入自訂歌詞《{title}》，共 {lines} 行',
    resultCopyHint     : '👆 點擊歌曲填入資訊，再從 Genius 複製歌詞',
    geniusLink         : '🔗 前往 Genius 複製歌詞',
    pasteHint          : '從上方 Genius 連結複製歌詞，然後貼到下方',
    cleaning           : '清理歌詞中…',
    sectionTagHide     : '🏷️ 隱藏段落標記 [Verse] ✔',
    sectionTagShow     : '🏷️ 顯示段落標記 [Verse]',
    // tutorial
    tutorialTitle      : '📖 使用說明',
    tutorialStep1Title : '🔍 搜尋歌曲',
    tutorialStep1Desc  : '輸入歌名或歌手名，找到你要的歌',
    tutorialStep2Title : '🔗 複製歌詞',
    tutorialStep2Desc  : '點擊搜尋結果或 🔗 按鈕，前往 Genius 網站複製歌詞',
    tutorialStep3Title : '📝 貼上歌詞',
    tutorialStep3Desc  : '把歌詞貼到「自訂模式」的文字框，按「載入歌詞」',
    tutorialStep4Title : '⬇️ 生成 Word',
    tutorialStep4Desc  : '調整版面和拼音設定，一鍵下載排版精美的 Word 文檔！',
    tutorialClose      : '👍 了解了！',
  },
  en: {
    brandName          : 'Lyrics Word Tool',
    brandSub           : 'Auto layout · A4 live preview · One-click export',
    searchLabel        : '🔍 Search Songs',
    searchPlaceholder  : 'Song title / Artist name…',
    searchBtn          : 'Search',
    resultsLabel       : 'Results',
    resultsEmpty       : 'Enter keywords to start searching',
    lyricsPreviewLabel : 'Lyrics',
    lyricsNotLoaded    : '(Not loaded yet)',
    previewLabel       : '📄 A4 Layout Preview',
    previewInitStatus  : 'Search and load lyrics to preview',
    placeholderText    : 'Load lyrics to see<br>the layout preview here',
    songInfoLabel      : '🎤 Loaded Song',
    songInfoEmpty      : '(None selected)',
    fontLabel          : '🔤 Font Size',
    titleFontLabel     : 'Title',
    bodyFontLabel      : 'Lyrics',
    layoutLabel        : '📐 Layout',
    autoLayout         : '↺  Auto-calculate best layout',
    statsLabel         : '📊 Layout Info',
    statLines          : 'Lyric lines',
    statPages          : 'Est. pages',
    statFill           : 'Page fill',
    generateBtn        : '⬇  Generate & Download Word',
    prevPage           : '‹',
    nextPage           : '›',
    // pinyin
    pinyinLabel        : '📖 Romanisation',
    pinyinOff          : 'Off',
    pinyinOn           : 'On ✔',
    pinyinLangLabel    : 'Language',
    pinyinLangCant     : 'Cantonese (Jyutping)',
    pinyinLangMand     : 'Mandarin (Pinyin)',
    pinyinLangJpn      : 'Japanese (Romaji)',
    pinyinLangKor      : 'Korean (Romanization)',
    pinyinSwapOff      : '🔄 Romaji↑ Lyrics↓',
    pinyinSwapOn       : '🔄 Lyrics↑ Romaji↓ ✔',
    // mobile tabs
    tabSearch           : 'Search',
    tabPreview          : 'Preview',
    tabSettings         : 'Settings',
    // dynamic
    pageLabel          : 'Page {cur} of {total}',
    singleColToggle    : 'Single column  (click to switch)',
    doubleColToggle    : 'Double column  (click to switch)',
    singleColShort     : 'Single',
    doubleColShort     : 'Double',
    layoutInfoText     : '{col} · {pt} pt',
    perfectLayout      : '✅ Perfect layout — fits one page!',
    goodLayout         : '✅ Fits 1 page, {pct}% fill',
    smallFont          : '⚠️ Fits 1 page, low fill at {pct}%',
    twoPages           : '📄 {pages} pages, page 2 at {p2}%',
    generating         : '⏳ Generating…',
    generateSuccess    : '✅ Download complete!',
    loadingSearch      : 'Searching…',
    loadingFetch       : 'Loading lyrics…',
    processing         : 'Processing…',
    noResults          : 'No songs found. Try different keywords.',
    loaded             : '✓ Loaded "{title}" — {lines} lines',
    autoLayoutDone     : '↺ Auto layout recalculated',
    downloadSuccess    : '🎉 Word document generated and downloaded',
    noQuery            : 'Please enter search keywords',
    noTitle            : 'Missing song information',
    searchError        : 'Search failed: {msg}',
    fetchError         : 'Load failed: {msg}',
    generateError      : 'Generate failed: {msg}',
    serverError        : 'Server error',
    // history
    historyLabel       : '📋 History',
    historyEmpty       : 'No history yet',
    // custom mode
    modeSearch         : '🔍 Search',
    modeCustom         : '✏️ Custom',
    customTitleLabel   : '🎵 Song Info',
    customTitlePlaceholder : 'Song title (required)',
    customArtistPlaceholder: 'Artist (optional)',
    customLyricsLabel  : '📝 Lyrics',
    customLyricsPlaceholder: 'Paste or type lyrics here…',
    customLoadBtn      : '📄 Load Lyrics',
    customNoTitle      : 'Please enter a song title',
    customNoLyrics     : 'Please enter lyrics',
    customLoaded       : '✓ Loaded custom lyrics "{title}" — {lines} lines',
    resultCopyHint     : '👆 Click a song to fill info, then copy lyrics from Genius',
    geniusLink         : '🔗 Copy lyrics from Genius',
    pasteHint          : 'Copy lyrics from the Genius link above, then paste below',
    cleaning           : 'Cleaning lyrics…',
    sectionTagHide     : '🏷️ Hide section tags [Verse] ✔',
    sectionTagShow     : '🏷️ Show section tags [Verse]',
    // tutorial
    tutorialTitle      : '📖 How to Use',
    tutorialStep1Title : '🔍 Search Songs',
    tutorialStep1Desc  : 'Enter a song title or artist name to find your song',
    tutorialStep2Title : '🔗 Copy Lyrics',
    tutorialStep2Desc  : 'Click a result or 🔗 button to visit Genius and copy the lyrics',
    tutorialStep3Title : '📝 Paste Lyrics',
    tutorialStep3Desc  : 'Paste lyrics into the text box in Custom mode, then click "Load Lyrics"',
    tutorialStep4Title : '⬇️ Generate Word',
    tutorialStep4Desc  : 'Adjust layout & romanisation settings, then download your formatted Word doc!',
    tutorialClose      : '👍 Got it!',
  },
};

let currentLang = localStorage.getItem('lang') || 'zh';

/** Translate a key, substituting {var} placeholders. */
function t (key, vars = {}) {
  const str = (TRANSLATIONS[currentLang] || {})[key]
           || (TRANSLATIONS['zh'] || {})[key]
           || key;
  return str.replace(/\{(\w+)\}/g, (_, k) => (vars[k] != null ? vars[k] : ''));
}

/** Toggle language and refresh all static text. */
function toggleLang () {
  currentLang = currentLang === 'zh' ? 'en' : 'zh';
  localStorage.setItem('lang', currentLang);
  applyLang();
}

/** Apply current language to all data-i18n elements + dynamic UI. */
function applyLang () {
  // Static text nodes
  document.querySelectorAll('[data-i18n]').forEach(el => {
    const key = el.dataset.i18n;
    // Use innerHTML so placeholderText with <br> renders
    el.innerHTML = t(key);
  });

  // Placeholder attributes (input)
  document.querySelectorAll('[data-i18n-key]').forEach(el => {
    const key  = el.dataset.i18nKey;
    const attr = el.dataset.i18nAttr || 'placeholder';
    el.setAttribute(attr, t(key));
  });

  // Toggle button label
  document.getElementById('langToggle').textContent =
    currentLang === 'zh' ? 'EN' : '中文';

  // Dynamic elements that depend on state
  document.getElementById('colToggleBtn').textContent =
    state.twoCols ? t('doubleColToggle') : t('singleColToggle');

  // Pinyin toggle button
  document.getElementById('pinyinToggleBtn').textContent =
    state.withPinyin ? t('pinyinOn') : t('pinyinOff');

  // Pinyin lang select options
  document.querySelectorAll('#pinyinLangSelect option').forEach(opt => {
    const key = opt.dataset.i18n;
    if (key) opt.textContent = t(key);
  });

  if (state.song) {
    document.getElementById('layoutInfo').textContent =
      t('layoutInfoText', {
        col: state.twoCols ? t('doubleColShort') : t('singleColShort'),
        pt : state.bodyPt.toFixed(1),
      });
    renderPreview();   // re-render status text
  } else {
    document.getElementById('previewStatus').textContent = t('previewInitStatus');
    renderPreview();   // placeholder text
  }

  // Section tag toggle button
  document.getElementById('sectionTagBtn').textContent =
    state.hideSectionTags ? t('sectionTagHide') : t('sectionTagShow');

  // html lang attribute
  document.documentElement.lang = currentLang === 'zh' ? 'zh-TW' : 'en';
}

// ── Layout formula constants (mirrors word_gen.py) ────────────────────────────
const TITLE_SPACE_BEFORE = 4.0;
const TITLE_LINE_FACTOR  = 1.25;
const TITLE_SPACE_AFTER  = 6.0;
const BODY_LINE_FACTOR   = 1.35;
const BODY_SPACE_AFTER   = 1.0;
const TWO_COL_OVERHEAD   = 25.0;
const TWO_COL_GAP_PT     = 0.8 * 28.3465;
const USABLE_H_PT        = (29.7 - 2.54 - 2.54) * (72 / 2.54);  // ≈698 pt
const USABLE_W_PT        = (21.0 - 2.0  - 2.0 ) * (72 / 2.54);  // ≈482 pt (matches Word 2.0cm margins)
const SAFETY             = 0.96;
const BODY_MIN           = 8.0;
const BODY_MAX           = 16.0;
const BODY_STEP          = 0.5;
const DEFAULT_TITLE_PT   = 22.0;
const PINYIN_FONT_RATIO  = 0.70;   // mirrors word_gen.py
const PINYIN_LINE_FACTOR_JS = 1.20;

// ── Title width estimation (mirrors word_gen._est_text_width) ─────────────────
function estTextWidth (text, fontPt) {
  let w = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0);
    if ((cp >= 0x4E00 && cp <= 0x9FFF) ||
        (cp >= 0x3400 && cp <= 0x4DBF) ||
        (cp >= 0xF900 && cp <= 0xFAFF) ||
        (cp >= 0x3000 && cp <= 0x303F) ||
        (cp >= 0xFF00 && cp <= 0xFFEF)) {
      w += fontPt;
    } else {
      w += fontPt * 0.55;
    }
  }
  return w;
}

// ── Title resolution (mirrors word_gen._resolve_title) ────────────────────────
function resolveTitle (title, artist, titlePt) {
  const one   = `\u300a${title}\u300b${artist}`;
  const line1 = `\u300a${title}\u300b`;
  const line2 = artist;

  for (let step = 0; step < 9; step++) {
    const pt = Math.max(14.0, Math.round((titlePt - step * 0.5) * 10) / 10);
    if (estTextWidth(one, pt) <= USABLE_W_PT)
      return { lines: [one], fontPt: pt };
    if (estTextWidth(line1, pt) <= USABLE_W_PT &&
        estTextWidth(line2, pt) <= USABLE_W_PT)
      return { lines: [line1, line2], fontPt: pt };
    if (pt <= 14.0) break;
  }
  return { lines: [line1, line2], fontPt: 14.0 };
}

// ── Height estimators ─────────────────────────────────────────────────────────
function titleH  (pt) { return TITLE_SPACE_BEFORE + pt * TITLE_LINE_FACTOR + TITLE_SPACE_AFTER; }
function bodyLineH (pt) { return pt * BODY_LINE_FACTOR + BODY_SPACE_AFTER; }
function pinyinLineH (pt) { return pt * PINYIN_FONT_RATIO * PINYIN_LINE_FACTOR_JS; }
function unitH (pt, withPinyin) { return bodyLineH(pt) + (withPinyin ? pinyinLineH(pt) : 0); }

const BLANK_LINE_PT = 12.0;

function estimateSingle (n, bodyPt, titlePt, withPinyin = false, nBlank = 0) {
  const nContent = n - nBlank;
  let h = titleH(titlePt) + unitH(bodyPt, withPinyin) * nContent;
  h += withPinyin ? BLANK_LINE_PT * nBlank : bodyLineH(bodyPt) * nBlank;
  return h;
}
function estimateTwoCol (n, bodyPt, titlePt, withPinyin = false, nBlank = 0) {
  const colLines   = Math.ceil(n / 2);
  const colBlank   = Math.ceil(nBlank / 2);
  const colContent = colLines - colBlank;
  let h = titleH(titlePt) + unitH(bodyPt, withPinyin) * colContent;
  h += withPinyin ? BLANK_LINE_PT * colBlank : bodyLineH(bodyPt) * colBlank;
  return h + TWO_COL_OVERHEAD;
}

function findOptimalLayout (lyrics, titlePt = DEFAULT_TITLE_PT, withPinyin = false) {
  const lines   = lyrics.split('\n');
  const n       = lines.length;
  const nBlank  = lines.filter(l => !l.trim()).length;
  const t1      = USABLE_H_PT * SAFETY;
  const t2      = 2 * USABLE_H_PT * SAFETY;
  const steps   = Math.round((BODY_MAX - BODY_MIN) / BODY_STEP) + 1;

  // Phase 1: single column, 1 page (always)
  for (let i = 0; i < steps; i++) {
    const sz = Math.round((BODY_MAX - i * BODY_STEP) * 10) / 10;
    if (estimateSingle(n, sz, titlePt, withPinyin, nBlank) <= t1) return { bodyPt: sz, twoCols: false };
  }

  if (withPinyin) {
    // Pinyin lines are Latin — they overflow in narrow two-col cells.
    // Stay single-column and allow up to 2 pages instead.
    for (let i = 0; i < steps; i++) {
      const sz = Math.round((BODY_MAX - i * BODY_STEP) * 10) / 10;
      if (estimateSingle(n, sz, titlePt, withPinyin, nBlank) <= t2) return { bodyPt: sz, twoCols: false };
    }
    return { bodyPt: BODY_MIN, twoCols: false };  // fallback: min font, single col
  } else {
    // No pinyin — original two-column logic unchanged
    for (let i = 0; i < steps; i++) {
      const sz = Math.round((BODY_MAX - i * BODY_STEP) * 10) / 10;
      if (estimateTwoCol(n, sz, titlePt, false, nBlank) <= t1) return { bodyPt: sz, twoCols: true };
    }
    for (let i = 0; i < steps; i++) {
      const sz = Math.round((BODY_MAX - i * BODY_STEP) * 10) / 10;
      if (estimateTwoCol(n, sz, titlePt, false, nBlank) <= t2) return { bodyPt: sz, twoCols: true };
    }
    return { bodyPt: BODY_MIN, twoCols: true };
  }
}

// ── State ─────────────────────────────────────────────────────────────────────
let state = {
  searchResults : [],
  selectedIdx   : -1,
  song          : null,
  titlePt       : DEFAULT_TITLE_PT,
  bodyPt        : 12.0,
  twoCols       : false,
  currentPage   : 1,
  totalPages    : 1,
  withPinyin    : false,
  pinyinSwapped : false,
  pinyinLang    : 'cantonese',
  annotated     : null,   // list of [rom, lyric] from /api/romanise
  hideSectionTags: true,  // hide [Verse], [Chorus] etc. by default
};

// Regex to match section tags like [Verse 1], [Chorus], [Bridge], etc.
const SECTION_TAG_RE = /^\[.*?\]\s*$/;

/** Return lyrics with section tags filtered out if hideSectionTags is on. */
function getDisplayLyrics () {
  if (!state.song) return '';
  if (!state.hideSectionTags) return state.song.lyrics;
  return state.song.lyrics
    .split('\n')
    .filter(l => !SECTION_TAG_RE.test(l.trim()))
    .join('\n');
}

/** Toggle section tag visibility and re-render. */
function toggleSectionTags () {
  state.hideSectionTags = !state.hideSectionTags;
  const btn = document.getElementById('sectionTagBtn');
  btn.textContent = state.hideSectionTags ? t('sectionTagHide') : t('sectionTagShow');
  btn.className = state.hideSectionTags
    ? 'btn btn-accent full-width mt-6'
    : 'btn btn-secondary full-width mt-6';
  state.currentPage = 1;
  if (state.song) autoLayout();
}

// ── A4 canvas scale ───────────────────────────────────────────────────────────
function applyA4Scale () {
  const wrapper = document.getElementById('a4Wrapper');
  const page    = document.getElementById('a4Page');
  if (!page) return;
  const availW  = wrapper.clientWidth - 32;
  const scale   = Math.min(1, availW / 794);
  page.style.transform = `scale(${scale})`;
  // Shrink the wrapper height to match scaled page so no dead space
  const pageH   = page.offsetHeight;
  wrapper.style.height = `${pageH * scale + 32}px`;
}
window.addEventListener('resize', applyA4Scale);

// ── Layout info ───────────────────────────────────────────────────────────────
function calcLayoutInfo () {
  if (!state.song) return null;
  const lines  = getDisplayLyrics().split('\n');
  const n      = lines.length;
  const nBlank = lines.filter(l => !l.trim()).length;
  const contentH = state.twoCols
    ? estimateTwoCol(n, state.bodyPt, state.titlePt, state.withPinyin, nBlank)
    : estimateSingle(n, state.bodyPt, state.titlePt, state.withPinyin, nBlank);
  const pages    = Math.max(1, Math.ceil(contentH / USABLE_H_PT));
  const fillRatio= contentH / (pages * USABLE_H_PT);
  return { n, contentH, pages, fillPct: Math.round(fillRatio * 100) };
}

function updateStats () {
  const info = calcLayoutInfo();
  if (!info) {
    ['statLines','statPages','statFill'].forEach(id =>
      document.getElementById(id).textContent = '—');
    return;
  }
  document.getElementById('statLines').textContent = info.n;
  document.getElementById('statPages').textContent = `${info.pages}`;
  const fillEl = document.getElementById('statFill');
  fillEl.textContent = `${info.fillPct}%`;
  fillEl.style.color = info.fillPct >= 85 ? 'var(--success)'
                     : info.fillPct >= 60 ? 'var(--warn)' : 'var(--info)';

  const statusEl = document.getElementById('previewStatus');
  if (info.pages === 1) {
    if (info.fillPct >= 92)
      statusEl.textContent = t('perfectLayout');
    else if (info.fillPct >= 70)
      statusEl.textContent = t('goodLayout', { pct: info.fillPct });
    else
      statusEl.textContent = t('smallFont', { pct: info.fillPct });
  } else {
    const p2 = Math.round(((info.contentH - USABLE_H_PT) / USABLE_H_PT) * 100);
    statusEl.textContent = t('twoPages', { pages: info.pages, p2 });
  }
  statusEl.style.color = info.pages === 1 ? 'var(--success)' : 'var(--warn)';
}

// ── Render A4 preview ─────────────────────────────────────────────────────────
function renderPreview () {
  const wrapper = document.getElementById('a4Wrapper');

  if (!state.song) {
    wrapper.innerHTML = `
      <div class="a4-page" id="a4Page">
        <div class="a4-placeholder">
          <div class="placeholder-icon">🎵</div>
          <div class="placeholder-text">${t('placeholderText')}</div>
        </div>
      </div>`;
    applyA4Scale();
    return;
  }

  const { title, artist } = state.song;
  const lyrics     = getDisplayLyrics();
  const lines      = lyrics.split('\n');
  const nLines     = lines.length;
  const titlePt    = state.titlePt;
  const bodyPt     = state.bodyPt;
  const twoCols    = state.twoCols;
  const withPinyin = state.withPinyin;

  const info = calcLayoutInfo();
  state.totalPages = info ? info.pages : 1;
  if (state.currentPage > state.totalPages) state.currentPage = 1;

  // Page indicator
  const pageInd = document.getElementById('pageIndicator');
  if (state.totalPages > 1) {
    pageInd.style.display = 'flex';
    document.getElementById('pageLabel').textContent =
      t('pageLabel', { cur: state.currentPage, total: state.totalPages });
    document.getElementById('prevPageBtn').disabled = state.currentPage === 1;
    document.getElementById('nextPageBtn').disabled = state.currentPage === state.totalPages;
  } else {
    pageInd.style.display = 'none';
  }

  // Row geometry — walk lines to find page breaks (accounts for blank line height)
  const titH       = titleH(titlePt);
  const contentUH  = unitH(bodyPt, withPinyin);
  const blankH     = (withPinyin) ? BLANK_LINE_PT : bodyLineH(bodyPt);
  const swapped    = state.pinyinSwapped;
  const ann        = state.withPinyin && state.annotated ? state.annotated : null;

  // Estimate how many visual lines a Latin text occupies at a given pt size
  // Average char width for Arial Latin ≈ 0.55 * font_size
  const CHAR_W_RATIO = 0.55;
  function wrapLines (text, fontPt, availW) {
    if (!text || !text.trim()) return 1;
    const charsPerLine = Math.floor(availW / (fontPt * CHAR_W_RATIO));
    return Math.max(1, Math.ceil(text.length / charsPerLine));
  }
  const colAvailW = twoCols ? (USABLE_W_PT - TWO_COL_GAP_PT) / 2 : USABLE_W_PT;

  let totalRows, col1Lines, col2Lines;
  if (twoCols) {
    const half = Math.ceil(nLines / 2);
    col1Lines  = lines.slice(0, half);
    col2Lines  = lines.slice(half);
    totalRows  = col1Lines.length;
  } else {
    totalRows = nLines;
  }

  // Build page break indices by accumulating height
  const sourceLines = twoCols ? col1Lines : lines;
  function findRowsForPage (startRow, availH) {
    let h = 0;
    let row = startRow;
    while (row < totalRows) {
      if (sourceLines[row].trim() === '') {
        const lineH = blankH;
        if (h + lineH > availH) break;
        h += lineH;
      } else if (ann && swapped) {
        const absIdx = twoCols ? row : row;
        const romText = (ann[absIdx] || [''])[0] || '';
        const nWrap = wrapLines(romText, bodyPt, colAvailW);
        const lineH = pinyinLineH(bodyPt) + bodyLineH(bodyPt) * nWrap;
        if (h + lineH > availH) break;
        h += lineH;
      } else if (ann) {
        const absIdx = row;
        const romText = (ann[absIdx] || [''])[0] || '';
        const pinPt   = bodyPt * PINYIN_FONT_RATIO;
        const nWrap = wrapLines(romText, pinPt, colAvailW);
        const lineH = pinyinLineH(bodyPt) * nWrap + bodyLineH(bodyPt);
        if (h + lineH > availH) break;
        h += lineH;
      } else {
        const lineH = contentUH;
        if (h + lineH > availH) break;
        h += lineH;
      }
      row++;
    }

    // ── Paragraph-aware break: snap to nearest blank line ──
    // Only when page is full (not at end of content) and we have enough lines
    if (row < totalRows && row - startRow >= 4) {
      const fitted = row - startRow;
      const maxLookback = Math.max(3, Math.floor(fitted * 0.3));
      for (let i = row - 1; i > startRow && (row - i) <= maxLookback; i--) {
        if (sourceLines[i].trim() === '') {
          // Break after this blank line — next page starts with a fresh verse
          row = i + 1;
          break;
        }
      }
    }

    return row;
  }

  let rowStart, rowEnd;
  if (state.currentPage === 1) {
    rowStart = 0;
    rowEnd   = findRowsForPage(0, USABLE_H_PT - titH);
  } else {
    // Find where page 1 ended
    let pgEnd = findRowsForPage(0, USABLE_H_PT - titH);
    for (let pg = 2; pg < state.currentPage; pg++) {
      pgEnd = findRowsForPage(pgEnd, USABLE_H_PT);
    }
    rowStart = pgEnd;
    rowEnd   = findRowsForPage(pgEnd, USABLE_H_PT);
  }
  rowEnd = Math.min(rowEnd, totalRows);

  // Title HTML (1 or 2 lines with possible font reduction)
  const { lines: titleLines, fontPt: effTitlePt } = resolveTitle(title, artist, titlePt);
  let titleHTML = '';
  for (let i = 0; i < titleLines.length; i++) {
    const mTop = i === 0
      ? `margin-top:${TITLE_SPACE_BEFORE}pt;`
      : 'margin-top:0;';
    const mBot = i === titleLines.length - 1
      ? `margin-bottom:${TITLE_SPACE_AFTER}pt;`
      : 'margin-bottom:0;';
    titleHTML += `<div class="lyric-title" style="font-size:${effTitlePt}pt;line-height:${TITLE_LINE_FACTOR};${mTop}${mBot}">${escHtml(titleLines[i])}</div>`;
  }

  // Body HTML — render (pinyin + lyric) pairs when annotated
  // swapped already declared above for pagination
  // When swapped: pinyin = big black, lyrics = small grey
  const mainPt  = swapped ? bodyPt * PINYIN_FONT_RATIO : bodyPt;
  const subPt   = swapped ? bodyPt : bodyPt * PINYIN_FONT_RATIO;
  const lineStyle   = `font-size:${mainPt}pt;line-height:${BODY_LINE_FACTOR};margin-bottom:${BODY_SPACE_AFTER}pt;`;
  const pinStyle    = `font-size:${subPt}pt;line-height:${swapped ? BODY_LINE_FACTOR : PINYIN_LINE_FACTOR_JS};${swapped ? '' : 'color:#808080;'}margin:0;`;
  const lyricSubStyle = swapped ? `font-size:${mainPt}pt;line-height:${PINYIN_LINE_FACTOR_JS};color:#808080;margin:0;` : null;
  // ann already declared above for pagination

  const blankStyle  = `font-size:12pt;line-height:1;margin:0;`;

  function renderLines (linesArr, startIdx) {
    return linesArr.map((l, i) => {
      const absIdx = startIdx + i;
      // Blank line = verse separator — single 12pt line, no pinyin
      if (ann && !l.trim()) {
        return `<p style="${blankStyle}">&nbsp;</p>`;
      }
      if (ann) {
        const rom = escHtml((ann[absIdx] || [''])[0] || '\u00a0');
        const lyr = escHtml(l) || '&nbsp;';
        if (swapped) {
          // Swapped: lyrics on top (small, grey) → pinyin below (big, black)
          return `<p class="pinyin-line" style="${lyricSubStyle}">${lyr}</p>` +
                 `<p class="pinyin-line" style="${pinStyle}">${rom}</p>`;
        } else {
          // Normal: pinyin on top (small, grey) → lyrics below (big, black)
          return `<p class="pinyin-line" style="${pinStyle}">${rom}</p>` +
                 `<p style="${lineStyle}">${lyr}</p>`;
        }
      }
      return `<p style="${lineStyle}">${escHtml(l) || '&nbsp;'}</p>`;
    }).join('');
  }

  let bodyHTML = '';
  if (twoCols) {
    const c1 = col1Lines.slice(rowStart, rowEnd);
    const c2 = col2Lines.slice(rowStart, rowEnd);
    // For two-col, column 1 starts at 0, column 2 starts at col1Lines.length
    const c2StartAbs = col1Lines.length;
    bodyHTML = `<div class="lyric-two-col">
      <div class="lyric-col">${renderLines(c1, rowStart)}</div>
      <div class="lyric-col">${renderLines(c2, c2StartAbs + rowStart)}</div>
    </div>`;
  } else {
    const rows = lines.slice(rowStart, rowEnd);
    bodyHTML = `<div class="lyric-body-single">${renderLines(rows, rowStart)}</div>`;
  }

  wrapper.innerHTML = `
    <div class="a4-page" id="a4Page">
      ${state.currentPage === 1 ? titleHTML : ''}
      ${bodyHTML}
    </div>`;

  applyA4Scale();
  updateStats();
  scaleMobileA4();
}

function escHtml (str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Page navigation ───────────────────────────────────────────────────────────
function changePage (delta) {
  state.currentPage = Math.max(1, Math.min(state.totalPages, state.currentPage + delta));
  renderPreview();
}

// ── Font step ─────────────────────────────────────────────────────────────────
function stepFont (which, delta) {
  if (which === 'title') {
    state.titlePt = Math.max(14, Math.min(32, Math.round((state.titlePt + delta) * 10) / 10));
    document.getElementById('titlePtDisplay').value = state.titlePt.toFixed(1);
  } else {
    state.bodyPt = Math.max(BODY_MIN, Math.min(BODY_MAX, Math.round((state.bodyPt + delta) * 10) / 10));
    document.getElementById('bodyPtDisplay').value = state.bodyPt.toFixed(1);
  }
  state.currentPage = 1;
  renderPreview();
}

function setFontDirect (which, val) {
  const v = parseFloat(val);
  if (isNaN(v)) return;
  if (which === 'title') {
    state.titlePt = Math.max(14, Math.min(32, Math.round(v * 10) / 10));
    document.getElementById('titlePtDisplay').value = state.titlePt.toFixed(1);
  } else {
    state.bodyPt = Math.max(BODY_MIN, Math.min(BODY_MAX, Math.round(v * 10) / 10));
    document.getElementById('bodyPtDisplay').value = state.bodyPt.toFixed(1);
  }
  state.currentPage = 1;
  renderPreview();
}

// ── Column toggle ─────────────────────────────────────────────────────────────
function toggleCols () {
  state.twoCols = !state.twoCols;
  document.getElementById('colToggleBtn').textContent =
    state.twoCols ? t('doubleColToggle') : t('singleColToggle');
  state.currentPage = 1;
  renderPreview();
}

// ── Pinyin toggle ────────────────────────────────────────────────────────────
function togglePinyin () {
  state.withPinyin = !state.withPinyin;
  const btn     = document.getElementById('pinyinToggleBtn');
  const langRow = document.getElementById('pinyinLangRow');
  const swapBtn = document.getElementById('pinyinSwapBtn');
  btn.textContent = state.withPinyin ? t('pinyinOn') : t('pinyinOff');
  btn.className   = state.withPinyin
    ? 'btn btn-accent full-width'
    : 'btn btn-secondary full-width';
  langRow.style.display  = state.withPinyin ? 'flex' : 'none';
  swapBtn.style.display  = state.withPinyin ? 'block' : 'none';

  if (!state.withPinyin) {
    state.pinyinSwapped = false;
    swapBtn.textContent = t('pinyinSwapOff');
    swapBtn.className   = 'btn btn-outline full-width mt-6';
  }

  if (state.withPinyin && state.song) {
    // Fetch romanisation first, then auto-layout
    fetchRomanisation().then(() => autoLayout());
  } else {
    state.annotated = null;
    autoLayout();
  }
}

function togglePinyinSwap () {
  state.pinyinSwapped = !state.pinyinSwapped;
  const btn = document.getElementById('pinyinSwapBtn');
  btn.textContent = state.pinyinSwapped ? t('pinyinSwapOn') : t('pinyinSwapOff');
  btn.className   = state.pinyinSwapped
    ? 'btn btn-accent full-width mt-6'
    : 'btn btn-outline full-width mt-6';
  renderPreview();
}

function onPinyinLangChange () {
  state.pinyinLang = document.getElementById('pinyinLangSelect').value;
  if (state.withPinyin && state.song) {
    fetchRomanisation().then(() => renderPreview());
  }
}

async function fetchRomanisation () {
  if (!state.song) return;
  try {
    const res  = await fetch('/api/romanise', {
      method : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body   : JSON.stringify({ lyrics: getDisplayLyrics(), lang: state.pinyinLang }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    state.annotated = data.annotated;  // [[rom, lyric], ...]
  } catch (err) {
    toast(`Romanisation error: ${err.message}`, 'error');
    state.annotated = null;
  }
}

// ── Auto layout ───────────────────────────────────────────────────────────────
function autoLayout () {
  if (!state.song) return;
  const { bodyPt, twoCols } = findOptimalLayout(
    getDisplayLyrics(), state.titlePt, state.withPinyin);
  state.bodyPt  = bodyPt;
  state.twoCols = twoCols;
  document.getElementById('bodyPtDisplay').value  = bodyPt.toFixed(1);
  document.getElementById('colToggleBtn').textContent   =
    twoCols ? t('doubleColToggle') : t('singleColToggle');
  state.currentPage = 1;
  renderPreview();
  toast(t('autoLayoutDone'), 'info');
}

// ── Search ────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('searchInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') doSearch();
  });
});

async function doSearch () {
  const query = document.getElementById('searchInput').value.trim();
  if (!query) return;
  showLoading(t('loadingSearch'));
  try {
    const res  = await fetch('/api/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    state.searchResults = data.results || [];
    state.selectedIdx   = -1;
    renderResults();
    if (state.searchResults.length === 0)
      toast(t('noResults'), 'error');
  } catch (err) {
    toast(t('searchError', { msg: err.message }), 'error');
  } finally {
    hideLoading();
  }
}

function renderResults () {
  const list = document.getElementById('resultsList');
  if (state.searchResults.length === 0) {
    list.innerHTML = `<li class="results-empty">${t('noResults')}</li>`;
    return;
  }
  list.innerHTML = state.searchResults.map((r, i) => `
    <li class="result-item ${i === state.selectedIdx ? 'selected' : ''}"
        onclick="selectResult(${i})">
      <div class="result-info">
        <div class="result-title">${escHtml(r.title)}</div>
        <div class="result-artist">${escHtml(r.artist)}</div>
      </div>
      <a class="result-genius-link" href="${escHtml(r.url)}" target="_blank" rel="noopener"
         onclick="event.stopPropagation(); fillFromResult(${i})" title="${t('geniusLink')}">
        🔗
      </a>
    </li>`).join('') +
    `<li class="results-hint">${t('resultCopyHint')}</li>`;
}

function selectResult (idx) {
  state.selectedIdx = idx;
  renderResults();
  const r = state.searchResults[idx];

  // Fill in custom mode fields
  document.getElementById('customTitle').value  = r.title;
  document.getElementById('customArtist').value = r.artist;

  // Switch to custom/paste mode
  switchMode('custom');

  // Open Genius page for user to copy lyrics
  if (r.url) {
    window.open(r.url, '_blank', 'noopener');
  }

  // Focus the lyrics textarea
  setTimeout(() => {
    document.getElementById('customLyrics').focus();
  }, 300);
}

/** Fill custom fields from a search result (called when clicking 🔗 Genius link). */
function fillFromResult (idx) {
  state.selectedIdx = idx;
  const r = state.searchResults[idx];

  // Fill in custom mode fields
  document.getElementById('customTitle').value  = r.title;
  document.getElementById('customArtist').value = r.artist;

  // Switch to custom/paste mode
  switchMode('custom');

  // Focus the lyrics textarea after a short delay
  setTimeout(() => {
    document.getElementById('customLyrics').focus();
  }, 300);
}

// ── Generate Word ─────────────────────────────────────────────────────────────
async function doGenerate () {
  if (!state.song) return;
  const btn      = document.getElementById('generateBtn');
  const statusEl = document.getElementById('generateStatus');
  btn.disabled = true;
  statusEl.textContent = t('generating');
  statusEl.style.color = 'var(--info)';

  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        title          : state.song.title,
        artist         : state.song.artist,
        lyrics         : getDisplayLyrics(),
        title_pt       : state.titlePt,
        body_pt        : state.bodyPt,
        two_cols       : state.twoCols,
        with_pinyin    : state.withPinyin,
        pinyin_lang    : state.pinyinLang,
        pinyin_swapped : state.pinyinSwapped,
      }),
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || t('serverError'));
    }

    const blob     = await res.blob();
    const url      = URL.createObjectURL(blob);
    const a        = document.createElement('a');
    const filename = state.song.artist
                       ? `《${state.song.title}》${state.song.artist}.docx`
                       : `《${state.song.title}》.docx`;
    a.href = url; a.download = filename; a.click();
    URL.revokeObjectURL(url);

    statusEl.textContent = t('generateSuccess');
    statusEl.style.color = 'var(--success)';
    addToHistory(state.song.title, state.song.artist);
    toast(t('downloadSuccess'), 'success');
  } catch (err) {
    statusEl.textContent = t('generateError', { msg: err.message });
    statusEl.style.color = 'var(--accent)';
    toast(t('generateError', { msg: err.message }), 'error');
  } finally {
    btn.disabled = false;
    setTimeout(() => { statusEl.textContent = ''; }, 5000);
  }
}

// ── Loading ───────────────────────────────────────────────────────────────────
function showLoading (msg) {
  document.getElementById('loadingMsg').textContent = msg || t('processing');
  document.getElementById('loadingOverlay').style.display = 'flex';
}
function hideLoading () {
  document.getElementById('loadingOverlay').style.display = 'none';
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function toast (msg, type = 'info') {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 3200);
}

// ── Tutorial ──────────────────────────────────────────────────────────────────
function showTutorial () {
  // Refresh i18n text inside the modal
  applyLang();
  document.getElementById('tutorialOverlay').style.display = 'flex';
}
function closeTutorial (e) {
  // If called from overlay click, only close if clicking the backdrop itself
  if (e && e.target !== e.currentTarget) return;
  document.getElementById('tutorialOverlay').style.display = 'none';
}

// ── History ───────────────────────────────────────────────────────────────────
const HISTORY_KEY     = 'lyrics_tool_history';
const HISTORY_MAX     = 20;

function loadHistory () {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_KEY)) || [];
  } catch { return []; }
}

function saveHistory (list) {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(list));
}

function addToHistory (title, artist) {
  let list = loadHistory();
  // Remove duplicate (same title+artist)
  list = list.filter(h => !(h.title === title && h.artist === artist));
  // Add to front
  list.unshift({ title, artist });
  // Trim to max
  if (list.length > HISTORY_MAX) list = list.slice(0, HISTORY_MAX);
  saveHistory(list);
  renderHistory();
}

function removeFromHistory (idx) {
  const list = loadHistory();
  list.splice(idx, 1);
  saveHistory(list);
  renderHistory();
}

function clickHistory (idx) {
  const list = loadHistory();
  const entry = list[idx];
  if (!entry) return;
  // Fill custom fields and switch to custom mode
  document.getElementById('customTitle').value  = entry.title;
  document.getElementById('customArtist').value = entry.artist;
  switchMode('custom');
  // Also trigger search to find the Genius link
  document.getElementById('searchInput').value = `${entry.title} ${entry.artist}`;
  doSearch();
}

function toggleHistory () {
  const panel = document.getElementById('historyPanel');
  const icon  = document.getElementById('historyToggleIcon');
  const open  = panel.style.display === 'none';
  panel.style.display = open ? '' : 'none';
  icon.classList.toggle('open', open);
}

function renderHistory () {
  const list = loadHistory();
  const ul   = document.getElementById('historyList');
  if (!ul) return;
  if (list.length === 0) {
    ul.innerHTML = `<li class="history-empty">${t('historyEmpty')}</li>`;
    return;
  }
  ul.innerHTML = list.map((h, i) => `
    <li class="history-item" onclick="clickHistory(${i})">
      <div class="history-item-info">
        <div class="history-item-title">${escHtml(h.title)}</div>
        <div class="history-item-artist">${escHtml(h.artist)}</div>
      </div>
      <button class="history-delete" onclick="event.stopPropagation(); removeFromHistory(${i})" title="刪除">✕</button>
    </li>`).join('');
}

// ── Mode switch (Search / Custom) ─────────────────────────────────────────────
function switchMode (mode) {
  const searchMode = document.getElementById('searchMode');
  const customMode = document.getElementById('customMode');
  const searchBtn  = document.getElementById('modeSearchBtn');
  const customBtn  = document.getElementById('modeCustomBtn');

  if (mode === 'custom') {
    searchMode.style.display = 'none';
    customMode.style.display = '';
    searchBtn.classList.remove('active');
    customBtn.classList.add('active');
  } else {
    searchMode.style.display = '';
    customMode.style.display = 'none';
    searchBtn.classList.add('active');
    customBtn.classList.remove('active');
  }
}

// ── Load custom lyrics ────────────────────────────────────────────────────────
async function loadCustomLyrics () {
  const title  = document.getElementById('customTitle').value.trim();
  const artist = document.getElementById('customArtist').value.trim();
  let   lyrics = document.getElementById('customLyrics').value.trim();

  if (!title) {
    toast(t('customNoTitle'), 'error');
    return;
  }
  if (!lyrics) {
    toast(t('customNoLyrics'), 'error');
    return;
  }

  // Clean the pasted lyrics via server
  showLoading(t('cleaning'));
  try {
    const res = await fetch('/api/clean', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lyrics }),
    });
    const data = await res.json();
    if (data.lyrics) {
      lyrics = data.lyrics;
      document.getElementById('customLyrics').value = lyrics;
    }
  } catch (err) {
    // If cleaning fails, use the raw lyrics as-is
  }
  hideLoading();

  // Build song data in the same format
  state.song = {
    title  : title,
    artist : artist || '',
    lyrics : lyrics,
  };
  state.currentPage = 1;

  // Auto layout
  const { bodyPt, twoCols } = findOptimalLayout(
    lyrics, state.titlePt, state.withPinyin);
  state.bodyPt  = bodyPt;
  state.twoCols = twoCols;

  // Update UI
  document.getElementById('bodyPtDisplay').value  = bodyPt.toFixed(1);
  document.getElementById('titlePtDisplay').value = state.titlePt.toFixed(1);
  document.getElementById('colToggleBtn').textContent   =
    twoCols ? t('doubleColToggle') : t('singleColToggle');
  document.getElementById('songInfo').textContent =
    artist ? `${title}\n${artist}` : title;
  document.getElementById('layoutInfo').textContent =
    t('layoutInfoText', {
      col: twoCols ? t('doubleColShort') : t('singleColShort'),
      pt : bodyPt.toFixed(1),
    });
  document.getElementById('lyricsPreviewText').textContent =
    artist ? `《${title}》${artist}\n\n${lyrics}` : `《${title}》\n\n${lyrics}`;
  document.getElementById('generateBtn').disabled = false;
  document.getElementById('generateBtnPreview').disabled = false;

  // If pinyin is on, fetch romanisation
  if (state.withPinyin) {
    await fetchRomanisation();
  }

  renderPreview();
  addToHistory(title, artist);
  switchMobileTab('preview');
  const lineCount = lyrics.split('\n').length;
  toast(t('customLoaded', { title, lines: lineCount }), 'success');
}

// ── Mobile Tab Bar ────────────────────────────────────────────────────────────
function isMobile () { return window.innerWidth <= 960; }

function switchMobileTab (tabName) {
  if (!isMobile()) return;
  // Toggle panel visibility
  document.querySelectorAll('.panel[data-tab]').forEach(p => {
    p.classList.toggle('mobile-active', p.dataset.tab === tabName);
  });
  // Toggle tab button active state
  document.querySelectorAll('.tab-item').forEach(btn => {
    btn.classList.toggle('active', btn.id === 'tab' + tabName.charAt(0).toUpperCase() + tabName.slice(1));
  });
  // Scale A4 preview when switching to preview tab
  // Double rAF + timeout to ensure panel is laid out before measuring
  if (tabName === 'preview') {
    requestAnimationFrame(() => requestAnimationFrame(scaleMobileA4));
    setTimeout(scaleMobileA4, 100);
  }
  window.scrollTo(0, 0);
}

function scaleMobileA4 () {
  if (!isMobile()) return;
  const wrapper = document.getElementById('a4Wrapper');
  const page    = document.getElementById('a4Page');
  if (!wrapper || !page) return;
  const availW = wrapper.clientWidth - 16;  // 8px padding each side
  const pageW  = 794;
  const scale  = Math.min(1, availW / pageW);
  page.style.transform = `scale(${scale})`;
  // Adjust wrapper height so it doesn't overflow
  wrapper.style.height = (page.offsetHeight * scale) + 'px';
}

function initMobile () {
  if (!isMobile()) {
    // Desktop: ensure all panels are visible (remove mobile-active class)
    document.querySelectorAll('.panel[data-tab]').forEach(p => {
      p.classList.remove('mobile-active');
    });
    return;
  }
  // Default: show search tab
  switchMobileTab('search');
}

// Re-check on resize (e.g. rotating phone, or resizing browser)
window.addEventListener('resize', () => {
  if (isMobile()) {
    // Ensure one tab is active
    const anyActive = document.querySelector('.panel[data-tab].mobile-active');
    if (!anyActive) switchMobileTab('search');
    scaleMobileA4();
  } else {
    // Desktop: remove mobile classes so all panels show
    document.querySelectorAll('.panel[data-tab]').forEach(p => {
      p.classList.remove('mobile-active');
    });
  }
});

// ── Init ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  applyLang();
  renderPreview();
  renderHistory();
  initMobile();
  if (!isMobile()) document.getElementById('searchInput').focus();
});
