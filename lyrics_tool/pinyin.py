"""
pinyin.py — Romanisation helpers for Cantonese (Jyutping) and Mandarin (Pinyin).

Usage:
    from lyrics_tool.pinyin import annotate_lyrics
    pairs = annotate_lyrics(lyrics, lang='auto')   # list of (pinyin_line, lyric_line)

Dependencies (install once):
    pip install pypinyin pycantonese
"""

import re

# ── Constants ─────────────────────────────────────────────────────────────────

# Characters that appear exclusively in written Cantonese (not standard Mandarin).
# Each character here is diagnostic: if it appears, the text is very likely Cantonese.
_CANTONESE_MARKERS = frozenset(
    '唔'   # negative (not / don't)     — Mandarin: 不/沒
    '嘅'   # possessive / de            — Mandarin: 的
    '喺'   # at / in (location)         — Mandarin: 在
    '冇'   # don't have                 — Mandarin: 沒有
    '嗰'   # that                       — Mandarin: 那
    '咁'   # so / like this             — Mandarin: 這樣/那樣
    '佢'   # he / she / it              — Mandarin: 他/她/它
    '㗎'   # sentence-final particle
    '喎'   # hearsay particle
    '嚟'   # come                       — Mandarin: 來
    '攞'   # take / get                 — Mandarin: 拿
    '囉'   # sentence-final particle
    '咋'   # only (Cantonese sense)     — Mandarin: 只
    '嗻'   # sentence-final particle
    '啩'   # maybe / I suppose
    '嘥'   # waste
    '冚'   # all / entire
    '掂'   # settled / done
    '嗱'   # hey! (attention grabber)
    '噃'   # sentence-final particle
    '囖'   # sentence-final particle
    '哋'   # plural suffix (佢哋/我哋)  — Mandarin: 們
)

# ── Language detection ────────────────────────────────────────────────────────

def detect_lang(text: str) -> str:
    """
    Return 'cantonese' if text contains Cantonese-specific characters,
    else 'mandarin'.
    """
    for ch in text:
        if ch in _CANTONESE_MARKERS:
            return 'cantonese'
    return 'mandarin'


# ── Cantonese Jyutping ────────────────────────────────────────────────────────

def _line_to_jyutping(line: str) -> str:
    """Convert one line to space-separated Jyutping. Non-CJK kept as-is."""
    try:
        import pycantonese
        pairs = pycantonese.characters_to_jyutping(line)
        # pairs: list of (character, jyutping_or_None)
        parts = []
        for ch, jp in pairs:
            if jp:
                parts.append(jp)
            elif ch.strip():
                parts.append(ch)   # punctuation / Latin — keep
        return ' '.join(parts)
    except Exception:
        return ''


# ── Mandarin Pinyin ───────────────────────────────────────────────────────────

def _line_to_pinyin(line: str) -> str:
    """Convert one line to space-separated pinyin with tone marks."""
    try:
        from pypinyin import lazy_pinyin, Style
        syllables = lazy_pinyin(line, style=Style.TONE)
        return ' '.join(syllables)
    except Exception:
        return ''


# ── Japanese Romaji ───────────────────────────────────────────────────────────

def _line_to_romaji(line: str) -> str:
    """Convert one line of Japanese text to space-separated romaji."""
    try:
        import pykakasi
        kakasi = pykakasi.kakasi()
        result = kakasi.convert(line)
        parts = []
        for item in result:
            rom = item.get('hepburn', '') or item.get('orig', '')
            if rom.strip():
                parts.append(rom)
        return ' '.join(parts)
    except Exception:
        return ''


# ── Korean Romanization ──────────────────────────────────────────────────────

def _line_to_korean_roman(line: str) -> str:
    """Convert one line of Korean (Hangul) text to romanized form."""
    try:
        from korean_romanizer.romanizer import Romanizer
        r = Romanizer(line)
        return r.romanize()
    except Exception:
        return ''


# ── Public API ────────────────────────────────────────────────────────────────

def line_to_romanisation(line: str, lang: str = 'auto') -> str:
    """
    Convert one lyric line to a romanisation string.

    lang:  'auto'       — detect from content
           'cantonese'  — force Jyutping
           'mandarin'   — force Pinyin
           'japanese'   — force Romaji
           'korean'     — force Korean Romanization

    Returns '' for blank lines.
    """
    if not line.strip():
        return ''

    if lang == 'japanese':
        return _line_to_romaji(line)

    if lang == 'korean':
        return _line_to_korean_roman(line)

    effective = detect_lang(line) if lang == 'auto' else lang

    if effective == 'cantonese':
        return _line_to_jyutping(line)
    else:
        return _line_to_pinyin(line)


def annotate_lyrics(
    lyrics: str,
    lang: str = 'auto',
) -> list[tuple[str, str]]:
    """
    Annotate every lyric line with romanisation.

    Returns:
        list of (romanisation_line, lyric_line)
        Both are empty strings for blank lines.
    """
    result = []
    # Detect language once from the full text (not per-line) when auto
    effective_lang = detect_lang(lyrics) if lang == 'auto' else lang

    for lyric_line in lyrics.splitlines():
        if lyric_line.strip():
            rom = line_to_romanisation(lyric_line, lang=effective_lang)
        else:
            rom = ''
        result.append((rom, lyric_line))
    return result
