"""
search.py — Search songs via Genius REST API (cloud-safe, no webpage scraping).

This module ONLY uses the Genius API endpoints (bearer-token auth),
which are NOT blocked by Cloudflare on cloud servers.  It does NOT
fetch lyrics from Genius webpages — users paste lyrics themselves.
"""

import re
import requests

import os as _os

# ── Genius API token ──────────────────────────────────────────────────────────
GENIUS_TOKEN = _os.environ.get(
    'GENIUS_ACCESS_TOKEN',
    'n6Px9y2EysJlD7D66OrmrgtyYAIqRPB3iO33PoS5OZYA5ma0fGJTW1QvzOb3Fx_9'
)

GENIUS_API_BASE = 'https://api.genius.com'

# ── Simplified ↔ Traditional Chinese converters ─────────────────────────────
try:
    import opencc as _opencc
    _CC_S2T = _opencc.OpenCC('s2twp')  # Simplified → Traditional (Taiwan)
    _CC_T2S = _opencc.OpenCC('tw2sp')  # Traditional (Taiwan) → Simplified

    def to_traditional(text: str) -> str:
        """Simplified → Traditional Chinese (for output)."""
        return _CC_S2T.convert(text) if text else text

    def to_simplified(text: str) -> str:
        """Traditional → Simplified Chinese (for Genius search queries)."""
        return _CC_T2S.convert(text) if text else text

except Exception:
    def to_traditional(text: str) -> str:   # type: ignore[misc]
        return text

    def to_simplified(text: str) -> str:    # type: ignore[misc]
        return text


# ── Section-tag regex (for cleaning user-pasted lyrics) ───────────────────────
SECTION_TAG_RE = re.compile(r'^\[.*?\]\s*$', re.MULTILINE)


def search_songs(query: str, max_results: int = 8):
    """
    Search Genius for songs matching the query via REST API.

    This endpoint uses api.genius.com (bearer auth) and is NOT affected
    by Cloudflare bot detection — works reliably on cloud servers.

    Returns a list of dicts:
        [{"title": ..., "artist": ..., "url": ..., "id": ...}, ...]
    """
    # Convert query: Traditional → Simplified so Genius can match
    genius_query = to_simplified(query)

    headers = {'Authorization': f'Bearer {GENIUS_TOKEN}'}
    params = {'q': genius_query, 'per_page': max_results}

    try:
        resp = requests.get(
            f'{GENIUS_API_BASE}/search',
            headers=headers,
            params=params,
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        raise RuntimeError(f"搜尋失敗：{e}")

    hits = []
    response = data.get('response', {})
    for hit in response.get('hits', []):
        if hit.get('type') == 'song':
            r = hit.get('result', {})
            hits.append({
                'id': r.get('id'),
                'title': r.get('title', ''),
                'artist': r.get('primary_artist', {}).get('name', ''),
                'url': r.get('url', ''),           # Genius webpage URL
                'full_title': r.get('full_title', ''),
            })
    return hits


def clean_lyrics(raw: str) -> str:
    """
    Clean user-pasted lyrics.

    Removes:
      - "Embed" / "123Embed" suffixes from Genius copy-paste
      - Excessive blank lines (3+ → 2)

    Note: Section tags like [Verse 1] are preserved so the frontend
    can toggle their visibility via a user control.
    """
    cleaned = raw

    # Remove the "Embed" or contributor suffix that Genius sometimes appends
    cleaned = re.sub(r'\d+Embed$', '', cleaned)
    cleaned = re.sub(r'Embed$', '', cleaned)

    # Collapse 3+ consecutive blank lines into 2
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)

    return cleaned.strip()
