"""
app.py — Flask web server for the Lyrics-to-Word tool (Cloud Edition).

This version does NOT scrape lyrics from Genius webpages.
Instead, it provides Genius search (via API) and lets users paste lyrics.

Run:    python app.py
Open:   http://localhost:5000
"""

import os
import tempfile
from flask import (Flask, render_template, request,
                   jsonify, send_file, after_this_request)

from lyrics_tool.search   import search_songs, clean_lyrics, to_traditional
from lyrics_tool.word_gen import (generate_word, find_optimal_layout,
                                  DEFAULT_FONT, DEFAULT_TITLE_PT)

app = Flask(__name__)

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


# ── API: search (Genius REST API — works on cloud servers) ────────────────────

@app.route('/api/search', methods=['POST'])
def api_search():
    data  = request.get_json(force=True)
    query = data.get('query', '').strip()
    if not query:
        return jsonify({'error': '請輸入搜尋關鍵字'}), 400
    try:
        results = search_songs(query, max_results=8)
        return jsonify({'results': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── API: clean pasted lyrics ──────────────────────────────────────────────────

@app.route('/api/clean', methods=['POST'])
def api_clean():
    """Clean user-pasted lyrics: remove section tags, convert to Traditional."""
    data   = request.get_json(force=True)
    raw    = data.get('lyrics', '')
    if not raw.strip():
        return jsonify({'error': '請貼上歌詞內容'}), 400
    cleaned = clean_lyrics(raw)
    cleaned = to_traditional(cleaned)
    return jsonify({'lyrics': cleaned})


# ── API: auto-layout suggestion ───────────────────────────────────────────────

@app.route('/api/layout', methods=['POST'])
def api_layout():
    data     = request.get_json(force=True)
    lyrics   = data.get('lyrics', '')
    title_pt = float(data.get('title_pt', DEFAULT_TITLE_PT))
    body_pt, two_cols = find_optimal_layout(lyrics, title_pt=title_pt)
    return jsonify({'body_pt': body_pt, 'two_cols': two_cols})


# ── API: romanise lyrics ──────────────────────────────────────────────────────

@app.route('/api/romanise', methods=['POST'])
def api_romanise():
    data   = request.get_json(force=True)
    lyrics = data.get('lyrics', '')
    lang   = data.get('lang', 'auto')
    try:
        from lyrics_tool.pinyin import annotate_lyrics
        pairs = annotate_lyrics(lyrics, lang=lang)
        return jsonify({'annotated': pairs})   # list of [rom, lyric]
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── API: generate Word document ───────────────────────────────────────────────

@app.route('/api/generate', methods=['POST'])
def api_generate():
    data      = request.get_json(force=True)
    title     = data.get('title',     '')
    artist    = data.get('artist',    '')
    lyrics    = data.get('lyrics',    '')
    font_name   = data.get('font_name', DEFAULT_FONT)
    title_pt    = float(data.get('title_pt', DEFAULT_TITLE_PT))
    body_pt     = (float(data['body_pt'])
                   if data.get('body_pt') is not None else None)
    two_cols    = (bool(data['two_cols'])
                   if data.get('two_cols') is not None else None)
    with_pinyin    = bool(data.get('with_pinyin', False))
    pinyin_lang    = data.get('pinyin_lang', 'cantonese')
    pinyin_swapped = bool(data.get('pinyin_swapped', False))

    safe = lambda s: ''.join(c for c in s if c not in r'\/:*?"<>|')
    if artist:
        filename = f'《{safe(title)}》{safe(artist)}.docx'
    else:
        filename = f'《{safe(title)}》.docx'

    fd, tmp_path = tempfile.mkstemp(suffix='.docx')
    os.close(fd)

    try:
        generate_word(
            output_path    = tmp_path,
            title          = title,
            artist         = artist,
            lyrics         = lyrics,
            font_name      = font_name,
            title_size     = title_pt,
            force_body_pt  = body_pt,
            force_two_cols = two_cols,
            with_pinyin    = with_pinyin,
            pinyin_lang    = pinyin_lang,
            pinyin_swapped = pinyin_swapped,
        )
    except Exception as e:
        os.unlink(tmp_path)
        return jsonify({'error': str(e)}), 500

    @after_this_request
    def _cleanup(response):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return response

    return send_file(
        tmp_path,
        as_attachment=True,
        download_name=filename,
        mimetype=('application/vnd.openxmlformats-officedocument'
                  '.wordprocessingml.document'),
    )


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    import socket
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = '127.0.0.1'

    print()
    print('  [*] Lyrics Word Tool - Cloud Edition')
    print('  ----------------------------------------')
    print(f'  Local:      http://localhost:{port}')
    print(f'  Same WiFi:  http://{ip}:{port}')
    print('  (Press Ctrl+C to stop)')
    print()
    app.run(host='0.0.0.0', port=port, debug=False)
