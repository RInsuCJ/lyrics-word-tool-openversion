# 🎵 Lyrics Word Tool

> Turn any song lyrics into a beautifully formatted, print-ready A4 Word document — with romanisation for Cantonese, Mandarin, Japanese & Korean.

<p align="center">
  <strong>🔍 Search → 📝 Paste → 📄 Preview → ⬇️ Download</strong>
</p>

---

## ✨ Features

- **🔍 Song Search** — Find any song via Genius, then copy lyrics from the link
- **📄 Live A4 Preview** — See exactly how your document will look before downloading
- **📐 Smart Auto-Layout** — Automatically picks the best font size and column layout to fit one A4 page
- **📖 Romanisation** — Add pronunciation guides above each line:
  - 🇭🇰 Cantonese Jyutping
  - 🇹🇼 Mandarin Pinyin
  - 🇯🇵 Japanese Romaji
  - 🇰🇷 Korean Romanization
- **🔄 Two-Column Layout** — Automatically switches to dual columns for long lyrics
- **🏷️ Section Tag Toggle** — Show or hide `[Verse]`, `[Chorus]` tags
- **🌐 Bilingual UI** — Full Traditional Chinese & English interface
- **📱 Mobile Friendly** — Works on phones and tablets with tabbed navigation

## 🎯 Who Is This For?

- 🎤 **Language learners** studying through music
- ⛪ **Churches & schools** that need formatted lyric sheets
- 🎵 **Music teachers** creating teaching materials
- 🖨️ **Anyone** who wants clean, printed lyrics

## 🚀 Quick Start

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/lyrics-word-tool.git
cd lyrics-word-tool

# Install dependencies
pip install -r requirements.txt

# Run
python app.py
```

Open **http://localhost:5050** and start searching!

## 📸 How It Works

| Step | Action |
|------|--------|
| **1. Search** | Enter a song title or artist name |
| **2. Copy** | Click the 🔗 link to open Genius, copy the lyrics |
| **3. Paste** | Paste lyrics into the text box, click "Load" |
| **4. Customize** | Adjust font size, columns, add romanisation |
| **5. Download** | One click to generate a perfect A4 Word document |

## 🛠️ Tech Stack

| Component | Technology |
|-----------|-----------|
| Backend | Python / Flask |
| Word Generation | python-docx (precise A4 geometry) |
| Romanisation | pypinyin, pycantonese, pykakasi, korean-romanizer |
| Character Conversion | OpenCC (Simplified ↔ Traditional) |
| Song Search | Genius REST API |
| Frontend | Vanilla HTML/CSS/JS (zero frameworks) |

## 📦 Deployment

This app is designed to run on cloud platforms like [Render](https://render.com):

```bash
# Procfile is included
web: gunicorn app:app
```

> **Note:** Unlike other lyrics tools, this app does **not** scrape webpages — it only uses the Genius API for search, which works reliably on cloud servers without being blocked by Cloudflare.

## 🤝 Contributing

Contributions are welcome! Some ideas:

- [ ] Add more romanisation systems (Thai, Vietnamese, etc.)
- [ ] PDF export option
- [ ] Batch processing for multiple songs
- [ ] Custom fonts upload

## 📄 License

MIT License — feel free to use, modify, and distribute.

---

<p align="center">
  Made with ❤️ for music lovers and language learners
</p>
