---
name: lyrics-to-word
description: >
  Paste any song lyrics and generate a beautifully formatted A4 Word document.
  Supports Cantonese Jyutping, Mandarin Pinyin, Japanese Romaji, and Korean Romanization.
  Auto-layout with single/dual column, font size optimization, and live A4 preview.
  Search songs by name to find Genius links, then paste lyrics for formatting.
  Perfect for language learning, karaoke lyrics printing, and music education.
---

# Lyrics to Word Tool — Cloud Edition

A web tool that transforms song lyrics into professionally formatted Word documents.

## What It Does

1. **Search songs** — Find any song by name or artist via Genius search
2. **Paste lyrics** — Copy lyrics from the search result link and paste them in
3. **Auto-layout** — Automatically calculates the best font size and column layout to fit A4
4. **Add romanisation** — Annotate lyrics with Cantonese Jyutping, Mandarin Pinyin, or Japanese Romaji
5. **Generate Word** — Download a print-ready `.docx` file with precise A4 layout

## Key Features

- **Smart auto-layout**: Scans font sizes from 16pt down to 8pt to find the largest that fits one A4 page
- **Two-column support**: Automatically switches to two-column layout for long lyrics
- **Live A4 preview**: See exactly how the document will look before downloading
- **Bilingual UI**: Full Traditional Chinese and English interface
- **Simplified ↔ Traditional conversion**: Auto-converts between Chinese character sets
- **Multiple romanisation systems**: Jyutping (粵語), Pinyin (國語), Romaji (日本語)

## Technical Stack

- **Backend**: Python / Flask
- **Word generation**: python-docx with precise A4 geometry calculations
- **Romanisation**: pycantonese (Jyutping), pypinyin (Pinyin), pykakasi (Romaji)
- **Character conversion**: OpenCC (s2twp / tw2sp)

## Usage

The user describes what song they want formatted, or directly provides lyrics text.
The tool searches for the song, helps them get the lyrics, and generates the Word document.
