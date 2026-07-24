# VerseConverter

Generate bilingual Korean/English PowerPoint slides from plain text scripture input — built for sermon and Bible study presentations.

---

## Overview

VerseConverter takes a single plain-text file containing both Korean and English scripture and outputs a formatted `.pptx` slide deck with one slide per verse. The Korean → English boundary is detected automatically, and verses are matched by chapter:verse number and laid out on the same slide with distinct styling.

---

## Usage

```bash
python make_bilingual_pptx.py combined.txt [output.pptx]
```

| Argument | Description |
|---|---|
| `combined.txt` | Single text file containing both Korean and English scripture, back to back |
| `output.pptx` | *(optional)* Output filename. Defaults to `bilingual_scripture.pptx` |

---

## Input Format

The file holds **both languages in one place**: all the Korean verses first, followed by all the English verses. Within each language block, verse text still appears *before* its reference line (**Layout A**), and the reference must end with a `chapter:verse` number (e.g. `26:1`).

```
태초에 하나님이 천지를 창조하시니라
창세기 1:1

땅이 혼돈하고 공허하며 흑암이 깊음 위에 있고
창세기 1:2

In the beginning God created the heavens and the earth.
Genesis 1:1

Now the earth was formless and empty, darkness was over the surface of the deep.
Genesis 1:2
```

**Notes:**
- Blank lines between verses are ignored
- Long verses split across multiple lines are automatically joined
- Trailing text after the final reference is appended to the last verse
- Verses with no match in the other language are reported as warnings but do not cause errors
- The Korean → English boundary is detected automatically as the first line containing no Hangul characters — no marker or separator is needed between the two blocks

---

## Matching & Merging Logic

1. **Split** — The combined file is divided into a Korean block and an English block at the first line containing no Hangul characters.
2. **Parse** — Each block is scanned for lines matching `BookName chapter:verse`. Text accumulated before each reference is assigned to that verse.
3. **Merge duplicates** — Consecutive entries sharing the same `chapter:verse` number are joined into a single verse (handles source files where one verse spans multiple blocks).
4. **Match** — Korean and English verses are paired by `chapter:verse` number. Unmatched verses are printed as warnings (`⚠ Korean-only` / `⚠ English-only`).

---

## Slide Layout

Each slide follows this structure (widescreen 13.33" × 7.5"):

```
┌─────────────────────────────────────────┐
│▓▓▓▓▓▓▓▓▓▓  gold top bar  ▓▓▓▓▓▓▓▓▓▓▓▓▓│
│                                         │
│       창세기 1:1  /  Genesis 1:1        │  ← gold, 27pt, Malgun Gothic
│                                         │
│         Korean verse text here          │  ← white, 36pt, Malgun Gothic
│                                         │
│          ───────────────                │  ← gold divider
│                                         │
│         English verse text here         │  ← light gold, 36pt, Georgia
│                                         │
│▓▓▓▓▓▓▓▓▓▓  gold bottom bar  ▓▓▓▓▓▓▓▓▓▓│
└─────────────────────────────────────────┘
```

### Colors

| Element | Color | Hex |
|---|---|---|
| Background | Black | `#000000` |
| Top/bottom bars, divider, reference | Gold | `#D4AF37` |
| Korean body text | White | `#FFFFFF` |
| English body text | Light gold | `#F0D980` |

### Fonts

| Text | Font |
|---|---|
| Korean (reference + body) | Malgun Gothic |
| English body | Georgia |

> Font selection is automatic: any text containing non-ASCII characters uses Malgun Gothic; pure ASCII text uses Georgia.

---

## Requirements

```
python-pptx>=0.6.21
```

Install with:

```bash
pip install python-pptx
```

---

## Key Functions

| Function | Description |
|---|---|
| `split_bilingual_file(filepath)` | Reads the combined `.txt` file and splits it into Korean and English line lists at the Hangul boundary |
| `parse_verses_from_lines(lines)` | Parses a list of lines (one language block); returns a list of `(reference, text)` tuples |
| `parse_verses(filepath)` | Reads and parses a single-language `.txt` file directly (kept for standalone use) |
| `extract_verse_num(ref)` | Extracts the bare `chapter:verse` digits from a full reference string |
| `match_verses(ko_verses, en_verses)` | Pairs KO and EN verse lists by `chapter:verse`; reports unmatched entries |
| `make_slide(prs, ref_ko, korean, ref_en, english)` | Adds one bilingual slide to the presentation |
| `add_textbox(slide, ...)` | Helper to add a styled text box with font, size, color, and alignment |
| `add_bg(slide, color)` | Sets the slide background to a solid color |

---

## Example

```bash
python make_bilingual_pptx.py genesis11.txt genesis11.pptx
```

```
Parsing genesis11.txt ...
  → 9 Korean verses after merging
  → 9 English verses after merging
Matching verses by number...
  → 9 matched pairs
✓ Saved 9 slides → genesis11.pptx
```
