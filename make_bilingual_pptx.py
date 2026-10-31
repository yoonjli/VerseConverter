import sys
import re
from fix_josa import fix_josa_spacing
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Usage ──────────────────────────────────────────────────────────────
# python make_bilingual_pptx.py combined.txt [output.pptx]
#
# combined.txt holds BOTH languages in one file, back to back:
#   - all the Korean verses first (text, then its "BookName 26:1" reference)
#   - then all the English verses (text, then its "BookName 26:1" reference)
# This is exactly the same "text before reference" layout as before, just
# with the Korean block and the English block concatenated into one file.
# The script finds the Korean → English boundary automatically by looking
# for the first reference line that contains no Hangul characters.
# ───────────────────────────────────────────────────────────────────────

REF_PATTERN = re.compile(r'^.{1,40}\s+\d+:\d+\s*$')


def contains_hangul(s):
    return any('\uac00' <= ch <= '\ud7a3' for ch in s)


def extract_verse_num(ref):
    """Extract just the chapter:verse digits (e.g. '26:7') from a reference string."""
    m = re.search(r'(\d+:\d+)', ref)
    return m.group(1) if m else ref


def book_name(ref):
    """'1 Corinthians 2:1' -> '1 Corinthians'"""
    return re.sub(r'\s*\d+:\d+\s*$', '', ref).strip()


def parse_verses_from_lines(lines):
    """
    Parse stripped, non-blank lines into a list of (reference, text, key).

    Works with ANY mix of layouts, even inside one file:
        text / reference      (reference AFTER its text)
        reference / text      (reference BEFORE its text)

    How: lines are grouped into alternating items - REF lines and TEXT chunks
    (consecutive text lines = one chunk, so a verse broken across lines is
    fine). Every reference must own exactly one adjacent text chunk, and
    scanning left to right there is only one way to do that: the first item
    pairs with its right neighbour, then the next unpaired item, and so on.

    `key` = (book_number, chapter, verse). The book number counts how many
    times the book name changed, so "Matthew 2:1" and "1 Corinthians 2:1"
    no longer collide, and Korean/English books line up by ORDER even though
    their names differ.

    Split verses (same reference appearing more than once in a row) are
    merged into one entry.
    """
    # --- Build items: ("R", ref_line) or ("T", [text lines]) ---
    items = []
    for line in lines:
        if line.lower().startswith("title:"):      # file header, not a verse
            continue
        if REF_PATTERN.match(line):
            items.append(["R", line.strip()])
        elif items and items[-1][0] == "T":
            items[-1][1].append(line)
        else:
            items.append(["T", [line]])

    # --- Pair each reference with its adjacent text chunk ---
    segments = []          # [ref, text]
    i = 0
    while i < len(items):
        if i + 1 < len(items) and items[i][0] != items[i + 1][0]:
            ref_item, text_item = (items[i], items[i + 1]) if items[i][0] == "R" \
                                  else (items[i + 1], items[i])
            segments.append([ref_item[1], " ".join(text_item[1]).strip()])
            i += 2
        else:
            what = "reference" if items[i][0] == "R" else "text"
            shown = items[i][1] if items[i][0] == "R" else " ".join(items[i][1])[:40]
            print(f"  ✗  Could not pair this {what} with a verse (skipped): {shown!r}")
            i += 1

    # --- Merge consecutive segments with the same book + verse number ---
    merged = []
    for ref, text in segments:
        if merged and merged[-1][0] == ref:
            merged[-1][1] += " " + text
        else:
            merged.append([ref, text])

    # --- Assign keys (book order, chapter, verse) ---
    out = []
    book_idx, last_book = 0, None
    for ref, text in merged:
        bk = book_name(ref)
        if last_book is not None and bk != last_book:
            book_idx += 1
        last_book = bk
        ch, vs = map(int, extract_verse_num(ref).split(':'))
        out.append((ref, fix_josa_spacing(text.strip()), (book_idx, ch, vs)))
    return out


def parse_verses(filepath):
    """Read a file and parse it (single-language layout)."""
    with open(filepath, encoding='utf-8') as f:
        lines = [l.strip() for l in f.readlines()]
    lines = [l for l in lines if l]
    return parse_verses_from_lines(lines)


def split_bilingual_file(filepath):
    """
    Read a combined Korean+English file and split it into two line lists
    at the point where reference lines stop containing Hangul.
    """
    with open(filepath, encoding='utf-8') as f:
        lines = [l.strip() for l in f.readlines()]
    lines = [l for l in lines if l and not l.lower().startswith("title:")]

    boundary = None
    for i, line in enumerate(lines):
        if not contains_hangul(line):
            boundary = i
            break

    if boundary is None:
        raise ValueError(
            "Could not find the Korean → English boundary. "
            "Expected the Korean references (e.g. '사도행전 15:1') to be "
            "followed later by English references (e.g. 'Acts 15:1')."
        )

    return lines[:boundary], lines[boundary:]


def fmt_key(key):
    return f"#{key[0] + 1} {key[1]}:{key[2]}"


def match_verses(ko_verses, en_verses):
    """
    Pair Korean and English verses by (book order, chapter, verse).
    Returns list of (ref_ko, text_ko, ref_en, text_en).
    Prints a report of any unmatched verses.
    """
    ko_dict = {key: (ref, text) for ref, text, key in ko_verses}
    en_dict = {key: (ref, text) for ref, text, key in en_verses}

    matched, only_ko, only_en = [], [], []
    for key in sorted(ko_dict.keys() | en_dict.keys()):
        if key in ko_dict and key in en_dict:
            matched.append((*ko_dict[key], *en_dict[key]))
        elif key in ko_dict:
            only_ko.append(ko_dict[key][0])
        else:
            only_en.append(en_dict[key][0])

    if only_ko:
        print(f"  ⚠  Korean-only verses (no English match): {only_ko}")
    if only_en:
        print(f"  ⚠  English-only verses (no Korean match): {only_en}")
    return matched


# ── Colors ─────────────────────────────────────────────────────────────
BLACK      = RGBColor(0x00, 0x00, 0x00)
GOLD       = RGBColor(0xD4, 0xAF, 0x37)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GOLD = RGBColor(0xF0, 0xD9, 0x80)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── Overflow estimation ────────────────────────────────────────────────
# python-pptx cannot measure rendered text, so we ESTIMATE how many lines a
# verse needs and compare with how many lines fit in its text box. Numbers
# are calibrated against a real render (36pt, 12.03in-wide box -> 4 Korean
# lines fit; a 5th line spills over the gold divider).
BODY_FONT_PT   = 36
MARGIN         = Inches(0.65)
BOX_H_IN       = 2.9          # height of each Korean / English text box
BOX_INSET_H_IN = 0.1          # PowerPoint default left/right text inset
BOX_INSET_V_IN = 0.05         # PowerPoint default top/bottom text inset
KO_LINE_FACTOR = 1.30         # line height = font size x factor (Korean font)
EN_LINE_FACTOR = 1.20         # (Georgia)
OPENERS        = set("\u201c\u2018([\"'")


def _char_em(ch):
    """Approximate glyph width in em."""
    if '\uac00' <= ch <= '\ud7a3' or '\u3131' <= ch <= '\u318e' or '\u4e00' <= ch <= '\u9fff':
        return 0.95
    if ch == ' ':
        return 0.25
    if ch in "iljtf.,;:'!|\u2019\u2018":
        return 0.30
    if ch in "\u201c\u201d\"":
        return 0.40
    if ch in "mwMW":
        return 0.85
    if ch.isdigit():
        return 0.55
    if ch.isupper():
        return 0.70
    return 0.50


def _units(word):
    """Split a word into unbreakable units: each Hangul syllable is its own
    unit (PowerPoint wraps Korean mid-word); punctuation/Latin sticks to the
    previous unit; an opening quote sticks to the next syllable."""
    units = []
    for ch in word:
        is_hangul = '\uac00' <= ch <= '\ud7a3'
        if not units:
            units.append(ch)
        elif is_hangul:
            if all(c in OPENERS for c in units[-1]):
                units[-1] += ch
            else:
                units.append(ch)
        else:
            units[-1] += ch
    return units


def estimate_line_count(text, font_pt, box_width_in):
    max_w = (box_width_in - 2 * BOX_INSET_H_IN) * 72
    space_w = _char_em(' ') * font_pt
    lines, cur = 1, 0.0
    for wi, word in enumerate(text.split()):
        if wi > 0 and cur > 0:
            cur += space_w
        for unit in _units(word):
            w = sum(_char_em(c) for c in unit) * font_pt
            if cur + w > max_w and cur > 0:
                lines += 1
                cur = w
            else:
                cur += w
    return lines


def max_lines_that_fit(font_pt, line_factor):
    usable = (BOX_H_IN - 2 * BOX_INSET_V_IN) * 72
    return int(usable // (font_pt * line_factor))


MIN_FONT_PT  = 24     # never shrink below this (still readable when projected)
FONT_STEP_PT = 2


def fit_font_size(text, line_factor, start_pt=BODY_FONT_PT, min_pt=MIN_FONT_PT):
    """Largest font size (start_pt, start_pt-2, ... min_pt) at which `text`
    is estimated to fit its box. Returns (size_pt, fits).
    If nothing down to min_pt fits, returns (min_pt, False)."""
    box_w_in = (SLIDE_W - MARGIN * 2) / 914400
    size = start_pt
    while size >= min_pt:
        if estimate_line_count(text, size, box_w_in) <= max_lines_that_fit(size, line_factor):
            return size, True
        size -= FONT_STEP_PT
    return min_pt, False


def add_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                font_size=22, bold=False, color=WHITE,
                align=PP_ALIGN.CENTER, font_name=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name or (
        "RIDIBatangSHL" if any(ord(c) > 127 for c in text) else "Georgia"
    )


def make_slide(prs, ref_ko, korean, ref_en, english):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_bg(slide, BLACK)
    W, H = SLIDE_W, SLIDE_H
    m = MARGIN

    # Top gold bar
    b = slide.shapes.add_shape(1, Inches(0), Inches(0), W, Inches(0.07))
    b.fill.solid(); b.fill.fore_color.rgb = GOLD; b.line.fill.background()

    # Combined reference label
    ref_label = f"{ref_ko}  /  {ref_en}"
    add_textbox(slide, ref_label, m, Inches(0.15), W - m * 2, Inches(0.55),
                font_size=27, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
                font_name="RIDIBatangSHL")

    # Pick a font size per section: normal size if it fits, otherwise shrink
    # step by step until the estimate says it fits (or MIN_FONT_PT is reached).
    ko_size, ko_fits = fit_font_size(korean, KO_LINE_FACTOR)
    en_size, en_fits = fit_font_size(english, EN_LINE_FACTOR)

    # Korean text (top half)
    add_textbox(slide, korean, m, Inches(0.85), W - m * 2, Inches(2.9),
                font_size=ko_size, color=WHITE, align=PP_ALIGN.CENTER,
                font_name="RIDIBatangSHL")

    # Divider
    div = slide.shapes.add_shape(1, Inches(1.5), Inches(3.85), W - Inches(3.0), Inches(0.04))
    div.fill.solid(); div.fill.fore_color.rgb = GOLD; div.line.fill.background()

    # English text (bottom half)
    add_textbox(slide, english, m, Inches(4.0), W - m * 2, Inches(2.9),
                font_size=en_size, color=LIGHT_GOLD, align=PP_ALIGN.CENTER,
                font_name="Georgia")

    # Bottom gold bar
    b2 = slide.shapes.add_shape(1, Inches(0), H - Inches(0.07), W, Inches(0.07))
    b2.fill.solid(); b2.fill.fore_color.rgb = GOLD; b2.line.fill.background()

    # Record (and print) every adjustment so the user can review those slides.
    # Printing means the GUI's on-screen log shows it too.
    slide_no = len(prs.slides)
    adjustments = []
    for section, size, fits in (("Korean", ko_size, ko_fits), ("English", en_size, en_fits)):
        if size != BODY_FONT_PT or not fits:
            adjustments.append({
                "slide": slide_no, "ref_ko": ref_ko, "ref_en": ref_en,
                "section": section, "from_pt": BODY_FONT_PT, "to_pt": size,
                "still_overflows": not fits,
            })
            if fits:
                print(f"  ⚠  Slide {slide_no} ({ref_ko} / {ref_en}): {section} font shrunk "
                      f"{BODY_FONT_PT}pt → {size}pt to fit - please review")
            else:
                print(f"  ✗  Slide {slide_no} ({ref_ko} / {ref_en}): {section} text may STILL "
                      f"overflow at the minimum {size}pt - please edit or split this verse")
    return adjustments


def print_review_summary(adjustments):
    """Print the end-of-run list of slides the user should open and confirm.
    `adjustments` is the combined list returned by make_slide() calls."""
    if not adjustments:
        print("\n✓ No font adjustments were needed (estimate - a quick visual check is still wise).")
        return
    shrunk = [a for a in adjustments if not a["still_overflows"]]
    stuck  = [a for a in adjustments if a["still_overflows"]]
    print("\n" + "=" * 64)
    print("REVIEW NEEDED - please open these slides and confirm they look right")
    print("=" * 64)
    if shrunk:
        print(f"\nFont was SHRUNK on {len({a['slide'] for a in shrunk})} slide(s):")
        for a in shrunk:
            print(f"  • Slide {a['slide']:>3}  {a['ref_ko']} / {a['ref_en']}  "
                  f"[{a['section']}]  {a['from_pt']}pt → {a['to_pt']}pt")
    if stuck:
        print(f"\nStill may OVERFLOW even at {MIN_FONT_PT}pt on {len({a['slide'] for a in stuck})} slide(s) "
              f"(edit text or split the verse):")
        for a in stuck:
            print(f"  • Slide {a['slide']:>3}  {a['ref_ko']} / {a['ref_en']}  [{a['section']}]")
    print("\nNote: sizes are estimated, not measured. Slides not listed should still get a quick look.")


def main():
    if len(sys.argv) < 2:
        print("Usage: python make_bilingual_pptx.py combined.txt [output.pptx]")
        sys.exit(1)

    in_file  = sys.argv[1]
    out_file = sys.argv[2] if len(sys.argv) > 2 else "bilingual_scripture.pptx"

    print(f"Parsing {in_file} ...")
    ko_lines, en_lines = split_bilingual_file(in_file)

    ko_verses = parse_verses_from_lines(ko_lines)
    print(f"  → {len(ko_verses)} Korean verses after merging")

    en_verses = parse_verses_from_lines(en_lines)
    print(f"  → {len(en_verses)} English verses after merging")

    print("Matching verses by number...")
    pairs = match_verses(ko_verses, en_verses)
    print(f"  → {len(pairs)} matched pairs")

    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    adjustments = []
    for ref_ko, text_ko, ref_en, text_en in pairs:
        adjustments.extend(make_slide(prs, ref_ko, text_ko, ref_en, text_en) or [])

    prs.save(out_file)
    print(f"✓ Saved {len(pairs)} slides → {out_file}")
    print_review_summary(adjustments)


if __name__ == "__main__":
    main()
