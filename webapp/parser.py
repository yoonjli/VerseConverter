import re

REF_PATTERN = re.compile(r'^.{1,40}\s+\d+:\d+\s*$')


def extract_verse_num(ref: str) -> str:
    """Extract chapter:verse digits from a reference string."""
    m = re.search(r'(\d+:\d+)', ref)
    return m.group(1) if m else ref


def parse_verses(text: str) -> list[tuple[str, str]]:
    """
    Parse scripture text into (reference, body) tuples.

    Handles:
    - Layout A: text before reference line
    - Split verses: same reference repeated across multiple chunks → merged
    - Trailing text after last reference → appended to last verse
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    segments: list[list] = []   # [[ref, text], ...]
    pending: list[str] = []

    for line in lines:
        if REF_PATTERN.match(line):
            segments.append([line.strip(), " ".join(pending).strip()])
            pending = []
        else:
            pending.append(line)

    # Trailing text after last reference
    if pending and segments:
        trailing = " ".join(pending).strip()
        if trailing:
            segments[-1][1] = (segments[-1][1] + " " + trailing).strip()

    # Drop empty text segments
    segments = [[r, t] for r, t in segments if t]

    # Merge consecutive segments sharing the same verse number
    merged: list[list] = []
    for ref, text in segments:
        num = extract_verse_num(ref)
        if merged and extract_verse_num(merged[-1][0]) == num:
            merged[-1][1] = merged[-1][1] + " " + text
        else:
            merged.append([ref, text])

    return [(ref, text.strip()) for ref, text in merged]


def match_verses(ko_verses, en_verses):
    """
    Match Korean and English verses by chapter:verse number.
    Returns:
        matched  - list of dicts with ref_ko, text_ko, ref_en, text_en, num
        only_ko  - verse numbers present only in Korean
        only_en  - verse numbers present only in English
    """
    ko_dict = {extract_verse_num(r): (r, t) for r, t in ko_verses}
    en_dict = {extract_verse_num(r): (r, t) for r, t in en_verses}

    all_nums = sorted(
        ko_dict.keys() | en_dict.keys(),
        key=lambda x: list(map(int, x.split(':')))
    )

    matched, only_ko, only_en = [], [], []

    for num in all_nums:
        if num in ko_dict and num in en_dict:
            ref_ko, text_ko = ko_dict[num]
            ref_en, text_en = en_dict[num]
            matched.append({
                "num": num,
                "ref_ko": ref_ko,
                "text_ko": text_ko,
                "ref_en": ref_en,
                "text_en": text_en,
            })
        elif num in ko_dict:
            only_ko.append(num)
        else:
            only_en.append(num)

    return matched, only_ko, only_en
