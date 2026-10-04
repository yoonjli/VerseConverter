"""
Fixes stray spaces between a Korean proper noun and the particle (조사)
that should be attached directly to it, e.g. "바울 은" -> "바울은",
"마케도니아 로" -> "마케도니아로".

Safe to run over a whole bilingual (KO+EN) file: the pattern only matches
runs of Hangul, so English lines pass through untouched.

Special case: "이" is both a subject particle ("바울 이" -> "바울이") and the
demonstrative "this" ("이 종"). If the previous word already ends in a
particle (e.g. "가다가는 이 재난"), the "이" must be the demonstrative, so the
space is kept.
"""

import re

# Particles / copula-endings, longest first so e.g. "에서는" wins over "에서".
JOSA = [
    "에게서", "으로부터", "에게로", "로부터",
    "이라고는", "이라고", "이라는", "이었다고", "이었다",
    "였다고", "였다는", "였다",
    "이다", "입니다", "이며", "이나", "이라", "이야", "이여",
    "라는", "라고", "라며", "라야", "같이",
    "에서는", "에서도", "에게는", "에게도", "으로는", "으로도",
    "까지는", "까지도", "부터는",
    "에서", "에게", "에는", "께서", "한테", "한테서", "이랑",  # (comma was missing after "이랑")
    "처럼", "만큼", "마다", "조차", "마저", "밖에", "뿐",
    "으로", "로서", "로써", "까지", "부터", "하고", "보다",
    "은", "는", "이", "가", "을", "를", "의", "에", "로",
    "와", "과", "도", "만", "께", "야", "여", "나", "씩",
]

_JOSA_PATTERN = "|".join(re.escape(j) for j in sorted(set(JOSA), key=len, reverse=True))

# group1: a run of Hangul syllables (the noun/name)
# group2: one of the particles above, as a *whole* token (checked via lookahead)
_PATTERN = re.compile(
    r"([가-힣]+)[ \t]+(" + _JOSA_PATTERN + r")(?=[\s,\.!?\"'\u201c\u201d\u2018\u2019\u2033)\]]|$)"
)

# If the previous word ends with one of these, it already carries a particle,
# so a following "이" is the demonstrative "this", not a particle.
# (Rare vocative endings 야/여/나/씩 are left out on purpose: too many names
# such as 이사야, 요나 end in them.)
_PARTICLE_ENDINGS = (
    "에서", "에게", "까지", "부터", "께서", "한테", "처럼", "보다", "마다",
    "조차", "마저", "하고",
    "은", "는", "을", "를", "가", "도", "만", "에", "로", "와", "과", "의", "께",
)

# Nouns/names that merely END in a particle-like syllable. They do NOT already
# carry a particle, so a following "이" is a real subject particle:
#   "사도 이" -> "사도이",  "교만 이" -> "교만이".
# Exact-match only (not endswith), otherwise "서울에서" would be exempted too.
# Add more here whenever you spot a miss.
NOUNS_ENDING_LIKE_JOSA = {
    "사도", "기도", "제도", "전도", "인도", "지도",   # ...도
    "교만",                                           # ...만
    "여호와", "하와",                                 # ...와
    "마가", "누가", "아가",                           # ...가
    "에서",                                           # Esau (…에서)
    "마을", "고을",                                   # ...을
}


def _already_has_particle(word: str) -> bool:
    return word.endswith(_PARTICLE_ENDINGS) and word not in NOUNS_ENDING_LIKE_JOSA


def _join(m: "re.Match") -> str:
    word, josa = m.group(1), m.group(2)
    if josa == "이" and _already_has_particle(word):
        return m.group(0)          # demonstrative "this": keep the space
    return word + josa


def fix_josa_spacing(text: str) -> str:
    """Collapse '<noun> <josa>' into '<noun><josa>'. Loops until stable
    to catch rare back-to-back cases like 'A 은 B 는' on one line."""
    prev = None
    while prev != text:
        prev = text
        text = _PATTERN.sub(_join, text)
    return text
