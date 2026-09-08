"""
Fixes stray spaces between a Korean proper noun and the particle (조사)
that should be attached directly to it, e.g. "바울 은" -> "바울은",
"마케도니아 로" -> "마케도니아로".

Safe to run over a whole bilingual (KO+EN) file: the pattern only matches
runs of Hangul, so English lines pass through untouched.
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
    "에서", "에게", "에는", "께서", "한테", "한테서", "이랑"
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


def fix_josa_spacing(text: str) -> str:
    """Collapse '<noun> <josa>' into '<noun><josa>'. Loops until stable
    to catch rare back-to-back cases like 'A 은 B 는' on one line."""
    prev = None
    while prev != text:
        prev = text
        text = _PATTERN.sub(r"\1\2", text)
    return text
