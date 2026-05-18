"""Presidential order number conversion.

Converts Arabic numerals in presidential order numbers to Chinese uppercase.
E.g., 第81号 → 第八十一号, 主席令第81号 → 主席令第八十一号.
"""
import re

# Patterns for presidential order numbers (before conversion)
# Covers: 第81号, 主席令第81号, 中华人民共和国主席令第81号
_PRESIDENTIAL_RE = re.compile(
    r'(?:(?:中华人民共和国)?主席令)?第(\d+)号'
)

# Already-converted pattern for validation
_PRESIDENTIAL_CN_RE = re.compile(
    r'(?:(?:中华人民共和国)?主席令)?第[一二三四五六七八九十百千万]+号'
)


def number_to_chinese(num: int) -> str:
    """Convert integer 1-9999 to Chinese numeral (e.g., 81 → 八十一)."""
    if not 1 <= num <= 9999:
        return str(num)

    digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
    units = ['', '十', '百', '千']

    s = str(num)
    n = len(s)
    result = []

    for i, ch in enumerate(s):
        d = int(ch)
        pos = n - i - 1  # 0-based from right
        if d == 0:
            # Skip consecutive zeros, but keep one 零 if not trailing
            if result and result[-1] != '零' and i < n - 1:
                result.append('零')
        else:
            result.append(digits[d] + units[pos])

    # Clean up: remove leading 一 for tens (一十 → 十)
    full = ''.join(result)
    if full.startswith('一十') and n == 2:
        full = full[1:]

    # Remove trailing 零
    full = full.rstrip('零')

    return full


def format_presidential_order_numbers(doc, *, _paragraphs=None) -> int:
    """Scan document for presidential order numbers and convert Arabic to Chinese.

    Returns number of conversions made.
    """
    paragraphs = _paragraphs if _paragraphs is not None else doc.paragraphs
    converted = 0

    for para in paragraphs:
        # Only modify paragraph text if it contains a presidential order pattern
        text = para.text
        if not text or '第' not in text or '号' not in text:
            continue

        new_text, count = _PRESIDENTIAL_RE.subn(_replace_num, text)
        if count > 0 and new_text != text:
            # Replace text in the first run (typical for this pattern)
            for run in para.runs:
                if _PRESIDENTIAL_RE.search(run.text):
                    run.text = _PRESIDENTIAL_RE.sub(_replace_num, run.text)
                    converted += count
                    break
            else:
                # Fallback: modify first text run
                for run in para.runs:
                    if run.text.strip():
                        old = run.text
                        run.text = _PRESIDENTIAL_RE.sub(_replace_num, old)
                        if run.text != old:
                            converted += count
                            break

    return converted


def _replace_num(m: re.Match) -> str:
    """Regex replacement callback: replace digit group with Chinese numeral."""
    num_str = m.group(1)
    num = int(num_str)
    cn = number_to_chinese(num)
    # Rebuild full match with Chinese numeral
    full = m.group(0)
    return full.replace(num_str, cn)


def check_presidential_order_numbers(paragraphs, issues: list) -> None:
    """Validate: flag any presidential order numbers still using Arabic numerals."""
    from tvba_core_validate import ValidationIssue

    for para in paragraphs:
        text = para.text
        if not text or '第' not in text or '号' not in text:
            continue

        for m in _PRESIDENTIAL_RE.finditer(text):
            issues.append(ValidationIssue(
                severity="warning",
                description=f"主席令编号应为中文大写（{m.group(0)}）",
                location=f'"{text[:50]}..."' if len(text) > 50 else f'"{text}"',
            ))
