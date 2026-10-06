"""Map text in a Frontiers author's proof to its printed margin line numbers."""
import re
import pymupdf


def margin_numbers(page):
    """Return [(side, number, y_center)] for the line numbers printed in the margins."""
    w = page.rect.width
    out = []
    for x0, y0, x1, y1, txt, *_ in page.get_text("words"):
        if not re.fullmatch(r"\d{1,4}", txt):
            continue
        if x1 < 45:
            out.append(("L", int(txt), (y0 + y1) / 2))
        elif x0 > w - 45:
            out.append(("R", int(txt), (y0 + y1) / 2))
    return out


def column_side(page, rect):
    """Which margin numbering applies: the first article page has a left sidebar (L) and one main column (R)."""
    split = 190 if page.number == 4 else page.rect.width / 2
    return "L" if (rect.x0 + rect.x1) / 2 < split else "R"


def line_of(page, rect, nums=None):
    """Printed line number for a rectangle on the page (column chosen by x position)."""
    nums = nums if nums is not None else margin_numbers(page)
    side = column_side(page, rect)
    cands = [(abs(y - (rect.y0 + rect.y1) / 2), n) for s, n, y in nums if s == side]
    if not cands:
        return None
    d, n = min(cands)
    return n if d < 9 else None


def dump(doc, first, last):
    """Yield (page_no, side, line_no, text) for each numbered text line."""
    for pno in range(first - 1, last):
        page = doc[pno]
        nums = margin_numbers(page)
        lines = {}
        for x0, y0, x1, y1, txt, *_ in page.get_text("words"):
            if x1 < 45 or x0 > page.rect.width - 45:
                continue
            r = pymupdf.Rect(x0, y0, x1, y1)
            ln = line_of(page, r, nums)
            side = column_side(page, r)
            key = (side, ln if ln is not None else 10000 + round(y0))
            lines.setdefault(key, []).append((x0, txt))
        for (side, ln), words in sorted(lines.items()):
            yield pno + 1, side, (ln if ln < 10000 else None), " ".join(t for _, t in sorted(words))
