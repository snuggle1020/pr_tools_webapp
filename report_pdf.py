# -*- coding: utf-8 -*-
"""
결과보고서(보도자료 배포 결과) PDF 빌더.
싱클리_K-커머스서밋2026_보도결과 PDF 양식을 그대로 따른다:
  1쪽 - 상단 요약줄/제목, 핵심 수치 카드(+기사 한 건당 막대 하나), 주요 매체 게재 카드,
        포털 노출 구분 / 매체 유형 막대
  2쪽~ - 게재 기사 전체 표 (제목 클릭 시 원문 이동)

데이터는 report_sheet.resolve_coverage() 결과(엑셀과 동일)를 받는다.
포털은 네이버 CP(인링크) / 네이버 검색제휴 / 다음 / 구글 순으로 구분한다.
다음·구글은 API가 없어 담당자가 앱에서 수동 입력한 값(report_sheet.apply_manual_portals)을 쓴다.
한글 폰트는 fonts/ 의 Pretendard(OFL)를 PDF에 임베드한다.
"""
import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

from media_order import CATEGORY_ORDER, media_category, media_sort_key

FONT_DIR = Path(__file__).parent / "fonts"
REG, MED, SEMI, BOLD = "Pretendard", "Pretendard-Medium", "Pretendard-SemiBold", "Pretendard-Bold"

INK = HexColor("#16191F")
MUTED = HexColor("#5A6170")
SUBTITLE = HexColor("#363B47")  # 보고서명 (기업명보다 한 톤 연하게)
LINE = HexColor("#E3E6EB")
BLUE = HexColor("#1F4ED8")      # 네이버 CP / 강조
LIGHT_BLUE = HexColor("#7EA0F0")  # 네이버 검색제휴
GREEN = HexColor("#2EA279")       # 다음
GRAY = HexColor("#C8CDD8")
ORANGE = HexColor("#B45309")     # [확인필요] 표시
PILL_BG = HexColor("#EEF0F4")    # 묶음 기사 표시 바탕
WHITE = HexColor("#FFFFFF")
WEEKDAYS = "월화수목금토일"

PAGE_W, PAGE_H = A4
X0 = 34
W = PAGE_W - 2 * X0
TOP = 42
BOTTOM = 40

PORTALS = [("네이버 CP", BLUE), ("네이버 검색제휴", LIGHT_BLUE), ("다음", GREEN), ("구글", GRAY)]
PORTAL_COLOR = dict(PORTALS)
PORTAL_RANK = {name: i for i, (name, _) in enumerate(PORTALS)}
REVIEW_PREFIX = "[확인필요] "
MAX_FEATURED = 6

# Pretendard에는 한자(韓·美·中 등)가 없어서, 없는 글자만 Noto Sans KR로 찍는다.
# fonts/NotoSansKR-Hanja-*.ttf 는 Noto Sans KR(OFL) 가변폰트를 굵기 고정 + 한자 영역만 남겨 줄인 파일.
HANJA, HANJA_BOLD = "NotoSansKR-Hanja-Regular", "NotoSansKR-Hanja-Bold"
FALLBACK = {REG: HANJA, MED: HANJA, SEMI: HANJA_BOLD, BOLD: HANJA_BOLD}

_fonts_registered = False
_cmaps = {}


def _register_fonts():
    global _fonts_registered
    if _fonts_registered:
        return
    files = [(REG, "Pretendard-Regular"), (MED, "Pretendard-Medium"), (SEMI, "Pretendard-SemiBold"),
             (BOLD, "Pretendard-Bold"), (HANJA, HANJA), (HANJA_BOLD, HANJA_BOLD)]
    for name, file in files:
        font = TTFont(name, str(FONT_DIR / f"{file}.ttf"))
        pdfmetrics.registerFont(font)
        _cmaps[name] = font.face.charToGlyph
    _fonts_registered = True


def _runs(text, font):
    """(문자열, 폰트) 조각 목록. 기본 폰트에 없는 글자(한자 등)는 보조 폰트로 찍고,
    둘 다 없는 글자나 깨진 문자(U+FFFD, 원문 인코딩 오류)는 공백으로 바꾼다 ("출범���SAP" -> "출범 SAP")."""
    fb = FALLBACK.get(font, font)
    main, sub = _cmaps[font], _cmaps[fb]
    text = "".join(ch if ch == " " or ord(ch) in main or ord(ch) in sub else " " for ch in text)
    text = re.sub(" {2,}", " ", text)
    runs = []
    for ch in text:
        f = font if ch == " " or ord(ch) in main else fb
        if runs and runs[-1][1] == f:
            runs[-1][0] += ch
        else:
            runs.append([ch, f])
    return runs


def _sw(text, font, size):
    return sum(pdfmetrics.stringWidth(t, f, size) for t, f in _runs(text, font))


def _wrap(text, font, size, width):
    """단어 단위로 줄바꿈하되, 한 단어가 너무 길면 글자 단위로 자른다."""
    lines, cur = [], ""
    for word in text.split(" "):
        cand = word if not cur else cur + " " + word
        if _sw(cand, font, size) <= width:
            cur = cand
            continue
        if cur:
            lines.append(cur)
        cur = ""
        for ch in word:
            if _sw(cur + ch, font, size) > width and cur:
                lines.append(cur)
                cur = ch
            else:
                cur += ch
    if cur:
        lines.append(cur)
    return lines or [""]


class _Page:
    """위에서부터 내려오는 좌표(top)로 그리기 위한 얇은 래퍼."""

    def __init__(self, c):
        self.c = c

    def y(self, top):
        return PAGE_H - top

    def text(self, x, top, s, font=REG, size=9.8, color=INK, align="left"):
        c = self.c
        c.setFillColor(color)
        y = self.y(top) - size * 0.8  # top 은 글자 상단 기준
        width = _sw(s, font, size)
        if align == "right":
            x -= width
        elif align == "center":
            x -= width / 2
        cx = x
        for t, f in _runs(s, font):
            c.setFont(f, size)
            c.drawString(cx, y, t)
            cx += pdfmetrics.stringWidth(t, f, size)
        return x + width

    def rect(self, x, top, w, h, fill=None, stroke=None, r=0, lw=0.8):
        c = self.c
        if fill is not None:
            c.setFillColor(fill)
        if stroke is not None:
            c.setStrokeColor(stroke)
            c.setLineWidth(lw)
        args = (x, self.y(top) - h, w, h)
        if r:
            c.roundRect(*args, r, stroke=1 if stroke is not None else 0, fill=1 if fill is not None else 0)
        else:
            c.rect(*args, stroke=1 if stroke is not None else 0, fill=1 if fill is not None else 0)

    def hline(self, x, top, w, color=LINE, lw=0.8):
        self.c.setStrokeColor(color)
        self.c.setLineWidth(lw)
        self.c.line(x, self.y(top), x + w, self.y(top))

    def vline(self, x, top, h, color=LINE, lw=0.8):
        self.c.setStrokeColor(color)
        self.c.setLineWidth(lw)
        self.c.line(x, self.y(top), x, self.y(top + h))

    def link(self, url, x, top, w, h):
        if not url:
            return
        self.c.linkURL(url, (x, self.y(top + h), x + w, self.y(top)), relative=0)


def _article_kind(title):
    """묶음: 여러 기업 소식을 함께 다룬 브리핑·뉴스레터형 (제목이 [코너명] 으로 시작)."""
    return "묶음" if title.startswith("[") else "단독"


def _prepare(rows):
    out = []
    for r in rows:
        title = r["제목"]
        review = title.startswith(REVIEW_PREFIX)
        if review:
            title = title[len(REVIEW_PREFIX):]
        portals = r.get("포털") or ["네이버"]
        naver = "네이버" in portals
        cp = naver and bool(r.get("네이버CP"))
        out.append({
            "media": r["매체명"],
            "title": title,
            "url": r["URL"],
            "dt": r["게재일자"],
            "review": review,
            "cp": cp,
            "naver": naver,
            # 대표 포털 (막대·점 색과 집계 기준): 네이버 CP > 네이버 검색제휴 > 다음 > 구글
            "portal": ("네이버 CP" if cp else "네이버 검색제휴") if naver else portals[0],
            "portals_text": "/".join(portals),
            "category": media_category(r["매체명"]),
            "kind": _article_kind(title),
        })
    # 네이버 CP사 -> 검색제휴사 -> 다음 -> 구글, 각각 매체 유형 순(통신사 > 일간지 > ... > 온라인매체)
    out.sort(key=lambda a: (PORTAL_RANK[a["portal"]], media_sort_key(a["media"]), a["dt"]))
    return out


def _shown_portals(arts):
    """네이버 두 항목은 항상, 다음/구글은 해당 기사가 있을 때만 표시."""
    present = {a["portal"] for a in arts}
    return [(n, c) for n, c in PORTALS if n.startswith("네이버") or n in present]


# ---------------------------------------------------------------- 1쪽

def _korean_date(d, with_year=True):
    return f"{d.year}년 {d.month}월 {d.day}일" if with_year else f"{d.month}월 {d.day}일"


def _period_text(start, end):
    """예: 10월 15일 ~ 10월 16일 (하루면 10월 15일)"""
    if start == end:
        return _korean_date(start, False)
    return f"{_korean_date(start, False)} ~ {_korean_date(end, False)}"


def _draw_header(p, company, title, dist_dt, last_dt):
    """[기업명 굵게 + 보고서명] 한 줄 > 정보 카드(보도자료 제목 / 배포일·집계일).
    정보 카드는 아래 핵심 수치 카드와 같은 모양·여백이고, 집계일 칸은 수치 카드의 3번째 칸 위치에 맞춘다."""
    company = re.sub(r"\s*[(（].*?[)）]", "", company).strip() or company  # 괄호 부분 제외
    size = 21
    report_name = "보도자료 배포 결과보고서"
    x = p.text(X0, TOP, company, font=BOLD, size=size)
    top = TOP
    if x + 8 + _sw(report_name, MED, size) <= X0 + W:
        p.text(x + 8, top, report_name, font=MED, size=size, color=SUBTITLE)
    else:  # 기업명이 길면 보고서명은 다음 줄로
        top += size + 9
        p.text(X0, top, report_name, font=MED, size=size, color=SUBTITLE)
    top += size + 18

    pad = 15
    col_w = (W - 2 * pad) / 5
    lines = _wrap(title, BOLD, 13, W - 2 * pad)
    meta_h = 34
    h = pad + 2 + 16 + 19 * len(lines) + 8 + 12 + meta_h + pad - 2
    p.rect(X0, top, W, h, fill=WHITE, stroke=LINE, r=10)

    t = top + pad + 2
    p.text(X0 + pad, t, "보도자료 제목", font=MED, size=9, color=MUTED)
    t += 16
    for line in lines:
        p.text(X0 + pad, t, line, font=BOLD, size=13)
        t += 19
    t += 8
    p.hline(X0 + pad, t, W - 2 * pad)
    t += 12

    meta = [
        ("배포일", f"{_korean_date(dist_dt)} {WEEKDAYS[dist_dt.weekday()]}요일"),
        ("집계일", _period_text(dist_dt, last_dt)),
    ]
    for i, (label, value) in enumerate(meta):
        cx = X0 + pad + i * col_w * 2 + (0 if i == 0 else 15)
        p.text(cx, t, label, font=MED, size=9, color=MUTED)
        p.text(cx, t + 15, value, font=SEMI, size=11)
        if i:
            p.vline(X0 + pad + i * col_w * 2, t - 2, meta_h - 4)
    return top + h + 12


def _draw_summary_card(p, top, arts, kpis):
    pad = 15
    inner_w = W - 2 * pad
    n = len(arts)
    gap = 4
    per_row, sq_h = (n, 24) if n <= 30 else (30, 24) if n <= 60 else (50, 12)
    sq_w = (inner_w - gap * (per_row - 1)) / per_row if per_row else 0
    sq_rows = -(-n // per_row) if n else 0

    kpi_h = 54
    h = pad + kpi_h + 22 + 14 + sq_rows * (sq_h + gap) + 8 + 12 + pad
    p.rect(X0, top, W, h, fill=WHITE, stroke=LINE, r=10)

    col_w = inner_w / len(kpis)
    for i, (num, label) in enumerate(kpis):
        cx = X0 + pad + i * col_w + (0 if i == 0 else 15)
        t = top + pad + 6
        x = p.text(cx, t, str(num), font=BOLD, size=22, color=BLUE if i == 0 else INK)
        p.text(x + 2, t + 9, "건", font=MED, size=11, color=MUTED)
        lt = t + 30
        for line in _wrap(label, REG, 9.8, col_w - 30):
            p.text(cx, lt, line, color=MUTED)
            lt += 14
        if i:
            p.vline(X0 + pad + i * col_w, top + pad, kpi_h + 10)

    t = top + pad + kpi_h + 22
    p.text(X0 + pad, t, "게재 기사 한눈에 보기 · 막대 하나가 기사 한 건입니다", size=9.3, color=MUTED)
    t += 14 + 5
    for i, a in enumerate(arts):
        row, col = divmod(i, per_row)
        color = PORTAL_COLOR[a["portal"]]
        p.rect(X0 + pad + col * (sq_w + gap), t + row * (sq_h + gap), sq_w, sq_h, fill=color, r=3)
    t += sq_rows * (sq_h + gap) + 8

    counts = Counter(a["portal"] for a in arts)
    x = X0 + pad
    for name, color in _shown_portals(arts):
        p.rect(x, t + 1.5, 8, 8, fill=color, r=1.5)
        x = p.text(x + 12, t, f"{name} {counts.get(name, 0)}건", size=9.3, color=MUTED) + 14
    return top + h


def _draw_featured(p, top, arts, limit):
    """선정 규칙:
    1) 네이버 CP사 기사가 있으면 CP사 기사 (최대 MAX_FEATURED건)
    2) 없으면 네이버 노출 기사 중 매체 영향력 순(media_order 우선순위) 상위 4건 - 한 매체당 1건
    3) 둘 다 없으면 이 칸을 그리지 않음
    limit: 이 높이(top 좌표)를 넘기지 않는 카드 줄까지만 그림 (1쪽 넘침 방지)."""
    cp = [a for a in arts if a["cp"]]
    if cp:
        subtitle, picks = "네이버 뉴스 콘텐츠제휴(CP) 매체에 실린 기사", cp[:MAX_FEATURED]
    else:
        subtitle, picks, seen = "네이버 뉴스 탭에 실린 주요 매체 기사", [], set()
        for a in sorted((a for a in arts if a["naver"]), key=lambda a: (media_sort_key(a["media"]), a["dt"])):
            if a["media"] not in seen:
                seen.add(a["media"])
                picks.append(a)
        picks = picks[:4]
    gap = 10
    cw = (W - gap) / 2
    card_rows = []
    head_h = 22 + 40
    used = top + head_h
    for i in range(0, len(picks), 2):
        pair = picks[i:i + 2]
        wrapped = [_wrap(a["title"], MED, 9.8, cw - 30) for a in pair]
        h = 40 + 15 * max(len(w) for w in wrapped)
        if used + h > limit:
            break
        card_rows.append((pair, wrapped, h))
        used += h + 8
    if not card_rows:
        return top

    top += 22
    p.text(X0, top, "주요 매체 게재", font=BOLD, size=13.5)
    p.text(X0, top + 21, subtitle, size=9.3, color=MUTED)
    top += 40

    for pair, wrapped, h in card_rows:
        for j, (a, lines) in enumerate(zip(pair, wrapped)):
            x = X0 + j * (cw + gap)
            p.rect(x, top, cw, h, fill=WHITE, stroke=LINE, r=6)
            p.rect(x, top, cw, 9, fill=BLUE, r=6)  # 상단 파란 띠 (모서리만 둥글게)
            p.rect(x + 0.8, top + 2.5, cw - 1.6, 7, fill=WHITE)
            p.text(x + 15, top + 15, a["media"], font=BOLD, size=9.3, color=BLUE)
            lt = top + 34
            for line in lines:
                p.text(x + 15, lt, line, font=MED, size=9.8)
                lt += 15
            p.link(a["url"], x, top, cw, h)
        top += h + 8
    return top


def _bar_cards_height(arts):
    n_cats = len({a["category"] for a in arts})
    return 16 + 70 + 21.5 * max(len(_shown_portals(arts)), n_cats, 1) + 10


def _draw_bar_cards(p, top, arts):
    portal_counts = Counter(a["portal"] for a in arts)
    portal_rows = [(name, portal_counts.get(name, 0), color) for name, color in _shown_portals(arts)]

    cat_counts = Counter(a["category"] for a in arts)
    order = list(CATEGORY_ORDER) + ["확인필요"]
    cats = [(c, cat_counts[c]) for c in order if cat_counts.get(c)]
    top_cat = max((n for _, n in cats), default=0)
    cat_rows = [(c, n, BLUE if n == top_cat else GRAY) for c, n in cats]

    top += 16
    gap = 10
    cw = (W - gap) / 2
    row_h = 21.5
    h = 70 + row_h * max(len(portal_rows), len(cat_rows), 1) + 10
    cards = [
        ("포털 노출 구분", "네이버 노출 기사는 CP·검색제휴로 나눠 집계", portal_rows),
        ("매체 유형", "게재 매체의 성격별 분포", cat_rows),
    ]
    for i, (title, sub, rows) in enumerate(cards):
        x = X0 + i * (cw + gap)
        p.rect(x, top, cw, h, fill=WHITE, stroke=LINE, r=10)
        p.text(x + 15, top + 22, title, font=BOLD, size=13)
        p.text(x + 15, top + 44, sub, size=9.3, color=MUTED)
        vmax = max((n for _, n, _ in rows), default=0) or 1
        bar_x, bar_w = x + 89, 126
        for k, (label, n, color) in enumerate(rows):
            t = top + 68 + k * row_h
            p.text(x + 15, t, label, size=9)
            p.rect(bar_x, t + 1.5, bar_w, 6, fill=LINE, r=3)
            if n:
                p.rect(bar_x, t + 1.5, max(6, bar_w * n / vmax), 6, fill=color, r=3)
            p.text(x + cw - 15, t, f"{n}건", font=BOLD, size=9, align="right")
    return top + h


# ---------------------------------------------------------------- 2쪽~

# (헤더, x 오프셋(카드 왼쪽 기준), 폭)
COLS = [
    ("No", 15, 18), ("매체명", 37, 74), ("제목", 115, 199),
    ("매체 유형", 324, 52), ("기사 유형", 378, 40), ("게재일", 422, 32), ("게재 포털", 455, 57),
]
ROW_FONT = 7.9
ROW_LEAD = 11


def _table_rows(arts):
    title_w = COLS[2][2]
    rows = []
    for a in arts:
        title = ("[확인필요] " if a["review"] else "") + a["title"]
        lines = _wrap(title, REG, ROW_FONT, title_w)
        media_lines = _wrap(a["media"], SEMI, ROW_FONT, COLS[1][2] - 10)
        n = max(len(lines), len(media_lines))
        rows.append((a, lines, media_lines, 12 + max(n, 1) * ROW_LEAD + (2 if n == 1 else 0)))
    return rows


TABLE_DESC = (
    "포털 노출 순(네이버 CP → 네이버 검색제휴 → 다음 → 구글)으로, 같은 포털 안에서는 매체 영향력 순"
    "(통신사 → 일간지 → 주요 경제지 → IT전문지 → 산업지 → 일반 경제지 → 온라인매체)으로 정렬했습니다. "
    "제목을 누르면 기사 원문으로 이동합니다."
)


def _draw_table_page(p, top, rows, first):
    """rows 를 한 페이지 카드에 들어가는 만큼 그리고, 남은 rows 반환."""
    x = X0
    head_top = top + 18
    if first:
        desc_lines = _wrap(TABLE_DESC, REG, 9.8, W - 50)
        hdr = head_top + 24 + 15 * len(desc_lines) + 10
    else:
        hdr = head_top

    avail_bottom = PAGE_H - BOTTOM - 50
    used = []
    t = hdr + 18
    for r in rows:
        if t + r[3] > avail_bottom and used:
            break
        used.append((t, r))
        t += r[3]
    card_h = t - top + 12

    p.rect(x, top, W, card_h, fill=WHITE, stroke=LINE, r=10)
    if first:
        p.text(x + 15, head_top, "게재 기사 전체", font=BOLD, size=13.5)
        t2 = head_top + 24
        for line in desc_lines:
            p.text(x + 15, t2, line, color=MUTED)
            t2 += 15

    for name, off, _ in COLS:
        p.text(x + off, hdr, name, font=SEMI, size=7.5, color=MUTED)
    p.hline(x + 15, hdr + 14, W - 30)

    for row_top, (a, lines, media_lines, h) in used:
        mid = row_top + h / 2  # 행 세로 가운데 - 모든 칸을 여기에 맞춤

        def ctop(n_lines):  # n줄짜리 글 묶음이 행 가운데 오도록 하는 첫 줄 top
            return mid - (n_lines - 1) * ROW_LEAD / 2 - 0.45 * ROW_FONT

        one = ctop(1)
        p.text(x + COLS[0][1], one, str(a["no"]), size=ROW_FONT, color=MUTED)
        p.c.setFillColor(PORTAL_COLOR[a["portal"]])
        p.c.circle(x + COLS[1][1] + 2.5, p.y(mid), 2.5, stroke=0, fill=1)
        mt = ctop(len(media_lines))
        for k, ml in enumerate(media_lines):
            p.text(x + COLS[1][1] + 10, mt + k * ROW_LEAD, ml, font=SEMI, size=ROW_FONT)
        tt = ctop(len(lines))
        for k, line in enumerate(lines):
            if k == 0 and a["review"]:
                tag = "[확인필요] "
                nx = p.text(x + COLS[2][1], tt, tag, font=SEMI, size=ROW_FONT, color=ORANGE)
                p.text(nx, tt, line[len(tag):], size=ROW_FONT)
            else:
                p.text(x + COLS[2][1], tt + k * ROW_LEAD, line, size=ROW_FONT)
        p.link(a["url"], x + COLS[2][1], row_top, COLS[2][2], h)
        p.text(x + COLS[3][1], one, a["category"], size=ROW_FONT, color=MUTED)

        if a["kind"] == "묶음":  # 단독과 같은 톤에서 옅은 회색 바탕으로만 구분
            p.rect(x + COLS[4][1], mid - 8, 29, 16, fill=PILL_BG, r=8)
        else:
            p.rect(x + COLS[4][1], mid - 8, 29, 16, stroke=LINE, r=8, lw=0.8)
        p.text(x + COLS[4][1] + 14.5, one, a["kind"], size=ROW_FONT, color=MUTED, align="center")

        p.text(x + COLS[5][1], one, f"{a['dt'].month}.{a['dt'].day:02d}", size=ROW_FONT, color=MUTED)
        p.text(x + COLS[6][1], one, a["portals_text"], size=ROW_FONT, color=MUTED)
        p.hline(x + 15, row_top + h, W - 30)

    rest = rows[len(used):]
    return top + card_h, rest


def _draw_footer(p, top, has_review, portals):
    top += 20
    p.hline(X0, top, W)
    left = f"본 보고서는 {'·'.join(portals)} 포털 모니터링 결과를 바탕으로 작성되었습니다."
    right = "Prepared by OPENPR"
    p.text(X0, top + 20, left, size=8.5, color=MUTED)
    p.text(X0 + W, top + 20, right, size=8.5, color=MUTED, align="right")
    if has_review:
        p.text(X0, top + 36, "[확인필요] 표시 기사는 해당 보도자료 기사인지 담당자 확인이 필요합니다", size=8.5, color=ORANGE)


# ---------------------------------------------------------------- 진입점

def build_pdf(title, company, distribution_date, rows, window_days=3, generated_on=None):
    """rows: report_sheet.resolve_coverage()가 반환한 리스트. 반환값: PDF bytes.
    (window_days / generated_on 은 이전 호출부 호환용)"""
    _register_fonts()
    dist_dt = datetime.strptime(distribution_date, "%Y-%m-%d").date()
    arts = _prepare(rows)
    for i, a in enumerate(arts, start=1):
        a["no"] = i

    last = max((a["dt"].date() for a in arts), default=dist_dt)

    total = len(arts)
    kpis = [
        (total, "총 게재 기사"),
        (sum(a["naver"] for a in arts), "네이버 뉴스 노출"),
        (sum(a["cp"] for a in arts), "네이버 CP사 기사"),
        (sum(a["kind"] == "단독" for a in arts), "단독 기사"),
        (sum(a["dt"].date() == dist_dt for a in arts), "배포 당일 게재"),
    ]

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"{company} {title[:40]} 보도 결과 보고서")
    c.setAuthor("오픈피알")
    p = _Page(c)

    top = _draw_header(p, company, title, dist_dt, max(last, dist_dt))
    top = _draw_summary_card(p, top, arts, kpis)
    top = _draw_featured(p, top, arts, limit=PAGE_H - 28 - _bar_cards_height(arts))  # 1쪽은 꼬리말이 없어 여백을 좁게
    top = _draw_bar_cards(p, top, arts)

    remaining = _table_rows(arts)
    first = True
    while True:
        c.showPage()
        if not remaining:
            if first:
                p.rect(X0, TOP, W, 80, fill=WHITE, stroke=LINE, r=10)
                p.text(X0 + 15, TOP + 18, "게재 기사 전체", font=BOLD, size=13.5)
                p.text(X0 + 15, TOP + 46, "집계 기간 내 수집된 기사가 없습니다.", color=MUTED)
                end = TOP + 80
            break
        end, remaining = _draw_table_page(p, TOP, remaining, first)
        first = False
        if not remaining:
            break
    used_portals = [n for n in ("네이버", "다음", "구글") if n == "네이버" or any(n in a["portals_text"] for a in arts)]
    _draw_footer(p, end, any(a["review"] for a in arts), used_portals)
    c.save()
    return buf.getvalue()
