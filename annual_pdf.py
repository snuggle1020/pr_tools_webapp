# -*- coding: utf-8 -*-
"""
연간 PR리포트 PDF (A4 세로, 월간 PR리포트와 같은 글꼴·색·표 모양).
쪽 구성: 1쪽 핵심 수치·요약·월별 추이·전년 대비 / 2쪽 연도별 추이·하이라이트·주제·요일
        3쪽 보도자료 순위·매체 유형·포털 / 4쪽 매체·기자 TOP 5·기자 미팅·총 정리
한 쪽에 안 들어가는 구역은 다음 쪽으로 넘기고, 쪽마다 남는 높이는 구역 사이에 고르게 나눈다.
입력은 annual_report.compose_annual() 결과. 데이터가 없는 구역(포털·기자·미팅 등)은 빠지고 번호가 당겨진다.
"""
import io
import math
from datetime import date

from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas

from monthly_pdf import BG, FAINT, ORANGE, PURPLE
from report_pdf import (
    BLUE, BOLD, GRAY, GREEN, INK, LIGHT_BLUE, LINE, MED, MUTED, PAGE_H, REG, SEMI, SUBTITLE, TOP, W, WHITE, X0,
    _clip_lines, _Page, _register_fonts, _sw, _wrap,
)

SOFT_BLUE = HexColor("#F2F5FD")
TRACK = HexColor("#EEF0F4")
SERIES = [("release", "보도자료", BLUE), ("plan", "기획", GREEN), ("interview", "인터뷰", PURPLE),
          ("contrib", "기고", ORANGE), ("mention", "단순언급", GRAY)]
WD = "월화수목금토일"
CONTENT_TOP = TOP + 48
CONTENT_BOTTOM = PAGE_H - 50
BASE_GAP, MAX_GAP = 22, 46
SEC_H = 26  # 구역 제목 줄


def _d(d):
    return f"{d.month}.{d.day} ({WD[d.weekday()]})" if d else "-"


def _sg(v):
    return "-" if v is None else (f"+{v}" if v > 0 else str(v))


def _pct(a, b):
    return round((a / b - 1) * 100) if b else None


class Block:
    """한 구역: 높이 h, draw(p, top) 로 그림."""

    def __init__(self, h, draw):
        self.h, self.draw = h, draw


def _section(p, x, top, num, title, w=W, note=None):
    nx = p.text(x, top + 2, num, font=BOLD, size=11, color=BLUE)
    p.text(nx + 7, top, title, font=BOLD, size=14)
    if note:
        note(x + w, top + 4) if callable(note) else p.text(x + w, top + 4, note, size=9, color=MUTED, align="right")


class _Tbl:
    """표 그리기 (월간 '주요 기사' 표와 같은 모양: 검은 머리글, 줄마다 아래 선)."""

    def __init__(self, cols, w, size=9, lead=13, pad=6, head_h=22):
        fixed = sum(c[1] for c in cols if c[1])
        self.cols = cols
        self.widths = [c[1] or (w - fixed) for c in cols]
        self.w, self.size, self.lead, self.pad, self.head_h = w, size, lead, pad, head_h

    def _lines(self, row, font):
        return [_wrap(str(v), font, self.size, cw - 12) for v, cw in zip(row, self.widths)]

    def height(self, rows, bold=()):
        return self.head_h + sum(self.pad * 2 + max(len(l) for l in self._lines(r, BOLD if i in bold else REG)) * self.lead
                                 for i, r in enumerate(rows))

    def draw(self, p, x, top, rows, bold=(), fill=(), links=None):
        xs = [x + sum(self.widths[:i]) for i in range(len(self.cols))]
        p.rect(x, top, self.w, self.head_h, fill=INK)
        for (name, _, al), cx, cw in zip(self.cols, xs, self.widths):
            tx = cx + 7 if al == "left" else cx + cw / 2 if al == "center" else cx + cw - 7
            p.text(tx, top + (self.head_h - 8.8) / 2, name, font=SEMI, size=8.8, color=WHITE, align=al)
        t = top + self.head_h
        for i, row in enumerate(rows):
            font = BOLD if i in bold else REG
            lines = self._lines(row, font)
            n = max(len(l) for l in lines)
            h = self.pad * 2 + n * self.lead
            if i in fill:
                p.rect(x, t, self.w, h, fill=BG)
            for (_, _, al), cx, cw, ls in zip(self.cols, xs, self.widths, lines):
                tx = cx + 7 if al == "left" else cx + cw / 2 if al == "center" else cx + cw - 7
                lt = t + self.pad + (n - len(ls)) * self.lead / 2 + (self.lead - self.size) / 2
                for line in ls:
                    p.text(tx, lt, line, font=font, size=self.size, align=al)
                    lt += self.lead
            if links and links[i]:
                k, url = links[i]
                p.link(url, xs[k], t, self.widths[k], h)
            t += h
            p.hline(x, t, self.w)
        return t - top


def _bars(p, x, top, w, rows, label_w, row_h=22, unit="건", x_w=26):
    """막대 줄: rows [(이름, 값, 보조 글, 파랑 여부|None)]. 파랑 여부가 None 이면 가장 큰 값만 파랑. 반환 높이."""
    mx = max((r[1] for r in rows), default=0) or 1
    val_w = 46
    for i, (name, v, extra, is_blue) in enumerate(rows):
        t = top + i * row_h
        mid = t + row_h / 2
        col = FAINT if not v else INK
        p.text(x, mid - 4.8, name, font=MED, size=9.5, color=col)
        bx = x + label_w
        bw = w - label_w - val_w - (x_w if extra is not None else 0) - 6
        p.rect(bx, mid - 3.5, bw, 7, fill=BG, r=3.5)
        if v:
            b = is_blue if is_blue is not None else v == mx
            p.rect(bx, mid - 3.5, max(7, bw * v / mx), 7, fill=BLUE if b else GRAY, r=3.5)
        vx = x + w - (x_w if extra is not None else 0)
        vs = f"{v:g}{unit}" if isinstance(v, (int, float)) else str(v)
        p.text(vx, mid - 4.8, vs, font=BOLD, size=9.5, color=col, align="right")
        if extra is not None:
            p.text(x + w, mid - 4.3, extra, size=8.5, color=MUTED, align="right")
    return len(rows) * row_h


# ---------------------------------------------------------------- 구역들

def _intro(d):
    kpis = [(f"{d['total']}", "건", "연간 총 보도 건수", BLUE)]
    if d["prev_total"]:
        kpis.append((_sg(_pct(d["total"], d["prev_total"])), "%", "전년 대비" if d["last_month"] == 12 else "전년 동기 대비", INK))
    kpis.append((f"{len(d['releases'])}", "건", "보도자료 배포 건수", INK))
    kpis.append((f"{d['rel_avg']:.1f}", "건", "보도자료당 평균 게재", INK))
    text = (d.get("summary") or "").strip()
    lines = []
    for para in text.splitlines():
        lines += _wrap(para.strip(), REG, 10.3, W - 32) if para.strip() else []
    kh = 62
    sh = (14 + 14 + 16.5 * len(lines) + 8) if lines else 0
    h = kh + (12 + sh if sh else 0)

    def draw(p, top):
        p.rect(X0, top, W, kh, fill=WHITE, stroke=LINE, r=9)
        cw = W / len(kpis)
        for i, (num, unit, label, color) in enumerate(kpis):
            mid = X0 + i * cw + cw / 2
            nx = p.text(mid - (_sw(num, BOLD, 24) + 2 + _sw(unit, SEMI, 10.5)) / 2, top + 12, num, font=BOLD, size=24, color=color)
            p.text(nx + 2, top + 22.5, unit, font=SEMI, size=10.5, color=MUTED)
            p.text(mid, top + 43, label, font=MED, size=9.5, color=MUTED, align="center")
            if i:
                p.vline(X0 + i * cw, top, kh)
        if sh:
            t = top + kh + 12
            p.rect(X0, t, W, sh, fill=BG, r=8)
            p.text(X0 + 16, t + 13, "올해의 요약", font=SEMI, size=9.3, color=BLUE)
            lt = t + 31
            for line in lines:
                p.text(X0 + 16, lt, line, size=10.3)
                lt += 16.5
    return Block(h, draw)


def _monthly(d):
    prev = d["prev_monthly"]
    last = d["last_month"]
    tot = d["monthly_total"]
    peak = max(range(last), key=lambda i: tot[i]) if any(tot[:last]) else None
    pad, chart_h, head_h, row_h = 12, 92, 17, 15.5
    box_h = pad + chart_h + 6 + head_h + 6 * row_h + pad
    h = SEC_H + box_h

    def legend(right, top):
        if prev:
            tx = p_ref[0].text(right, top, f"{d['year'] - 1}년 월별 합계", size=8.5, color=MUTED, align="right")
            x2 = right - _sw(f"{d['year'] - 1}년 월별 합계", REG, 8.5) - 5
            c = p_ref[0].c
            c.setDash(3, 2)
            c.setStrokeColor(INK)
            c.setLineWidth(0.9)
            c.line(x2 - 14, p_ref[0].y(top + 4.5), x2, p_ref[0].y(top + 4.5))
            c.setDash()

    p_ref = [None]

    def draw(p, top):
        p_ref[0] = p
        _section(p, X0, top, d["_num"]["01"], "월별 보도 추이", note=legend if prev else None)
        bt = top + SEC_H
        p.rect(X0, bt, W, box_h, fill=WHITE, stroke=LINE, r=8)
        ix, iw = X0 + 12, W - 24
        lbl, totw = 58, 34
        slot = (iw - lbl - totw) / 12
        cx = lambda i: ix + lbl + slot * i + slot / 2
        base = bt + pad + chart_h
        mx = max(max(tot), max(prev) if prev else 0) or 1
        sc = (chart_h - 14) / mx
        # 가장 많은 달 세로 띠
        if peak is not None:
            p.rect(cx(peak) - slot / 2, base + 6, slot, head_h + 6 * row_h, fill=SOFT_BLUE)
        bw = min(17, slot * 0.6)
        for i in range(12):
            if i >= last:
                continue
            y = base
            for key, _, col in SERIES:
                v = d["monthly"][key][i]
                if v:
                    p.rect(cx(i) - bw / 2, y - v * sc, bw, v * sc, fill=col)
                    y -= v * sc
            pk = i == peak
            p.text(cx(i), y - 11, str(tot[i]), font=BOLD if pk else MED, size=8.3, color=INK if pk else MUTED, align="center")
        if prev:
            c = p.c
            c.setDash(3, 2)
            c.setStrokeColor(INK)
            c.setLineWidth(0.9)
            path = c.beginPath()
            for i in range(12):
                X, Y = cx(i), p.y(base - prev[i] * sc)
                path.moveTo(X, Y) if i == 0 else path.lineTo(X, Y)
            c.drawPath(path, stroke=1, fill=0)
            c.setDash()
        # 형식별 건수 표 (막대와 같은 세로줄)
        t = base + 6
        p.text(ix, t + 4.5, "형식", font=SEMI, size=8.5, color=MUTED)
        for i in range(12):
            p.text(cx(i), t + 4.5, f"{i + 1}월", font=BOLD if i == peak else SEMI, size=8.5, color=INK if i == peak else MUTED, align="center")
        p.text(ix + iw, t + 4.5, "합계", font=SEMI, size=8.5, color=MUTED, align="right")
        p.hline(ix, t + head_h, iw, color=INK, lw=0.8)
        t += head_h
        for key, name, col in SERIES:
            p.rect(ix, t + 4.5, 6.5, 6.5, fill=col, r=1.5)
            p.text(ix + 10, t + 4, name, size=8.5)
            for i in range(12):
                v = d["monthly"][key][i]
                p.text(cx(i), t + 4, str(v) if v and i < last else "–", size=8.5, color=INK if v and i < last else FAINT, align="center")
            p.text(ix + iw, t + 4, str(d["types"][key]), font=BOLD, size=8.5, align="right")
            t += row_h
        p.hline(ix, t, iw, color=INK, lw=0.8)
        p.text(ix, t + 4, "합계", font=BOLD, size=8.5)
        for i in range(12):
            p.text(cx(i), t + 4, str(tot[i]) if i < last else "–", font=BOLD, size=8.5, color=INK if i < last else FAINT, align="center")
        p.text(ix + iw, t + 4, str(d["total"]), font=BOLD, size=8.5, color=BLUE, align="right")
    return Block(h, draw)


def _yoy(d):
    py, cy = d["year"] - 1, d["year"]
    pt = d["prev_types"]
    rows, rates = [], []
    for key, name, _ in SERIES:
        cur = d["types"][key]
        pv = pt[key] if pt else None
        r = _pct(cur, pv) if pv else None
        rates.append(r)
        rows.append([name, f"{pv}건" if pv is not None else "-", f"{cur}건", (_sg(cur - pv) + "건") if pv is not None else "-", ""])
    tr = _pct(d["total"], d["prev_total"]) if d["prev_total"] else None
    rows.append(["합계", f"{d['prev_total']}건" if d["prev_total"] else "-", f"{d['total']}건",
                 (_sg(d["total"] - d["prev_total"]) + "건") if d["prev_total"] else "-", ""])
    tbl = _Tbl([("형식", 104, "center"), (f"{py}년", 84, "center"), (f"{cy}년", 84, "center"), ("증감", 78, "center"),
                ("증감률", None, "center")], W, pad=6.5)
    th = tbl.height(rows)
    note = None if d["last_month"] == 12 else f"1~{d['last_month']}월 기준 (전년 같은 기간과 비교)"
    h = SEC_H + th

    def draw(p, top):
        _section(p, X0, top, d["_num"]["02"], "전년 대비 증감률", note=note)
        t0 = top + SEC_H
        tbl.draw(p, X0, t0, rows, bold={len(rows) - 1}, fill={len(rows) - 1})
        # 증감률 칸: 막대 + 숫자
        cx0 = X0 + sum(tbl.widths[:4])
        cw = tbl.widths[4]
        mx = max((abs(r) for r in rates if r is not None), default=0) or 1
        t = t0 + tbl.head_h
        rh = tbl.pad * 2 + tbl.lead
        for i, r in enumerate(rates + [tr]):
            mid = t + rh / 2
            tw, nw = 90, 36
            bx = cx0 + (cw - tw - 6 - nw) / 2
            if i < len(rates):
                p.rect(bx, mid - 3, tw, 6, fill=TRACK, r=3)
                if r is not None and r != 0:
                    p.rect(bx, mid - 3, max(6, tw * abs(r) / mx), 6, fill=BLUE if r > 0 else GRAY, r=3)
            p.text(bx + tw + 6 + nw, mid - 4.5, (_sg(r) + "%") if r is not None else "-",
                   font=BOLD if i == len(rates) else REG, size=9, align="right")
            t += rh
    return Block(h, draw)


def _yearly(d):
    ys = d["yearly"]
    mx = max(y[1] for y in ys) or 1
    bar_max, box_h = 100, 196

    def draw(p, top):
        _section(p, X0, top, d["_num"]["03"], "연도별 보도 추이")
        bt = top + SEC_H
        p.rect(X0, bt, W, box_h, fill=WHITE, stroke=LINE, r=8)
        colw = W / max(len(ys), 1)
        base = bt + 16 + 16 + bar_max
        for i, (y, n, rels, (m1, m2), chg) in enumerate(ys):
            mid = X0 + colw * i + colw / 2
            cur = y == d["year"]
            bh = n / mx * bar_max
            p.rect(mid - 17, base - bh, 34, bh, fill=BLUE if cur else GRAY, r=2)
            label = f"{n}"
            vx = p.text(mid - (_sw(label, BOLD, 11.5) + 1.5 + _sw("건", MED, 8.5)) / 2, base - bh - 16, label,
                        font=BOLD, size=11.5, color=BLUE if cur else MUTED)
            p.text(vx + 1.5, base - bh - 13.5, "건", font=MED, size=8.5, color=BLUE if cur else MUTED)
            p.hline(mid - min(42, colw / 2 - 6), base, min(84, colw - 12), color=INK, lw=0.8)
            yl = f"{y}년" + (f" ({m1}~{m2}월)" if (m1, m2) != (1, 12) else "")
            p.text(mid, base + 7, yl, font=SEMI, size=10, color=BLUE if cur else INK, align="center")
            p.text(mid, base + 23, f"보도자료 {rels}건", size=8.5, color=MUTED, align="center")
            if chg is not None:
                lab = "전년 동기 대비" if cur and d["last_month"] < 12 else "전년 대비"
                p.text(mid, base + 35, f"{lab} {_sg(chg)}%", size=8.5, color=MUTED, align="center")
    return Block(SEC_H + box_h, draw)


def _highlights(d):
    hs = d["highlights"]
    gap = 8
    cw = (W - 2 * gap) / 3
    card_h = 80
    rows = math.ceil(len(hs) / 3)
    h = SEC_H + rows * card_h + (rows - 1) * gap

    def draw(p, top):
        _section(p, X0, top, d["_num"]["04"], "올해의 하이라이트", note="제목을 누르면 기사로 이동")
        for i, a in enumerate(hs):
            x = X0 + (i % 3) * (cw + gap)
            t = top + SEC_H + (i // 3) * (card_h + gap)
            p.rect(x, t, cw, card_h, fill=WHITE, stroke=LINE, r=8)
            p.text(x + 12, t + 11, a["label"], font=SEMI, size=8.8, color=BLUE)
            lines = _clip_lines(_wrap(a["title"], BOLD, 10.5, cw - 24), BOLD, 10.5, cw - 24, 2)
            lt = t + 27
            for line in lines:
                p.text(x + 12, lt, line, font=BOLD, size=10.5)
                lt += 15
            p.text(x + 12, t + card_h - 19, f"{a['media']} · {_d(a['date'])} 게재", size=8.5, color=MUTED)
            p.link(a.get("url"), x, t, cw, card_h)
    return Block(h, draw)


def _pair(left, right):
    """나란히 놓는 두 구역 [(번호키, 제목, 메모, 상자 안 높이, 상자 그리기(p,x,top,w,h))]. 상자 높이는 큰 쪽에 맞춤."""
    gap = 14
    items = [x for x in (left, right) if x]
    cw = (W - gap) / 2 if len(items) == 2 else W
    inner = max(it[3] for it in items)
    h = SEC_H + inner

    def draw(p, top, d):
        for j, (key, title, note, ih, fn) in enumerate(items):
            x = X0 + j * (cw + gap)
            _section(p, x, top, d["_num"][key], title, w=cw, note=note)
            p.rect(x, top + SEC_H, cw, inner, fill=WHITE, stroke=LINE, r=8) if fn[1] else None
            fn[0](p, x, top + SEC_H, cw, inner)
    return h, draw


def _box_bars(rows, label_w, row_h=22):
    """상자 안 막대 줄들을 위아래 가운데에."""
    ih = len(rows) * row_h + 24

    def fn(p, x, top, w, h):
        _bars(p, x + 14, top + (h - len(rows) * row_h) / 2, w - 28, rows, label_w, row_h)
    return ih, (fn, True)


def _topic_weekday(d):
    trows = [(t, round(avg, 1), f"({n})", None) for t, avg, n in d["topics"]]
    days = [w for w in d["weekdays"] if w[2] >= 2]
    best = max(days, key=lambda w: w[1])[0] if days else None
    wrows = [(name, round(avg, 1), f"({n})", (name == best) if best else None) for name, avg, n in d["weekdays"]]
    lh, lf = _box_bars(trows, 64, 24)
    rh, rf = _box_bars(wrows, 52, 24)
    note = "( ) 안은 배포 건수"
    h, draw = _pair(("05", "주제별 평균 게재 수", note, lh, lf), ("06", "요일별 평균 게재 수", note, rh, rf))
    return Block(h, lambda p, top: draw(p, top, d))


def _ranking(d):
    rels = sorted(d["releases"], key=lambda r: -r["n"])
    rows = [[str(i + 1), r["title"], r["topic"] or "-", _d(r["date"]), f"{r['n']}건", f"{r['key']}곳"] for i, r in enumerate(rels)]
    bold = {i for i, r in enumerate(rels) if r["n"] >= d["rel_avg"]}
    tbl = _Tbl([("순위", 36, "center"), ("보도자료", None, "left"), ("주제", 70, "center"), ("배포일", 62, "center"),
                ("게재 수", 52, "center"), ("주요 매체", 54, "center")], W, pad=6)
    th = tbl.height(rows, bold)
    h = SEC_H + th + 22

    def draw(p, top):
        _section(p, X0, top, d["_num"]["07"], "보도자료별 성과 순위", note="게재 많은 순으로 정렬")
        tbl.draw(p, X0, top + SEC_H, rows, bold=bold)
        p.text(X0, top + SEC_H + th + 9, f"굵은 글씨는 연 평균({d['rel_avg']:.1f}건) 이상 실린 보도자료입니다. "
                                          "주요 매체는 통신사·일간지·주요 경제지 중 실린 곳 수입니다.", size=8.8, color=MUTED)
    return Block(h, draw)


def _media_portal(d):
    tot = d["media_total"] or 1
    mrows = [(c, n, f"{round(n / tot * 100)}%", None) for c, n in d["media_types"]]
    lh, lf = _box_bars(mrows, 56, 22)
    right = None
    if d["portals"]:
        ptot = sum(n for _, n in d["portals"]) or 1
        cols = [BLUE, LIGHT_BLUE, GRAY]
        ph = 24 + 10 + 10 + 3 * 20 + 22

        def pfn(p, x, top, w, h):
            ix, iw = x + 14, w - 28
            t = top + (h - (ph - 24)) / 2
            bx = ix
            for (_, n), c in zip(d["portals"], cols):
                if n:
                    p.rect(bx, t, iw * n / ptot, 10, fill=c)
                    bx += iw * n / ptot
            t += 20
            for (name, n), c in zip(d["portals"], cols):
                p.rect(ix, t + 2, 7, 7, fill=c, r=1.5)
                p.text(ix + 11, t, name, size=9.5)
                p.text(ix + iw - 26, t, f"{n}건", font=BOLD, size=9.5, align="right")
                p.text(ix + iw, t + 0.5, f"{round(n / ptot * 100)}%", size=8.5, color=MUTED, align="right")
                t += 20
            p.text(ix, t + 4, "네이버 게재 기사의 경우 다음·구글 집계에서 제외", size=8.5, color=MUTED)
        right = ("09", "포털별 게재 수", None, ph, (pfn, True))
    h, draw = _pair(("08", "매체 유형별 게재 수", None, lh, lf), right)
    return Block(h, lambda p, top: draw(p, top, d))


def _tops(d):
    gap = 14
    two = bool(d["top_reporters"])
    cw = (W - gap) / 2 if two else W
    mt = _Tbl([("No", 30, "center"), ("매체", 92, "center"), ("유형", None, "center"), ("게재", 48, "center")], cw, pad=5.5)
    mrows = [[str(i + 1), m, c, f"{n}건"] for i, (m, c, n) in enumerate(d["top_media"])]
    rrows = [[str(i + 1), f"{r} 기자", m, f"{n}건"] for i, (r, m, n) in enumerate(d["top_reporters"] or [])]
    rt = _Tbl([("No", 30, "center"), ("기자", 92, "center"), ("매체", None, "center"), ("게재", 48, "center")], cw, pad=5.5)
    th = max(mt.height(mrows), rt.height(rrows) if two else 0)
    h = SEC_H + th + 22

    def draw(p, top):
        _section(p, X0, top, d["_num"]["10"], "매체 TOP 5", w=cw)
        mh = mt.draw(p, X0, top + SEC_H, mrows)
        p.text(X0, top + SEC_H + mh + 9, "올해 기사를 가장 많이 게재한 매체", size=8.8, color=MUTED)
        if two:
            x = X0 + cw + gap
            _section(p, x, top, d["_num"]["11"], "기자 TOP 5", w=cw)
            rh = rt.draw(p, x, top + SEC_H, rrows)
            p.text(x, top + SEC_H + rh + 9, "올해 기사를 가장 많이 작성한 기자", size=8.8, color=MUTED)
    return Block(h, draw)


def _meetings(d):
    m = d["meet"]
    box_top, box_sub, arrow = 64, 40, 18
    h = SEC_H + box_top + box_sub + 22
    cells = [("기자 미팅", f"{m['meetings']}", "회", [f"미팅 매체 {m['outlets']}곳"]),
             ("기사 발생 미팅", f"{m['converted']}", "회",
              [f"총 게재 기사 {m['stories'] + m['releases']}건", f"인터뷰·기획 {m['stories']}건 / 보도자료 {m['releases']}건"]),
             ("기사 전환율", f"{round(m['converted'] / m['meetings'] * 100) if m['meetings'] else 0}", "%",
              ["인터뷰·기획 기사로 이어진 미팅",
               f"{m['conv_planned']}회 / {m['meetings']}회 ({round(m['conv_planned'] / m['meetings'] * 100) if m['meetings'] else 0}%)"])]

    def draw(p, top):
        _section(p, X0, top, d["_num"]["12"], "기자 미팅 성과")
        bw = (W - 2 * arrow) / 3
        t = top + SEC_H
        for i, (label, num, unit, subs) in enumerate(cells):
            x = X0 + i * (bw + arrow)
            on = i == 2
            p.rect(x, t, bw, box_top + box_sub, fill=SOFT_BLUE if on else WHITE, stroke=BLUE if on else LINE, r=8, lw=1.1 if on else 0.8)
            col = BLUE if on else INK
            p.text(x + bw / 2, t + 14, label, font=SEMI, size=10, color=col, align="center")
            nw = _sw(num, BOLD, 23) + 2 + _sw(unit, SEMI, 10)
            nx = p.text(x + (bw - nw) / 2, t + 32, num, font=BOLD, size=23, color=col)
            p.text(nx + 2, t + 42, unit, font=SEMI, size=10, color=MUTED)
            p.hline(x, t + box_top, bw, color=HexColor("#D3DDF8") if on else LINE)
            st = t + box_top + (box_sub - len(subs) * 13) / 2 + 2
            for s in subs:
                p.text(x + bw / 2, st, s, size=9, color=MUTED, align="center")
                st += 13
            if i < 2:
                p.text(x + bw + arrow / 2, t + (box_top + box_sub) / 2 - 7, "→", size=13, color=FAINT, align="center")
        p.text(X0, t + box_top + box_sub + 9, "기사 발생 미팅: 미팅 후 해당 매체에 기사 1건 이상 게재", size=8.8, color=MUTED)
    return Block(h, draw)


def _wrapup(d):
    w = d["wrap"] or {}
    groups = [(t, [x.strip() for x in w.get(k, []) if x.strip()]) for k, t in
              (("good", "잘 된 점"), ("bad", "아쉬운 점"), ("plan", f"{d['year'] + 1} 제언"))]
    groups = [(t, [_wrap(x, REG, 10.3, W - 46) for x in items]) for t, items in groups if items]
    hs = [12 + 14 + sum(len(ls) for ls in items) * 16 + 8 for _, items in groups]
    h = SEC_H + sum(hs) + 6 * (len(hs) - 1)

    def draw(p, top):
        _section(p, X0, top, d["_num"]["13"], "총 정리")
        t = top + SEC_H
        for (title, items), bh in zip(groups, hs):
            p.rect(X0, t, W, bh, fill=BG, r=8)
            p.text(X0 + 16, t + 11, title, font=SEMI, size=9.3, color=BLUE)
            lt = t + 28
            for ls in items:
                p.c.setFillColor(MUTED)
                p.c.circle(X0 + 19, p.y(lt + 5), 1.6, stroke=0, fill=1)
                for line in ls:
                    p.text(X0 + 27, lt, line, size=10.3)
                    lt += 16
            t += bh + 6
    return Block(h, draw), bool(groups)


# ---------------------------------------------------------------- 조립

def build_annual_pdf(d):
    """d: annual_report.compose_annual() 결과. 반환: PDF bytes."""
    _register_fonts()
    d = dict(d)
    # 하이라이트 카드 위 글: 형식
    names = {"release": "보도자료", "plan": "기획", "interview": "인터뷰", "contrib": "기고"}
    d["highlights"] = [dict(a, label=names.get(a["kind"], a["kind"])) for a in d["highlights"]]

    # 들어갈 구역 (키 순서 = 번호 순서). 데이터가 없으면 빠지고 번호가 당겨짐
    plan = [["01", "02"], ["03", "04", "05"], ["07", "08"], ["10", "12", "13"]]
    has = {"01": True, "02": True, "03": True, "04": bool(d["highlights"]), "05": bool(d["releases"]),
           "06": bool(d["releases"]), "07": bool(d["releases"]), "08": bool(d["media_types"]),
           "09": bool(d["portals"]), "10": bool(d["top_media"]), "11": bool(d["top_reporters"]) and bool(d["top_media"]),
           "12": bool(d["meet"] and d["meet"]["meetings"]), "13": False}
    wrap_block, has["13"] = _wrapup(d)
    order = [k for k in ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13"] if has[k]]
    d["_num"] = {k: f"{i + 1:02d}" for i, k in enumerate(order)}
    makers = {"01": _monthly, "02": _yoy, "03": _yearly, "04": _highlights, "05": _topic_weekday, "07": _ranking,
              "08": _media_portal, "10": _tops, "12": _meetings, "13": lambda d: wrap_block}
    pages = []
    for i, keys in enumerate(plan):
        blocks = ([_intro(d)] if i == 0 else []) + [makers[k](d) for k in keys if has[k]]
        pages.append(blocks)

    # 넘치면 뒤 구역을 다음 쪽으로
    avail = CONTENT_BOTTOM - CONTENT_TOP
    out, carry = [], []
    for blocks in pages + [[]]:
        blocks = carry + blocks
        carry = []
        while len(blocks) > 1 and sum(b.h for b in blocks) + BASE_GAP * (len(blocks) - 1) > avail:
            carry.insert(0, blocks.pop())
        if blocks:
            out.append(blocks)
    while carry:
        out.append([carry.pop(0)])
    i = 0
    while i < len(out) - 1:
        nxt = out[i + 1][0]
        if sum(b.h for b in out[i]) + BASE_GAP * len(out[i]) + nxt.h <= avail:
            out[i].append(out[i + 1].pop(0))
            if not out[i + 1]:
                out.pop(i + 1)
            continue
        i += 1

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"{d['company']} {d['year']} 연간 PR 리포트")
    c.setAuthor("오픈피알")
    p = _Page(c)
    last = d["last_month"]
    end_day = (date(d["year"] + (last == 12), last % 12 + 1, 1) - date.resolution).day
    for n, blocks in enumerate(out, 1):
        # 머리글
        x = p.text(X0, TOP, d["company"], font=BOLD, size=22)
        p.text(x + 7, TOP, f"{d['year']} 연간 PR 리포트", font=MED, size=22, color=SUBTITLE)
        p.text(X0 + W, TOP + 12, f"분석 기간 {d['year']}.1.1 – {last}.{end_day}", size=8.8, color=MUTED, align="right")
        p.hline(X0, TOP + 32, W, color=INK, lw=2)
        # 구역: 남는 높이를 사이 간격으로 고르게
        used = sum(b.h for b in blocks)
        gap = BASE_GAP if len(blocks) < 2 else min(MAX_GAP, max(BASE_GAP, (avail - used) / (len(blocks) - 1)))
        t = CONTENT_TOP
        for b in blocks:
            b.draw(p, t)
            t += b.h + gap
        # 꼬리말
        p.hline(X0, PAGE_H - 34, W)
        p.text(X0, PAGE_H - 26, "Prepared by OPENPR", size=8, color=MUTED)
        p.text(X0 + W, PAGE_H - 26, str(n), size=8, color=MUTED, align="right")
        c.showPage()
    c.save()
    return buf.getvalue()
