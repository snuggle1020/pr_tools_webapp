# -*- coding: utf-8 -*-
"""
월간 PR리포트 PDF 빌더.
다 채운 월간 PR리포트 엑셀(월 시트 + Summary 시트)을 읽어 A4 PDF로 만든다.
  1쪽 - 핵심 수치, 이달의 요약, 보도자료별 게재 현황, 기사 형식별 막대, 월별 추이, 기자 미팅, 다음 달 목표
  2쪽~ - 주요 기사 전체 표(보도자료별로 묶음, 제목 클릭 시 원문 이동), 단순 언급 기사(주제 + 관련 기사 수)

월 시트는 B열의 '1. 기사형식별 보도 수' 같은 구역 제목을 기준으로 읽는다.
그리기 도구(폰트·줄바꿈·좌표 래퍼)는 결과보고서(report_pdf)의 것을 그대로 쓴다.
"""
import base64
import io
import json
import re
from collections import Counter
from datetime import date, datetime, timedelta

import openpyxl
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as rl_canvas

from report_pdf import (
    BOLD, MED, REG, SEMI, INK, MUTED, SUBTITLE, LINE, BLUE, GREEN, GRAY, WHITE, WEEKDAYS,
    PAGE_H, X0, W, TOP, _Page, _register_fonts, _sw, _wrap, _clip_lines,
)

BG = HexColor("#F4F6F9")
ORANGE = HexColor("#D9822B")
RED = HexColor("#C2410C")
FAINT = HexColor("#9AA1AE")
BOTTOM_Y = PAGE_H - 46
PURPLE = HexColor("#7A4FD1")
# 월별 추이 막대: 아래에서부터 쌓는 순서 (키, 범례 이름, 색)
TREND_SERIES = [("release", "보도자료", BLUE), ("plan", "기획", GREEN), ("interview", "인터뷰", PURPLE),
                ("contrib", "기고", ORANGE), ("mention", "단순언급", GRAY)]  # 본문이 내려올 수 있는 한계 (아래는 꼬리말)

TYPE_ORDER = ["인터뷰", "기고", "기획(오픈피알)", "기획(외부)", "보도자료", "단순언급"]
SECTION_KEYS = {  # 구역 제목에 들어 있는 말 -> 구역 이름 (위에서부터 먼저 맞는 것; '단순 언급 주요 기사'라서 단순 언급이 먼저)
    "types": "기사형식별", "releases": "보도자료별", "mentions": "단순 언급",
    "articles": "주요 기사", "meetings": "기자 미팅", "plans": "목표",
}


# ---------------------------------------------------------------- 엑셀 읽기

def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", str(v or ""))
    return date(int(m[1]), int(m[2]), int(m[3])) if m else None


def _int(v):
    m = re.search(r"-?\d+", str(v if v is not None else ""))
    return int(m[0]) if m else 0


def _url(cell):
    if cell is None:
        return ""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    s = str(cell.value or "").strip()
    return s if s.startswith("http") else ""


def _sections(ws):
    """{구역 이름: {'title': 제목, 'rows': [{열 문자: 셀}]}} - 제목 바로 다음 줄(표 머리글)은 건너뜀."""
    out, cur, head_row = {}, None, 0
    for r in range(1, ws.max_row + 1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and re.match(r"\s*\d+\.\s", b):
            cur = next((k for k, word in SECTION_KEYS.items() if word in b), None)
            if cur:
                out[cur] = {"title": re.sub(r"^\s*\d+\.\s*", "", b).strip(), "rows": []}
                head_row = r + 1
            continue
        if isinstance(b, str) and b.startswith("참고."):  # 앱이 쓴 기록용 부록 (단순 언급 기사 전체) - 읽지 않음
            cur = None
            continue
        if cur is None or r == head_row:
            continue
        cells = {c.column_letter: c for c in ws[r] if c.value is not None and str(c.value).strip() != ""}
        if cells:
            out[cur]["rows"].append(cells)
    return out


def month_sheets(file):
    file.seek(0)
    wb = openpyxl.load_workbook(file, read_only=True)
    names = [s for s in wb.sheetnames if re.fullmatch(r"\d{1,2}월", s.strip())]
    wb.close()
    return names


def read_month_report(file, sheet_name):
    """월간 PR리포트 엑셀의 한 달 시트를 읽어 PDF에 필요한 값을 돌려준다."""
    file.seek(0)
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[sheet_name]
    month = int(re.sub(r"\D", "", sheet_name))
    sec = _sections(ws)
    val = lambda row, col: row[col].value if col in row else None

    title = str(ws["B2"].value or "")
    company = re.sub(r"\s*\d{1,2}월.*$", "", title).strip() or title

    types = {k: 0 for k in TYPE_ORDER}
    for r in sec.get("types", {}).get("rows", []):
        name = str(val(r, "B") or "").strip()
        if name and name != "합계" and not name.startswith("*"):
            types[name] = _int(val(r, "C"))

    releases = [{"title": str(val(r, "C")).strip(), "count": _int(val(r, "F"))}
                for r in sec.get("releases", {}).get("rows", []) if val(r, "C")]

    arts = []
    for r in sec.get("articles", {}).get("rows", []):
        if not str(val(r, "B") or "").strip().isdigit() or not val(r, "D"):
            continue
        arts.append({"media": str(val(r, "C") or "").strip(), "title": str(val(r, "D")).strip(),
                     "url": _url(r.get("G")) or _url(r.get("D")), "date": _to_date(val(r, "J")),
                     "kind": str(val(r, "K") or "").strip()})

    mentions = []
    for r in sec.get("mentions", {}).get("rows", []):
        if not str(val(r, "B") or "").strip().isdigit() or not val(r, "D"):
            continue
        mentions.append({"media": str(val(r, "C") or "").strip(), "title": str(val(r, "D")).strip(),
                         "url": _url(r.get("G")) or _url(r.get("D")), "count": _int(val(r, "J")) or 1})

    meetings = [(str(val(r, "B") or ""), str(val(r, "C") or ""), str(val(r, "D") or ""))
                for r in sec.get("meetings", {}).get("rows", []) if not str(val(r, "B") or "").startswith("*")]
    plans = [(str(val(r, "B") or ""), str(val(r, "C") or ""), str(val(r, "E") or ""))
             for r in sec.get("plans", {}).get("rows", []) if not str(val(r, "B") or "").startswith("*")]
    plan_title = sec.get("plans", {}).get("title") or f"{month % 12 + 1}월 PR 목표"

    # Summary 시트: 월별 형식 추이 (이번 달 칸이 아직 비어 있으면 월 시트 값으로 채움)
    year, rows = summary_rows(wb)
    trend = None
    if rows:
        if sum(rows[k][month - 1] for k in TYPE_ORDER) == 0:
            for k in TYPE_ORDER:
                rows[k][month - 1] = types[k]
        trend = trend_from_rows(rows)
    if year is None:
        dates = [a["date"] for a in arts if a["date"]]
        year = dates[0].year if dates else date.today().year
    wb.close()

    return {"company": company, "year": year, "month": month, "types": types, "releases": releases,
            "articles": arts, "mentions": mentions, "meetings": meetings, "plans": plans,
            "plan_title": plan_title, "trend": trend}


PLANNED_TYPES = ("인터뷰", "기고", "기획(오픈피알)", "기획(외부)")


def summary_rows(wb):
    """Summary 시트의 형식별 1~12월 값. 반환: (연도|None, {형식: [12개]}|None)"""
    sm_name = next((n for n in wb.sheetnames if n.lower().startswith("summary")), None)
    if not sm_name:
        return None, None
    sm = wb[sm_name]
    m = re.search(r"(\d{4})", str(sm["B2"].value or "") + sm_name)
    rows = {}
    for r in range(1, sm.max_row + 1):
        name = str(sm.cell(r, 2).value or "").strip()
        if name in TYPE_ORDER:
            rows[name] = [_int(sm.cell(r, c).value) for c in range(3, 15)]
    if not rows:
        return (int(m[1]) if m else None), None
    for k in TYPE_ORDER:
        rows.setdefault(k, [0] * 12)
    return (int(m[1]) if m else None), rows


def trend_from_rows(rows):
    """{형식: [12개]} -> 막대 그래프용. 'planned'(기획·인터뷰·기고 합)는 합계 계산용,
    'interview','contrib','plan'(기획 오픈피알+외부)은 막대를 형식별로 나눠 그리는 데 쓴다."""
    return {
        "release": list(rows["보도자료"]),
        "planned": [sum(rows[k][i] for k in PLANNED_TYPES) for i in range(12)],
        "mention": list(rows["단순언급"]),
        "interview": list(rows["인터뷰"]),
        "contrib": list(rows["기고"]),
        "plan": [rows["기획(오픈피알)"][i] + rows["기획(외부)"][i] for i in range(12)],
    }


def assign_release(article, dates):
    """보도자료 기사가 몇 번째 보도자료 건인지 (배포일이 기사 게재일 이전인 것 중 가장 늦은 것). 없으면 None."""
    if article.get("kind") != "보도자료" or not article.get("date"):
        return None
    cand = [i for i, d in enumerate(dates) if d and d <= article["date"]]
    return cand[-1] if cand else None


def guess_release_dates(data):
    """보도자료 기사가 가장 많이 몰린 날짜들을 보도자료 수만큼 골라 배포일로 추정 (빠른 날짜부터)."""
    n = len(data["releases"])
    counts = Counter(a["date"] for a in data["articles"] if a["kind"] == "보도자료" and a["date"])
    top = sorted(d for d, _ in counts.most_common(n))
    first = date(data["year"], data["month"], 1)
    return top + [first] * (n - len(top))


def default_summary(data):
    total = sum(data["types"].values())
    parts = []
    if data["releases"]:
        parts.append(f"보도자료 {len(data['releases'])}건을 배포해 총 {total}건이 보도됐습니다.")
    else:
        parts.append(f"이번 달은 보도자료 배포 없이 총 {total}건이 보도됐습니다.")
    tr = data["trend"]
    if tr and data["month"] > 1:
        i = data["month"] - 2
        prev = tr["release"][i] + tr["planned"][i] + tr["mention"][i]
        d = total - prev
        parts.append(f"전월({prev}건) 대비 {abs(d)}건 {'증가' if d >= 0 else '감소'}했습니다.")
    return " ".join(parts)


# ---------------------------------------------------------------- 그리기

class _Doc:
    def __init__(self, c, data):
        self.c, self.p, self.data = c, _Page(c), data
        self.page_no = 1
        self.top = TOP

    def header(self):
        """모든 쪽 같은 머리글: 기업명 / N월 PR 리포트 + 오른쪽 분석 기간"""
        p, d = self.p, self.data
        size = 22
        x = p.text(X0, TOP, d["company"], font=BOLD, size=size)
        p.text(x + 7, TOP, f"{d['month']}월 PR 리포트", font=MED, size=size, color=SUBTITLE)
        last = (date(d["year"], d["month"] % 12 + 1, 1) if d["month"] < 12 else date(d["year"] + 1, 1, 1)) - timedelta(days=1)
        p.text(X0 + W, TOP + size - 10, f"분석 기간 {d['year']}.{d['month']}.1 – {d['month']}.{last.day}",
               size=8.8, color=MUTED, align="right")
        t = TOP + size + 10
        p.hline(X0, t, W, color=INK, lw=2)
        self.top = t + 16

    def footer(self):
        p = self.p
        p.hline(X0, PAGE_H - 34, W)
        p.text(X0, PAGE_H - 26, "Prepared by OPENPR", size=8, color=MUTED)
        p.text(X0 + W, PAGE_H - 26, str(self.page_no), size=8, color=MUTED, align="right")

    def new_page(self):
        self.footer()
        self.c.showPage()
        self.page_no += 1
        self.header()

    def ensure(self, h):
        if self.top + h > BOTTOM_Y:
            self.new_page()

    def section(self, num, title, note="", x=X0, w=W):
        p = self.p
        nx = p.text(x, self.top + 2, f"{num:02d}", font=BOLD, size=11, color=BLUE)
        p.text(nx + 7, self.top, title, font=BOLD, size=14)
        if note:
            p.text(x + w, self.top + 4, note, size=9, color=MUTED, align="right")
        return self.top + 26


def _draw_kpis(doc):
    p, d = doc.p, doc.data
    total = sum(d["types"].values())
    tr = d["trend"]
    cells = [(str(total), "건", "이번 달 총 보도 건수", BLUE)]
    if tr and d["month"] > 1:
        i = d["month"] - 2
        prev = tr["release"][i] + tr["planned"][i] + tr["mention"][i]
        diff = total - prev
        cells.append((f"{'+' if diff >= 0 else '−'}{abs(diff)}", "건", "전월 대비 증감",
                      GREEN if diff >= 0 else RED))
    cells.append((str(len(d["releases"])), "건", "보도자료 배포 건수", INK))
    if tr:
        ytd = sum(tr["release"][i] + tr["planned"][i] + tr["mention"][i] for i in range(d["month"]))
        cells.append((str(ytd), "건", "연간 누적 건수", INK))

    top, h = doc.top, 66
    p.rect(X0, top, W, h, fill=WHITE, stroke=LINE, r=9)
    cw = W / len(cells)
    for i, (num, unit, label, color) in enumerate(cells):  # 숫자+단위 묶음과 설명을 칸 가운데에
        mid = X0 + i * cw + cw / 2
        x = p.text(mid - (_sw(num, BOLD, 25) + 2 + _sw(unit, SEMI, 10.5)) / 2, top + 13, num, font=BOLD, size=25, color=color)
        p.text(x + 2, top + 24, unit, font=SEMI, size=10.5, color=MUTED)
        p.text(mid, top + 45, label, font=MED, size=9.5, color=MUTED, align="center")
        if i:
            p.vline(X0 + i * cw, top, h)
    doc.top = top + h + 12


def _draw_summary(doc, text):
    if not text.strip():
        return
    p = doc.p
    lines = []
    for para in text.strip().splitlines():
        lines += _wrap(para.strip(), REG, 10.3, W - 32) if para.strip() else [""]
    h = 14 + 14 + 16.5 * len(lines) + 10
    p.rect(X0, doc.top, W, h, fill=BG, r=8)
    p.text(X0 + 16, doc.top + 13, "이달의 요약", font=SEMI, size=9.3, color=BLUE)
    t = doc.top + 31
    for line in lines:
        p.text(X0 + 16, t, line, size=10.3)
        t += 16.5
    doc.top += h + 8


def _draw_releases(doc, dates):
    p, d = doc.p, doc.data
    rels = d["releases"]
    doc.top += 14
    top = doc.section(1, "보도자료별 게재 현황")
    if not rels:
        p.rect(X0, top, W, 40, fill=BG, r=8)
        p.text(X0 + 16, top + 14, "이번 달 배포한 보도자료가 없습니다.", size=10, color=MUTED)
        doc.top = top + 40
        return
    gap = 10
    cw = (W - gap) / 2
    for i in range(0, len(rels), 2):
        pair = rels[i:i + 2]
        wrapped = [_clip_lines(_wrap(r["title"], BOLD, 11.5, cw - 34), BOLD, 11.5, cw - 34, 2) for r in pair]
        h = 16 + 17 * max(len(w) for w in wrapped) + 40
        for j, (r, lines) in enumerate(zip(pair, wrapped)):
            k = i + j
            color = BLUE  # 보도자료 카드는 한 가지 색
            x = X0 + j * (cw + gap)
            p.rect(x, top, cw, h, fill=WHITE, stroke=LINE, r=8)
            p.rect(x, top, 4, h, fill=color)
            t = top + 16
            for line in lines:
                p.text(x + 16, t, line, font=BOLD, size=11.5)
                t += 17
            dt = dates[k] if k < len(dates) else None
            if dt:
                p.text(x + 16, top + h - 20, f"{dt.month}월 {dt.day}일 ({WEEKDAYS[dt.weekday()]}) 배포", size=9.3, color=MUTED)
            p.text(x + cw - 14, top + h - 21, "건", font=SEMI, size=10.5, color=MUTED, align="right")
            p.text(x + cw - 14 - _sw("건", SEMI, 10.5) - 3, top + h - 35, str(r["count"]), font=BOLD, size=25,
                   color=color, align="right")
        top += h + 8
    doc.top = top - 8


def _draw_types_and_trend(doc):
    p, d = doc.p, doc.data
    doc.top += 22
    gap = 14
    cw = (W - gap) / 2
    h = 170
    doc.ensure(26 + h)
    top = doc.top
    doc.section(2, "기사 형식별 보도 수", x=X0, w=cw)
    box = top + 26
    # 형식별 막대
    p.rect(X0, box, cw, h, fill=WHITE, stroke=LINE, r=8)
    types = d["types"]
    mx = max(types.values()) or 1
    t = box + 14
    for name in TYPE_ORDER:
        n = types.get(name, 0)
        col = FAINT if not n else INK
        p.text(X0 + 14, t, name, font=MED, size=9.5, color=col)
        bx, bw = X0 + 92, cw - 92 - 50
        p.rect(bx, t + 2.5, bw, 7, fill=BG, r=3.5)
        if n:
            p.rect(bx, t + 2.5, max(7, bw * n / mx), 7, fill=BLUE if n == mx else GRAY, r=3.5)
        p.text(X0 + cw - 14, t, f"{n}건", font=BOLD, size=9.5, color=col, align="right")
        t += 20
    p.hline(X0 + 14, t + 2, cw - 28)
    p.text(X0 + 14, t + 11, "합계", font=BOLD, size=10)
    p.text(X0 + cw - 14, t + 11, f"{sum(types.values())}건", font=BOLD, size=10, align="right")

    # 월별 추이
    x0 = X0 + cw + gap
    doc.section(3, "월별 보도 추이", x=x0, w=cw)
    p.rect(x0, box, cw, h, fill=WHITE, stroke=LINE, r=8)
    tr = d["trend"]
    if not tr:
        p.text(x0 + 14, box + 16, "Summary 시트가 없어 추이를 그리지 않았습니다.", size=9, color=MUTED)
    else:
        m = d["month"]
        tot = [tr["release"][i] + tr["planned"][i] + tr["mention"][i] for i in range(m)]
        mx = max(tot) or 1
        chart_h, base = 92, box + 128
        slot = (cw - 24) / m
        bw = min(22, slot * 0.6)
        for i in range(m):
            cx = x0 + 12 + slot * i + slot / 2
            cur = i == m - 1
            y = base
            for key, _, color in TREND_SERIES:
                seg = tr[key][i] / mx * chart_h
                if seg > 0:
                    p.rect(cx - bw / 2, y - seg, bw, seg, fill=color)
                    y -= seg
            if cur:
                p.c.setStrokeColor(INK)
                p.c.setLineWidth(1.4)
                p.c.rect(cx - bw / 2 - 2, p.y(base) - 2 + 0, bw + 4, (base - y) + 4, stroke=1, fill=0)
            p.text(cx, y - 13, str(tot[i]), font=BOLD if cur else MED, size=8.5, color=INK if cur else MUTED, align="center")
            p.text(cx, base + 5, f"{i + 1}월", font=BOLD if cur else REG, size=8.5, align="center")
        lx = x0 + 14
        for _, name, color in TREND_SERIES:
            p.rect(lx, box + h - 17, 7, 7, fill=color, r=1.5)
            lx = p.text(lx + 10, box + h - 18.5, name, size=8.3, color=MUTED) + 9
    doc.top = box + h


def _table(doc, cols, rows, x=X0, w=W, font_size=9, lead=13):
    """cols: [(머리글, 폭, 정렬)] - 폭 None 은 남는 폭.
    rows: {'cells': [(글, 폰트, 색)], 'link': (칸 번호, url)}.
    페이지를 넘기면 머리글을 다시 그린다."""
    p = doc.p
    fixed = sum(c[1] for c in cols if c[1])
    widths = [c[1] or (w - fixed) for c in cols]
    xs = [x + sum(widths[:i]) for i in range(len(cols))]

    def head():
        p.rect(x, doc.top, w, 22, fill=INK)
        for (name, _, align), cx, cw in zip(cols, xs, widths):
            tx = cx + 7 if align == "left" else cx + cw / 2 if align == "center" else cx + cw - 7
            p.text(tx, doc.top + 6.5, name, font=SEMI, size=8.8, color=WHITE, align=align)
        doc.top += 22

    head()
    for r in rows:
        wrapped = []
        for (txt, font, _), cw in zip(r["cells"], widths):
            wrapped.append(_wrap(str(txt), font, font_size, cw - 14))
        n = max(len(l) for l in wrapped)
        h = 10 + n * lead
        if doc.top + h > BOTTOM_Y:
            doc.new_page()
            head()
        for (txt, font, color), lines, (_, _, align), cx, cw in zip(r["cells"], wrapped, cols, xs, widths):
            tx = cx + 7 if align == "left" else cx + cw / 2 if align == "center" else cx + cw - 7
            t = doc.top + 5 + (n - len(lines)) * lead / 2 + 1
            for line in lines:
                p.text(tx, t, line, font=font, size=font_size, color=color, align=align)
                t += lead
        if r.get("link"):
            k, url = r["link"]
            p.link(url, xs[k], doc.top, widths[k], h)
        doc.top += h
        p.hline(x, doc.top, w)


def _draw_meetings_and_plans(doc):
    d = doc.data
    gap = 14
    cw = (W - gap) / 2
    doc.top += 22
    doc.ensure(26 + 22 + 30 * max(len(d["meetings"]), len(d["plans"]), 1))
    start = doc.top
    page = doc.page_no

    doc.section(4, "기자 미팅", x=X0, w=cw)
    doc.top = start + 26
    if d["meetings"]:
        rows = [{"cells": [(a, SEMI, INK), (b, REG, INK), (c, REG, INK)]} for a, b, c in d["meetings"]]
        _table(doc, [("매체", 82, "left"), ("기자", 50, "left"), ("주요 내용", None, "left")], rows, x=X0, w=cw)
    else:
        doc.p.text(X0, doc.top, "이번 달 정리할 미팅이 없습니다.", size=9.5, color=MUTED)
        doc.top += 16
    left_end = doc.top

    if doc.page_no != page:  # 왼쪽 표가 페이지를 넘겼으면 오른쪽 표는 새 줄에 전체 폭으로
        doc.top += 22
        x, w = X0, W
        doc.ensure(60)
        start2 = doc.top
    else:
        x, w, start2 = X0 + cw + gap, cw, start
    doc.top = start2
    doc.section(5, d["plan_title"], x=x, w=w)
    doc.top = start2 + 26
    if d["plans"]:
        rows = [{"cells": [(a, SEMI, INK), (b, SEMI, INK), (c, REG, INK)]} for a, b, c in d["plans"]]
        _table(doc, [("형식", 54, "left"), ("내용", None, "left"), ("진행 상황", 86, "left")], rows, x=x, w=w)
    else:
        doc.p.text(x, doc.top, "등록된 목표가 없습니다.", size=9.5, color=MUTED)
        doc.top += 16
    doc.top = max(doc.top, left_end) if doc.page_no == page else doc.top


def _draw_articles(doc, dates):
    """기사 전체를 게재일·매체 순으로 한 표에 (보도자료별 묶음 줄 없이)"""
    arts = doc.data["articles"]
    doc.section(6, "주요 기사 리스트", "제목을 누르면 기사로 이동")
    doc.top += 26
    rows = []
    for no, a in enumerate(arts, 1):
        dt = f"{a['date'].month}.{a['date'].day}" if a["date"] else ""
        rows.append({"cells": [(no, REG, MUTED), (a["media"], SEMI, INK), (a["title"], REG, INK),
                               (dt, REG, MUTED), (a["kind"] or "-", MED, MUTED)],
                     "link": (2, a["url"])})
    if not arts:
        doc.p.text(X0, doc.top, "정리된 기사가 없습니다.", size=9.5, color=MUTED)
        doc.top += 16
        return
    _table(doc, [("No", 28, "center"), ("매체", 84, "left"), ("제목", None, "left"),
                 ("게재일", 40, "center"), ("형식", 54, "center")], rows)


def _draw_mentions(doc):
    d = doc.data
    ms = d["mentions"]
    doc.top += 24
    doc.ensure(26 + 22 + 30)
    doc.section(7, "단순 언급 기사")
    doc.top += 26
    if not ms:
        doc.p.text(X0, doc.top, "이번 달 단순 언급 기사가 없습니다.", size=9.5, color=MUTED)
        doc.top += 16
        return
    rows = [{"cells": [(i, REG, MUTED), (m["title"], SEMI, INK), (m["media"], REG, INK), (f"{m['count']}건", MED, INK)],
             "link": (1, m["url"])} for i, m in enumerate(ms, 1)]
    _table(doc, [("No", 28, "center"), ("제목", None, "left"), ("대표 매체", 84, "left"),
                 ("관련 기사", 54, "center")], rows)  # 위 기사 표와 같은 글자 크기·칸 너비
    doc.top += 8
    doc.p.text(X0, doc.top, "* 단순 언급: 본문에 함께 언급된 기사. 같은 주제의 기사는 하나로 묶어 대표 기사와 관련 기사 수로 표시",
               size=8.3, color=MUTED)


HISTORY_PREFIX = "prhist:"


def read_history(pdf_bytes):
    """앱이 만든 월간 리포트 PDF에 숨겨 둔 월별 기록. 반환: {'company','year','month','rows': {형식: [12]}} 또는 None."""
    try:
        from pypdf import PdfReader
        kw = (PdfReader(io.BytesIO(pdf_bytes)).metadata or {}).get("/Keywords") or ""
        if not kw.startswith(HISTORY_PREFIX):
            return None
        h = json.loads(base64.b64decode(kw[len(HISTORY_PREFIX):]).decode("utf-8"))
        h["rows"] = {k: [int(x) for x in h["rows"].get(k, [0] * 12)] for k in TYPE_ORDER}
        h["year"], h["month"] = int(h["year"]), int(h["month"])
        return h
    except Exception:
        return None


def build_monthly_pdf(data, release_dates, summary_text=""):
    """data: read_month_report() / monthly_report.compose() 결과, release_dates: 보도자료 순서대로 배포일(date).
    data['history'] 가 있으면 다음 달에 이 PDF만 올려도 월별 추이가 이어지도록 PDF 정보(키워드)에 숨겨 넣는다.
    반환값: PDF bytes."""
    _register_fonts()
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"{data['company']} {data['year']}년 {data['month']}월 PR 리포트")
    c.setAuthor("오픈피알")
    if data.get("history"):
        raw = json.dumps(data["history"], ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        c.setKeywords(HISTORY_PREFIX + base64.b64encode(raw).decode("ascii"))
    doc = _Doc(c, data)

    doc.header()
    _draw_kpis(doc)
    _draw_summary(doc, summary_text)
    _draw_releases(doc, release_dates)
    _draw_types_and_trend(doc)
    _draw_meetings_and_plans(doc)

    doc.new_page()
    _draw_articles(doc, release_dates)
    _draw_mentions(doc)
    doc.footer()
    c.save()
    return buf.getvalue()
