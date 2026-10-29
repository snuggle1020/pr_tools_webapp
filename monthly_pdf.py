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

from media_order import media_category

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


def _clip(text, n):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text if len(text) <= n else text[:n].rstrip(" ,·…") + "…"


def _topic_label(title):
    """단순 언급 주제를 짧게: 제목의 따옴표 속 핵심어와 맨 앞 회사명('LX판토스, ...')을 뽑는다.
    예) "LX판토스, ‘로봇 물류센터’ 만든다…" -> LX판토스 ‘로봇 물류센터’ / "전통시장서 …'AI 관광집사'가 …" -> ‘AI 관광집사’"""
    t = re.sub(r"\[[^\]]*\]", "", title or "").strip()
    quoted = re.search(r"[‘'\"“]([^’'\"”]{2,20})[’'\"”]", t)
    subject = re.match(r"([^\s,·…]{2,15}),", t)
    if quoted and subject:
        return f"{subject.group(1)} ‘{quoted.group(1)}’"
    if quoted:
        return f"‘{quoted.group(1)}’"
    if subject:
        return subject.group(1)
    return f"‘{_clip(t, 20)}’"


def _change(label, prev, cur):
    """형식별 증감을 보고서 말투로 (문장 끝 '습니다'는 빼고 줘서 '~했으며,'로 이어 붙일 수 있게).
    예) 인터뷰 기사 1건이 새롭게 게재됐 / 기획 기사는 1건으로 전월 수준을 유지했 / 기고는 3건으로 2건 증가했"""
    if not prev:
        return f"{label} {cur}건이 새롭게 게재됐"
    if not cur:
        return f"{label}는 이번 달 게재되지 않았"
    if prev == cur:
        return f"{label}는 {cur}건으로 전월 수준을 유지했"
    return f"{label}는 {cur}건으로 {abs(cur - prev)}건 {'증가' if cur > prev else '감소'}했"


def _deleted_note(data):
    """게재 후 삭제한 보도자료 알림 한 문장 (없으면 빈 문자열)."""
    gone = [f"‘{_clip(r['title'], 32)}’" for r in data["releases"] if r.get("deleted")]
    return f"{', '.join(gone)} 보도자료는 게재 후 삭제 요청해 집계에서 제외했습니다." if gone else ""


def _join(clauses):
    """'A했으며, B했습니다'처럼 한 문장으로 잇는다."""
    return "으며, ".join(clauses) + "습니다." if clauses else ""


def default_summary(data, release_dates=None):
    """이달의 요약 기본 문장 (규칙으로 만든 초안, 앱에서 담당자가 고쳐 씀).
    전월 건수가 있으면 형식별로 전월과 비교: 보도자료 배포 → 단순 언급 → 기획·인터뷰·기고.
    전월 건수가 없으면(1월이거나 기록 없음) 이달 내용만 정리한다."""
    m, t = data["month"], data["types"]
    tr = data["trend"]
    if not (tr and m > 1):
        return _summary_without_prev(data, release_dates)
    i = m - 2
    total = sum(t.values())
    prev_total = tr["release"][i] + tr["planned"][i] + tr["mention"][i]
    up, down = total > prev_total, total < prev_total
    rels = [r for r in data["releases"] if not r.get("deleted")]
    parts = []

    # 보도자료: 배포 유무가 바뀌었는지, 전체 증감을 이끌었는지
    p_rel, c_rel = tr["release"][i], t.get("보도자료", 0)
    led = (c_rel > p_rel and up) or (c_rel < p_rel and down)
    lead = f"전체 게재 수 {'증가' if up else '감소'}를 이끌었습니다."
    if p_rel == 0 and rels:
        parts.append(f"전월에는 보도자료 배포가 없었으나, {m}월 보도자료 {len(rels)}건을 배포해 관련 기사 {c_rel}건이 " +
                     (f"게재되며 {lead}" if led else "게재됐습니다."))
    elif p_rel and not rels:
        parts.append(f"{m}월은 보도자료 배포가 없어 관련 기사가 {p_rel}건에서 0건으로 " +
                     (f"줄며 {lead}" if led else "감소했습니다."))
    elif rels:
        head = f"{m}월은 보도자료 {len(rels)}건을 배포해 관련 기사 {c_rel}건이 게재되며, "
        if p_rel == c_rel:
            parts.append(f"{head}전월 수준을 유지했습니다.")
        else:
            how = f"전월 대비 {abs(c_rel - p_rel)}건 {'증가' if c_rel > p_rel else '감소'}"
            parts.append(f"{head}{how}" + (f"해 {lead}" if led else "했습니다."))
    else:
        parts.append(f"{m}월은 전월과 마찬가지로 보도자료 배포가 없었습니다.")
    parts.append(_deleted_note(data))

    # 단순 언급: 비슷하면 '비슷한 수준', 크게 늘면 어떤 소식에서 거론됐는지
    p_men, c_men = tr["mention"][i], t.get("단순언급", 0)
    big = [mn for mn in data["mentions"] if mn["count"] >= 2][:2]
    topics = ", ".join(f"{_topic_label(mn['title'])}({mn['count']}건)" for mn in big)
    if p_men == c_men:
        if c_men:
            parts.append(f"단순 언급은 {c_men}건으로 전월 수준을 유지했습니다.")
    elif not c_men:
        parts.append("단순 언급 기사는 이번 달 게재되지 않았습니다.")
    elif abs(c_men - p_men) <= max(2, p_men * 0.2):
        parts.append(f"단순 언급은 {c_men}건으로 전월과 비슷한 수준입니다.")
    elif c_men > p_men and topics:
        parts.append(f"단순 언급은 {topics} 소식에서 거론되며 {c_men}건으로 {c_men - p_men}건 증가했습니다.")
    else:
        parts.append(_join([_change("단순 언급", p_men, c_men).replace("단순 언급는", "단순 언급은")]))

    # 기획 · 인터뷰 · 기고: 이번 달이나 전월에 있었던 것만, 한 문장으로
    parts.append(_join([_change(label, tr[key][i], now) for label, key, now in [
        ("기획 기사", "plan", t.get("기획(오픈피알)", 0) + t.get("기획(외부)", 0)),
        ("인터뷰 기사", "interview", t.get("인터뷰", 0)),
        ("기고", "contrib", t.get("기고", 0))] if tr[key][i] or now]))
    return " ".join(x for x in parts if x)


def _summary_without_prev(data, release_dates=None):
    """전월 건수가 없을 때: ① 보도를 이끈 보도자료 ② 게재 매체 유형 ③ 단순 언급이 나온 주요 소식 ④ 기획 기사 유무."""
    m = data["month"]
    total = sum(data["types"].values())
    dates = [d for d, r in zip(list(release_dates or []), data["releases"]) if not r.get("deleted")]
    rels = [r for r in data["releases"] if not r.get("deleted")]
    parts = []

    # ① 보도를 이끈 것
    if rels:
        k, top = max(enumerate(rels), key=lambda x: x[1]["count"])
        d = dates[k] if k < len(dates) else None
        same_day = sum(1 for a in data["articles"] if d and a["date"] == d and assign_release(a, dates) == k)
        title = _clip(top["title"], 32)
        if len(rels) == 1:
            if d and same_day and same_day < top["count"]:
                how = f"배포 당일 {same_day}건을 포함해 {top['count']}건 게재되며"
            elif d and same_day:
                how = f"배포 당일 {top['count']}건 게재되며"
            else:
                how = f"{top['count']}건 게재되며"
            parts.append(f"{m}월은 ‘{title}’ 보도자료 {how} 보도를 이끌었습니다.")
        else:
            parts.append(f"{m}월은 보도자료 {len(rels)}건으로 {sum(r['count'] for r in rels)}건이 게재됐고, "
                         f"그중 ‘{title}’이(가) {top['count']}건으로 가장 많이 실렸습니다.")
    else:
        parts.append(f"{m}월은 보도자료 배포 없이 총 {total}건이 보도됐습니다.")
    parts.append(_deleted_note(data))

    # ② 게재 매체 유형
    cats = Counter(media_category(a["media"]) for a in data["articles"]).most_common()
    if len(data["articles"]) < 3:  # 기사가 몇 건 안 되면 매체 유형 이야기는 뺌
        pass
    elif len(cats) == 1:
        parts.append(f"게재 매체는 {cats[0][0]}({cats[0][1]}건)이었습니다.")
    elif cats:
        others = "·".join(c for c, _ in cats[1:4])
        parts.append(f"게재 매체는 {cats[0][0]}({cats[0][1]}건) 중심으로 {others}에도 실렸습니다.")

    # ③ 단순 언급 맥락
    big = [mn for mn in data["mentions"] if mn["count"] >= 2][:2]
    if big:
        topics = ", ".join(f"{_topic_label(mn['title'])}({mn['count']}건)" for mn in big)
        parts.append(f"단순 언급은 {topics} 소식에서 함께 거론됐습니다.")
    elif data["mentions"]:
        parts.append(f"단순 언급은 {data['types'].get('단순언급', 0)}건으로, 여러 소식에서 개별적으로 거론됐습니다.")

    # ④ 기획 기사
    planned = sum(data["types"].get(k, 0) for k in PLANNED_TYPES)
    parts.append("기획·인터뷰 기사는 없었습니다." if not planned else f"기획·인터뷰 기사도 {planned}건 실렸습니다.")
    return " ".join(x for x in parts if x)


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
    cells.append((str(sum(1 for r in d["releases"] if not r.get("deleted"))), "건", "보도자료 배포 건수", INK))
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
            deleted = r.get("deleted")
            color = GRAY if deleted else BLUE  # 보도자료 카드는 한 가지 색, 게재 후 삭제한 건만 회색
            x = X0 + j * (cw + gap)
            p.rect(x, top, cw, h, fill=WHITE, stroke=LINE, r=8)
            p.rect(x, top, 4, h, fill=color)
            t = top + 16
            for line in lines:
                p.text(x + 16, t, line, font=BOLD, size=11.5, color=MUTED if deleted else INK)
                t += 17
            dt = dates[k] if k < len(dates) else None
            if dt:
                p.text(x + 16, top + h - 20, f"{dt.month}월 {dt.day}일 ({WEEKDAYS[dt.weekday()]}) 배포"
                       + (" 후 삭제 요청" if deleted else ""), size=9.3, color=MUTED)
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
    doc.section(6, "주요 기사", "제목을 누르면 기사로 이동")
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
    rows = [{"cells": [(i, REG, MUTED), (m["title"], REG, INK), (m["media"], REG, INK), (f"{m['count']}건", MED, INK)],
             "link": (1, m["url"])} for i, m in enumerate(ms, 1)]
    _table(doc, [("No", 28, "center"), ("제목", None, "left"), ("대표 매체", 84, "left"),
                 ("관련 기사", 54, "center")], rows)  # 위 기사 표와 같은 글자 크기·칸 너비


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
