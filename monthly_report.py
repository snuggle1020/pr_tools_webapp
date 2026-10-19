# -*- coding: utf-8 -*-
"""
월간 PR리포트를 앱 안에서 편집해 PDF로 만드는 흐름의 데이터 처리.
  - 수집한 기사에 첫 형식 넣기 / 보도자료 목록 추정
  - 앱에서 고친 기사·보도자료·미팅·목표 -> PDF용 데이터(monthly_pdf.build_monthly_pdf 입력) 계산
  - 기록용 엑셀: 기존 PR리포트 파일에 이번 달 시트를 기존과 같은 칸 배치로 추가하고 Summary 시트 갱신
  - 작업 파일(.json) 저장 / 불러오기 (수집한 기사 목록까지 통째로 저장해서 다시 검색하지 않음)
"""
import io
import json
from collections import Counter
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from media_order import media_sort_key
from monthly_pdf import PLANNED_TYPES, TYPE_ORDER, assign_release, summary_rows, trend_from_rows

WORK_VERSION = 1
KIND_OPTIONS = TYPE_ORDER  # 앱 '형식' 드롭다운


# ---------------------------------------------------------------- 수집 직후

def initial_kind(title, keyword, auto_result):
    """자동 분류가 보도자료면 보도자료, 제목에 키워드가 없으면 단순언급, 나머지는 담당자가 고르도록 공란."""
    if auto_result == "보도자료":
        return "보도자료"
    if keyword and keyword not in title:
        return "단순언급"
    return ""


def guess_releases(articles, keyword):
    """보도자료 기사가 3건 이상 몰린 날짜마다 보도자료 1건으로 보고, 그날 기사 중 키워드로 시작하는 첫 제목을 붙인다.
    앞 배포일에서 3일 안에 몰린 날짜는 같은 보도자료의 후속 게재로 본다."""
    by_date = Counter(a["date"] for a in articles if a["kind"] == "보도자료" and a["date"])
    dates = []
    for d in sorted(d for d, n in by_date.items() if n >= 3):
        if not dates or (d - dates[-1]).days > 3:
            dates.append(d)
    out = []
    for d in dates:
        same = [a for a in articles if a["kind"] == "보도자료" and a["date"] == d]
        lead = next((a for a in same if a["title"].startswith(keyword)), same[0])
        out.append({"title": lead["title"], "date": d})
    return out


def sort_articles(articles):
    return sorted(articles, key=lambda a: (a["date"] or date.max, media_sort_key(a["media"])))


# ---------------------------------------------------------------- 편집 결과 -> PDF 데이터

def compose(meta, articles, releases, meetings, plans, trend_rows):
    """meta: {'company','year','month'} / articles: [{'url','media','title','date','kind','include','group'}]
    releases: [{'title','date'}] / meetings, plans: [(a,b,c)] / trend_rows: Summary 값 {형식: [12]} 또는 None.
    반환: (monthly_pdf 용 data, 보도자료 배포일 목록, 형식 미지정 기사 수)"""
    used = [a for a in articles if a["include"] and a["kind"]]
    blank = sum(1 for a in articles if a["include"] and not a["kind"])
    dates = [r["date"] for r in releases]

    types = {k: 0 for k in TYPE_ORDER}
    for a in used:
        types[a["kind"]] = types.get(a["kind"], 0) + 1

    rel_counts = Counter(assign_release(a, dates) for a in used)
    rels = [{"title": r["title"], "count": rel_counts.get(i, 0)} for i, r in enumerate(releases)]

    main = [{"media": a["media"], "title": a["title"], "url": a["url"], "date": a["date"], "kind": a["kind"]}
            for a in sort_articles(a for a in used if a["kind"] != "단순언급")]

    groups = {}
    for a in used:
        if a["kind"] == "단순언급":
            groups.setdefault((a.get("group") or "").strip() or a["title"], []).append(a)
    mentions = []
    for members in groups.values():
        lead = min(members, key=lambda a: (media_sort_key(a["media"]), a["date"] or date.max))
        mentions.append({"media": lead["media"], "title": lead["title"], "url": lead["url"],
                         "count": len(members), "first": min(a["date"] or date.max for a in members)})
    mentions.sort(key=lambda m: (-m["count"], m["first"]))

    trend = None
    if trend_rows:
        rows = {k: list(v) for k, v in trend_rows.items()}
        for k in TYPE_ORDER:
            rows[k][meta["month"] - 1] = types[k]
        trend = trend_from_rows(rows)

    nm = meta["month"] % 12 + 1
    data = {"company": meta["company"], "year": meta["year"], "month": meta["month"], "types": types,
            "releases": rels, "articles": main, "mentions": mentions,
            "meetings": [m for m in meetings if any(m)], "plans": [p for p in plans if any(p)],
            "plan_title": f"{nm}월 PR 목표", "trend": trend}
    return data, dates, blank


# ---------------------------------------------------------------- 기록용 엑셀

THIN = Side(style="thin", color="D0D4DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="E8EEFC")
TITLE_FONT = Font(name="맑은 고딕", size=16, bold=True)
SEC_FONT = Font(name="맑은 고딕", size=12, bold=True)
HEAD_FONT = Font(name="맑은 고딕", size=10, bold=True)
BODY_FONT = Font(name="맑은 고딕", size=10)
LINK_FONT = Font(name="맑은 고딕", size=10, color="1F4ED8", underline="single")


def _cells(ws, row, spec, head=False):
    """spec: [(시작열, 끝열, 값)] - 시작~끝 병합, 테두리·글꼴 적용. 값이 ('link', url) 이면 하이퍼링크."""
    for c1, c2, v in spec:
        cell = ws.cell(row, c1)
        if isinstance(v, tuple) and v[0] == "link":
            cell.value = v[1]
            if v[1]:
                cell.hyperlink = v[1]
            cell.font = LINK_FONT
        else:
            cell.value = v
            cell.font = HEAD_FONT if head else BODY_FONT
        for c in range(c1, c2 + 1):
            ws.cell(row, c).border = BOX
            if head:
                ws.cell(row, c).fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center" if head else "left", vertical="center", wrap_text=not head)
        if c2 > c1:
            ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)


def _write_month_sheet(ws, data, articles_all):
    """기존 월 시트와 같은 칸 배치(B~K열, '1.'~'6.' 구역)로 이번 달 내용을 쓴다 -> monthly_pdf.read_month_report 로 다시 읽힘."""
    m = data["month"]
    ws["B2"] = f"{data['company']} {m}월 PR리포트"
    ws["B2"].font = TITLE_FONT
    r = 4

    def section(title):
        nonlocal r
        ws.cell(r, 2, title).font = SEC_FONT
        r += 1

    section("1. 기사형식별 보도 수")
    _cells(ws, r, [(2, 2, ""), (3, 3, "게재 수")], head=True); r += 1
    for k in TYPE_ORDER:
        _cells(ws, r, [(2, 2, k), (3, 3, data["types"][k])]); r += 1
    _cells(ws, r, [(2, 2, "합계"), (3, 3, sum(data["types"].values()))], head=True); r += 2

    section("2. 보도자료별 게재 현황")
    _cells(ws, r, [(2, 2, "No."), (3, 5, "주제"), (6, 6, "게재 수")], head=True); r += 1
    for i, rel in enumerate(data["releases"], 1):
        _cells(ws, r, [(2, 2, i), (3, 5, rel["title"]), (6, 6, rel["count"])]); r += 1
    r += 1

    section("3. 주요 기사 리스트")
    _cells(ws, r, [(2, 2, "No."), (3, 3, "매체"), (4, 6, "제목"), (7, 9, "URL"), (10, 10, "날짜"), (11, 11, "기사형식")], head=True)
    r += 1
    for i, a in enumerate(data["articles"], 1):
        d = a["date"].isoformat() if a["date"] else ""
        _cells(ws, r, [(2, 2, i), (3, 3, a["media"]), (4, 6, a["title"]), (7, 9, ("link", a["url"])),
                       (10, 10, d), (11, 11, a["kind"])]); r += 1
    r += 1

    section("4. 단순 언급 주요 기사")
    _cells(ws, r, [(2, 2, "No."), (3, 3, "매체"), (4, 6, "제목"), (7, 9, "URL"), (10, 10, "관련 기사 수"), (11, 11, "기사형식")], head=True)
    r += 1
    for i, mn in enumerate(data["mentions"], 1):
        _cells(ws, r, [(2, 2, i), (3, 3, mn["media"]), (4, 6, mn["title"]), (7, 9, ("link", mn["url"])),
                       (10, 10, f"{mn['count']}건"), (11, 11, "")]); r += 1
    r += 1

    section("5. 기자 미팅 리스트")
    _cells(ws, r, [(2, 2, "매체"), (3, 3, "기자명"), (4, 11, "주요 내용")], head=True); r += 1
    for a, b, c in data["meetings"]:
        _cells(ws, r, [(2, 2, a), (3, 3, b), (4, 11, c)]); r += 1
    r += 1

    section(f"6. {data['plan_title']}")
    _cells(ws, r, [(2, 2, "기사형식"), (3, 4, "내용"), (5, 6, "진행 상황")], head=True); r += 1
    for a, b, c in data["plans"]:
        _cells(ws, r, [(2, 2, a), (3, 4, b), (5, 6, c)]); r += 1
    r += 1

    # 단순 언급 기사 원본 목록 (기록용 - PDF에는 묶음만 나감)
    mention_arts = [a for a in sort_articles(articles_all) if a["include"] and a["kind"] == "단순언급"]
    if mention_arts:
        ws.cell(r, 2, "참고. 단순 언급 기사 전체").font = SEC_FONT
        r += 1
        _cells(ws, r, [(2, 2, "No."), (3, 3, "매체"), (4, 6, "제목"), (7, 9, "URL"), (10, 10, "날짜"), (11, 11, "묶음")], head=True)
        r += 1
        for i, a in enumerate(mention_arts, 1):
            _cells(ws, r, [(2, 2, i), (3, 3, a["media"]), (4, 6, a["title"]), (7, 9, ("link", a["url"])),
                           (10, 10, a["date"].isoformat() if a["date"] else ""), (11, 11, a.get("group") or "")]); r += 1

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 16
    for col in "CDEFGHIJK":
        ws.column_dimensions[col].width = 16
    ws.sheet_view.showGridLines = False


def _update_summary(wb, data):
    sm_name = next((n for n in wb.sheetnames if n.lower().startswith("summary")), None)
    if sm_name is None:
        sm = wb.create_sheet(f"Summary ({data['year']})", 0)
        sm["B2"] = f"{data['year']}년 월별 커버리지"
        sm["B2"].font = SEC_FONT
        _cells(sm, 3, [(2, 2, "")] + [(3 + i, 3 + i, f"{i + 1}월") for i in range(12)] + [(15, 15, "합계"), (16, 16, "비율")], head=True)
        for j, k in enumerate(TYPE_ORDER + ["합계"]):
            row = 4 + j
            _cells(sm, row, [(2, 2, k)] + [(3 + i, 3 + i, 0) for i in range(12)], head=(k == "합계"))
            _cells(sm, row, [(15, 15, f"=SUM(C{row}:N{row})"), (16, 16, f"=IFERROR(O{row}/O10,0)")], head=(k == "합계"))
            if k == "합계":
                for i in range(12):
                    col = openpyxl.utils.get_column_letter(3 + i)
                    sm.cell(row, 3 + i).value = f"=SUM({col}4:{col}9)"
        sm.column_dimensions["B"].width = 16
        return
    sm = wb[sm_name]
    col = 2 + data["month"]
    for r in range(1, sm.max_row + 1):
        name = str(sm.cell(r, 2).value or "").strip()
        cell = sm.cell(r, col)
        if name in data["types"]:
            cell.value = data["types"][name]
        elif name == "합계" and not str(cell.value or "").startswith("="):
            cell.value = sum(data["types"].values())


def build_record_xlsx(base_bytes, data, articles_all):
    """base_bytes: 기존 PR리포트 엑셀(없으면 새 파일). 이번 달 시트를 Summary 바로 뒤에 넣고(같은 이름은 교체) Summary 갱신."""
    if base_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(base_bytes))
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    name = f"{data['month']}월"
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        del wb[name]
    else:
        sm = next((i for i, n in enumerate(wb.sheetnames) if n.lower().startswith("summary")), None)
        idx = 0 if sm is None else sm + 1
    ws = wb.create_sheet(name, idx)
    _write_month_sheet(ws, data, articles_all)
    _update_summary(wb, data)
    wb.active = wb.sheetnames.index(name)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def trend_rows_from_xlsx(base_bytes):
    """기존 PR리포트 엑셀의 Summary 값 {형식: [12]} 과 연도. 없으면 (None, None)."""
    if not base_bytes:
        return None, None
    wb = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=True)
    year, rows = summary_rows(wb)
    wb.close()
    return year, rows


# ---------------------------------------------------------------- 작업 파일

def dump_work(meta, articles, releases, meetings, plans, summary, trend_rows) -> bytes:
    d = lambda x: x.isoformat() if x else None
    out = {
        "version": WORK_VERSION, "kind": "monthly",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "meta": meta, "summary": summary, "trend_rows": trend_rows,
        "articles": [{**a, "date": d(a["date"])} for a in articles],
        "releases": [{"title": r["title"], "date": d(r["date"])} for r in releases],
        "meetings": [list(m) for m in meetings], "plans": [list(p) for p in plans],
    }
    return json.dumps(out, ensure_ascii=False, indent=2).encode("utf-8")


def load_work(raw: bytes) -> dict:
    try:
        w = json.loads(raw.decode("utf-8"))
        if w.get("kind") != "monthly":
            raise ValueError
        d = lambda x: date.fromisoformat(x) if x else None
        w["articles"] = [{**a, "date": d(a.get("date"))} for a in w["articles"]]
        w["releases"] = [{"title": r["title"], "date": d(r.get("date"))} for r in w.get("releases", [])]
        w["meetings"] = [tuple(m) for m in w.get("meetings", [])]
        w["plans"] = [tuple(p) for p in w.get("plans", [])]
        w["meta"]["year"], w["meta"]["month"] = int(w["meta"]["year"]), int(w["meta"]["month"])
    except Exception:
        raise ValueError("월간 리포트 작업 파일을 읽을 수 없습니다. 이 탭의 '💾 작업 저장'으로 받은 .json 파일인지 확인하세요.")
    return w
