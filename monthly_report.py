# -*- coding: utf-8 -*-
"""
월간 PR리포트를 앱 안에서 편집해 PDF로 만드는 흐름의 데이터 처리.
  - 수집한 기사에 첫 형식 넣기 / 보도자료 목록 추정
  - 제목이 비슷한 기사 찾기: 우리 보도자료와 같은 소식이면 보도자료로, 단순언급끼리 같은 소식이면 한 묶음으로
  - 앱에서 고친 기사·보도자료·미팅·목표 -> PDF용 데이터(monthly_pdf.build_monthly_pdf 입력) 계산
  - 월별 기록(형식별 1~12월 건수): 지난달 PDF에 숨겨 둔 기록(monthly_pdf.read_history) 또는
    처음 한 번 기존 PR리포트 엑셀의 Summary 시트에서 읽어 오고, 이번 달 PDF에 다시 숨겨 넣는다
  - 작업 파일(.json) 저장 / 불러오기 (수집한 기사 목록까지 통째로 저장해서 다시 검색하지 않음)
"""
import io
import json
import re
from collections import Counter
from datetime import date, datetime

import openpyxl

from media_order import media_sort_key
from monthly_pdf import TYPE_ORDER, assign_release, summary_rows, trend_from_rows

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


# ---------------------------------------------------------------- 제목이 비슷한 기사 찾기

SAME_STORY_DAYS = 4  # 같은 소식으로 볼 게재일 차이


def _bigrams(title, keyword=""):
    """제목의 2글자 조각. [코너명]·키워드·문장부호는 빼고 본다."""
    t = re.sub(r"\[[^\]]*\]", "", title or "")
    if keyword:
        t = t.replace(keyword, "")
    t = re.sub(r"[^0-9a-z가-힣]", "", t.lower())
    return {t[i:i + 2] for i in range(len(t) - 1)}


def _same_story(a, b, ba, bb):
    """같은 보도자료·같은 소식을 옮긴 기사인지: 제목 2글자 조각이 5개 이상, 짧은 쪽의 30% 이상 겹치고 게재일이 가까움.
    실측(딥파인 10월 기사 45건): LX판토스 3건, 관광공사 AI 관광집사 18건이 정확히 묶이고 무관 기사는 섞이지 않음."""
    if a["date"] and b["date"] and abs((a["date"] - b["date"]).days) > SAME_STORY_DAYS:
        return False
    shared = len(ba & bb)
    return shared >= 5 and shared / (min(len(ba), len(bb)) or 1) >= 0.3


def _lead(members):
    """묶음의 대표 기사: 매체 영향력이 가장 큰 기사, 같으면 먼저 실린 기사 (PDF 대표 기사와 같은 기준)."""
    return min(members, key=lambda a: (media_sort_key(a["media"]), a["date"] or date.max))


def auto_mark_releases(articles, releases, keyword, only=None):
    """제목에 키워드가 없어 단순언급·공란으로 들어온 기사 중, 우리 보도자료 기사와 같은 소식이면 보도자료로 바꾼다.
    기준: 형식이 보도자료인 기사나 ① 보도자료 제목과 _same_story. only: 이 URL들만 바꿈 (None 이면 전부).
    반환: 바꾼 기사 수."""
    refs = [(a, _bigrams(a["title"], keyword)) for a in articles if a["kind"] == "보도자료"]
    refs += [({"date": r["date"]}, _bigrams(r["title"], keyword)) for r in releases if r.get("title")]
    changed = 0
    for a in articles:
        if a["kind"] not in ("", "단순언급") or (only is not None and a["url"] not in only):
            continue
        ba = _bigrams(a["title"], keyword)
        if any(_same_story(a, r, ba, br) for r, br in refs):
            a["kind"] = "보도자료"
            a["group"] = ""
            changed += 1
    return changed


def auto_group_mentions(articles):
    """단순언급 기사 중 같은 소식끼리 묶음 이름을 채운다. 이미 묶음 이름이 있는 기사는 그대로 두고,
    비어 있는 기사가 그 기사와 같은 소식이면 같은 이름을 쓴다. 혼자인 기사는 비워 둠 (그래도 1건짜리 주제로 나감).
    묶음 이름은 대표 기사 제목 앞부분. 반환: 묶음이 생긴 기사 수."""
    ms = [a for a in articles if a["kind"] == "단순언급"]
    bg = [_bigrams(a["title"]) for a in ms]
    parent = list(range(len(ms)))

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(len(ms)):
        for j in range(i + 1, len(ms)):
            if _same_story(ms[i], ms[j], bg[i], bg[j]):
                parent[root(i)] = root(j)
    clusters = {}
    for i, a in enumerate(ms):
        clusters.setdefault(root(i), []).append(a)
    filled = 0
    for members in clusters.values():
        blank = [a for a in members if not (a.get("group") or "").strip()]
        if len(members) < 2 or not blank:
            continue
        named = Counter(a["group"].strip() for a in members if (a.get("group") or "").strip())
        name = named.most_common(1)[0][0] if named else re.sub(r"\s+", " ", _lead(members)["title"])[:24].strip()
        for a in blank:
            a["group"] = name
            filled += 1
    return filled


# ---------------------------------------------------------------- 편집 결과 -> PDF 데이터

def compose(meta, articles, releases, meetings, plans, trend_rows):
    """meta: {'company','year','month'} / articles: [{'url','media','title','date','kind','include','group'}]
    releases: [{'title','date','deleted','count'(삭제 건만 직접 넣은 게재 수)}] / meetings, plans: [(a,b,c)] / trend_rows: 월별 기록 {형식: [12]} 또는 None.
    반환: (monthly_pdf 용 data, 보도자료 배포일 목록, 형식 미지정 기사 수)
    이번 달 칸은 표에서 센 값으로 덮어쓰고, 그 기록을 data['history'] 로 PDF에 숨겨 넣게 한다.
    '게재 후 삭제'한 보도자료의 기사는 건수·기사 목록·월별 기록에서 모두 뺀다 (카드만 회색으로 남김)."""
    dates = [r["date"] for r in releases]
    gone = {i for i, r in enumerate(releases) if r.get("deleted")}
    used = [a for a in articles if a["include"] and a["kind"] and assign_release(a, dates) not in gone]
    blank = sum(1 for a in articles if a["include"] and not a["kind"])

    types = {k: 0 for k in TYPE_ORDER}
    for a in used:
        types[a["kind"]] = types.get(a["kind"], 0) + 1

    rel_counts = Counter(assign_release(a, dates) for a in used)
    rels = [{"title": r["title"], "deleted": bool(r.get("deleted")),
             "count": int(r.get("count") or 0) if r.get("deleted") else rel_counts.get(i, 0)}
            for i, r in enumerate(releases)]

    main = [{"media": a["media"], "title": a["title"], "url": a["url"], "date": a["date"], "kind": a["kind"]}
            for a in sort_articles(a for a in used if a["kind"] != "단순언급")]

    groups = {}
    for a in used:
        if a["kind"] == "단순언급":
            groups.setdefault((a.get("group") or "").strip() or a["title"], []).append(a)
    mentions = []
    for members in groups.values():
        lead = _lead(members)
        mentions.append({"media": lead["media"], "title": lead["title"], "url": lead["url"],
                         "count": len(members), "first": min(a["date"] or date.max for a in members)})
    mentions.sort(key=lambda m: (-m["count"], m["first"]))

    m = meta["month"]
    rows = {k: list((trend_rows or {}).get(k) or [0] * 12) for k in TYPE_ORDER}
    for k in TYPE_ORDER:
        rows[k][m - 1] = types[k]
    has_past = any(rows[k][i] for k in TYPE_ORDER for i in range(m - 1))
    trend = trend_from_rows(rows) if has_past else None  # 지난 기록이 없으면 추이·전월 비교는 빼고 그림

    data = {"company": meta["company"], "year": meta["year"], "month": m, "types": types,
            "releases": rels, "articles": main, "mentions": mentions,
            "meetings": [x for x in meetings if any(x)], "plans": [p for p in plans if any(p)],
            "plan_title": f"{m % 12 + 1}월 PR 목표", "trend": trend,
            "history": {"company": meta["company"], "year": meta["year"], "month": m, "rows": rows,
                        "detail": month_detail(main, rels, dates, [x for x in meetings if any(x)])}}
    return data, dates, blank


def month_detail(main, rels, dates, meetings):
    """연간 리포트가 월간 PDF에서 다시 읽는 이번 달 원본 (PDF에 월별 기록과 함께 숨겨 넣음).
    주요 기사(단순언급 제외)·보도자료(제목·배포일·게재 수, 게재 후 삭제 건은 뺌)·기자 미팅."""
    d = lambda x: x.isoformat() if x else None
    return {
        "articles": [{"media": a["media"], "title": a["title"], "url": a["url"], "date": d(a["date"]), "kind": a["kind"]}
                     for a in main],
        "releases": [{"title": r["title"], "date": d(dt), "count": r["count"]}
                     for r, dt in zip(rels, dates) if not r["deleted"]],
        "meetings": [list(x) for x in meetings],
    }


# ---------------------------------------------------------------- 월별 기록 불러오기

def history_entry(secrets, company, year):
    """Streamlit Secrets 의 [monthly_history."고객사"."연도"] 표 (저장소가 공개라 코드 대신 Secrets에 둠).
    고객사 이름은 그대로 또는 서로 포함되면 맞는 것으로 본다. 없으면 None.
    형식별 1~12월 건수 외에 연간 리포트용 '보도자료 배포'(그해 보도자료 수)가 들어 있을 수 있다."""
    try:
        hist = secrets["monthly_history"]
        name = next((k for k in hist if k == company), None) or next(
            (k for k in hist if k in company or company in k), None)
        if name is None or str(year) not in hist[name]:
            return None
        return hist[name][str(year)]
    except Exception:
        return None


def history_from_secrets(secrets, company, year):
    """Secrets 에 넣어 둔 형식별 1~12월 건수 {형식: [12]}. 없으면 None."""
    src = history_entry(secrets, company, year)
    if src is None:
        return None
    try:
        return {k: [int(x) for x in (list(src.get(k, [])) + [0] * 12)[:12]] for k in TYPE_ORDER}
    except Exception:
        return None


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
        "releases": [{"title": r["title"], "date": d(r["date"]), "deleted": r.get("deleted", False),
                      "count": int(r.get("count") or 0)} for r in releases],
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
        w["releases"] = [{"title": r["title"], "date": d(r.get("date")), "deleted": r.get("deleted", False),
                          "count": int(r.get("count") or 0)}
                         for r in w.get("releases", [])]
        w["meetings"] = [tuple(m) for m in w.get("meetings", [])]
        w["plans"] = [tuple(p) for p in w.get("plans", [])]
        w["meta"]["year"], w["meta"]["month"] = int(w["meta"]["year"]), int(w["meta"]["month"])
    except Exception:
        raise ValueError("월간 리포트 작업 파일을 읽을 수 없습니다. 이 탭의 '💾 작업 저장'으로 받은 .json 파일인지 확인하세요.")
    return w
