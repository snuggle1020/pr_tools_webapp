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

from media_order import MEDIA_GROUPS, media_category, media_sort_key
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

def compose(meta, articles, releases, meetings, plans, trend_rows, competitors=None):
    """meta: {'company','year','month'} / articles: [{'url','media','title','date','kind','include','group'}]
    releases: [{'title','date','deleted','count'(삭제 건만 직접 넣은 게재 수)}] / meetings, plans: [(a,b,c)] / trend_rows: 월별 기록 {형식: [12]} 또는 None.
    반환: (monthly_pdf 용 data, 보도자료 배포일 목록, 형식 미지정 기사 수)
    이번 달 칸은 표에서 센 값으로 덮어쓰고, 그 기록을 data['history'] 로 PDF에 숨겨 넣게 한다.
    '게재 후 삭제'한 보도자료의 기사는 건수·기사 목록·월별 기록에서 모두 뺀다 (카드만 회색으로 남김).
    competitors: [{'name', 'category'(분류, 예: DAO), 'articles': [{'url','media','title','date','kind'(COMP_KINDS)}]}]
    - 있으면 PDF 09 경쟁사 비교."""
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
    trend = trend_from_rows(rows)  # 지난 달이 전부 0이어도 월별 추이 그래프는 그림 (전월 비교·누적은 has_past 일 때만)

    data = {"company": meta["company"], "year": meta["year"], "month": m, "types": types,
            "releases": rels, "articles": main, "mentions": mentions,
            "meetings": [x for x in meetings if any(x)], "plans": [p for p in plans if any(p)],
            "plan_title": f"{m % 12 + 1}월 PR 목표", "trend": trend, "has_past": has_past,
            "media_types": _media_types(used),
            "portals": [(k, sum(1 for a in used if article_portal(a) == k)) for k in PORTAL_OPTIONS],
            "competitors": _competitor_rows(meta["company"], used, rels, competitors) if competitors else None,
            "history": {"company": meta["company"], "year": meta["year"], "month": m, "rows": rows,
                        "detail": month_detail(main, rels, dates, [x for x in meetings if any(x)])}}
    return data, dates, blank


# ---------------------------------------------------------------- 경쟁사 비교 (PDF 09)

COMP_KINDS = ["보도자료", "기획·인터뷰", "주가·시황", "단순언급"]
# 포털별 게재 수 (PDF 05). 네이버에 실린 기사는 다음·구글에도 있어도 네이버로만 셈
PORTAL_OPTIONS = ["네이버 CP", "네이버 검색제휴", "다음", "구글"]


def naver_portal(link):
    """네이버 검색 결과의 link 가 n.news.naver.com 이면 네이버 CP(네이버 뉴스 제휴), 아니면 검색제휴."""
    return "네이버 CP" if "n.news.naver.com" in (link or "") else "네이버 검색제휴"


def article_portal(a):
    """기사 포털: 저장된 값, 없으면 네이버로 모은 기사는 검색제휴·직접 넣은 기사는 구글로 봄 (예전 작업 파일)."""
    return a.get("portal") or ("네이버 검색제휴" if a.get("naver", True) else "구글")
# 종목 나열·시세 기사 제목에 자주 나오는 말 (경쟁사 기사 자동 분류용)
STOCK_WORDS = re.compile(r"주[가식]|테마주|관련주|[가-힣]+주에|[가-힣]+주|株|매수|순매도|상승|하락|급등|급락|강세|약세|시황|"
                         r"코스피|코스닥|랠리|희비|트래커|DAT|상폐|청산|비트코인|가상자산|가상화폐|코인|실적")
COMP_RELEASE_DAYS = 2  # 경쟁사 보도자료: 첫 기사부터 며칠 안의 같은 소식까지 한 건으로


def _story_groups(articles, keyword):
    """먼저 실린 기사부터, COMP_RELEASE_DAYS 안의 같은 소식(_same_story)끼리 묶음."""
    groups = []
    for a in sorted(articles, key=lambda a: a["date"] or date.max):
        ba = _bigrams(a["title"], keyword)
        g = next((g for g in groups if a["date"] and g[0][0]["date"]
                  and (a["date"] - g[0][0]["date"]).days <= COMP_RELEASE_DAYS
                  and any(_same_story(a, x, ba, bx) for x, bx in g)), None)
        (g.append((a, ba)) if g else groups.append([(a, ba)]))
    return [[a for a, _ in g] for g in groups]


def classify_competitor(name, articles):
    """경쟁사 기사 형식 자동 분류 (articles 의 'kind' 를 채움). 10월 버넥트·비트맥스 63건 중 62건이 사람 판단과 같았음.
    보도자료: 같은 소식 2건 이상 + 제목 절반 이상이 기업명으로 시작 / 주가·시황: STOCK_WORDS /
    기획·인터뷰: 제목에 기업명 / 단순언급: 나머지."""
    rel = set()
    for g in _story_groups(articles, name):
        if len(g) >= 2 and sum(a["title"].startswith(name) for a in g) * 2 >= len(g):
            rel |= {id(a) for a in g}
    for a in articles:
        a["kind"] = ("보도자료" if id(a) in rel else "주가·시황" if STOCK_WORDS.search(a["title"])
                     else "기획·인터뷰" if name in a["title"] else "단순언급")
    return articles


def competitor_summary(name, articles):
    """담당자가 확인한 형식으로 센 경쟁사 한 줄 + 보도자료 목록(같은 소식끼리 한 건, 대표 기사는 기업명으로 시작하는 첫 기사)."""
    rel_arts = [a for a in articles if a["kind"] == "보도자료"]
    releases = []
    for g in _story_groups(rel_arts, name):
        lead = next((a for a in g if a["title"].startswith(name)), _lead(g))
        releases.append({"title": lead["title"], "url": lead["url"], "date": g[0]["date"], "count": len(g)})
    n = Counter(a["kind"] for a in articles)
    return {"name": name, "n_rel": len(releases), "releases": releases, "rel_articles": len(rel_arts),
            "plan": n.get("기획·인터뷰", 0), "stock": n.get("주가·시황", 0), "mention": n.get("단순언급", 0),
            "total": len(articles)}


def _competitor_rows(company, used, rels, competitors):
    """09 표: 우리 회사(경쟁사와 같은 네이버 기준 - ②-1 로 직접 넣은 기사는 뺌) + 경쟁사들."""
    nav = Counter(a["kind"] for a in used if a.get("naver", True))
    own = {"name": company, "own": True, "n_rel": sum(1 for r in rels if not r["deleted"]), "releases": [],
           "rel_articles": nav.get("보도자료", 0),
           "plan": sum(nav.get(k, 0) for k in ("기획(오픈피알)", "기획(외부)", "인터뷰", "기고")),
           "stock": 0, "mention": nav.get("단순언급", 0)}
    own["total"] = own["rel_articles"] + own["plan"] + own["mention"]
    return [own] + [{**competitor_summary(c["name"], c["articles"]), "category": (c.get("category") or "").strip()}
                    for c in competitors]


def _media_types(used):
    """매체 유형별 건수 [(유형, 건수)] - 보고서 매체 순서(통신사·일간지·경제지…)대로, 0건 유형은 뺌."""
    counts = Counter(media_category(a["media"]) for a in used)
    order = list(dict.fromkeys(g[1] for g in MEDIA_GROUPS)) + ["확인필요"]
    return [(c, counts[c]) for c in order if counts.get(c)] + [(c, n) for c, n in counts.items() if c not in order]


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

def dump_work(meta, articles, releases, meetings, plans, summary, trend_rows, competitors=None) -> bytes:
    d = lambda x: x.isoformat() if x else None
    out = {
        "version": WORK_VERSION, "kind": "monthly",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "meta": meta, "summary": summary, "trend_rows": trend_rows,
        "articles": [{**a, "date": d(a["date"])} for a in articles],
        "releases": [{"title": r["title"], "date": d(r["date"]), "deleted": r.get("deleted", False),
                      "count": int(r.get("count") or 0)} for r in releases],
        "meetings": [list(m) for m in meetings], "plans": [list(p) for p in plans],
        "competitors": [{"name": c["name"], "category": c.get("category", ""),
                         "articles": [{**a, "date": d(a["date"])} for a in c["articles"]]}
                        for c in (competitors or [])],
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
        w["competitors"] = [{"name": c["name"], "category": c.get("category", ""),
                             "articles": [{**a, "date": d(a.get("date"))} for a in c["articles"]]}
                            for c in w.get("competitors", [])]
        w["meta"]["year"], w["meta"]["month"] = int(w["meta"]["year"]), int(w["meta"]["month"])
    except Exception:
        raise ValueError("월간 리포트 작업 파일을 읽을 수 없습니다. 이 탭의 '💾 작업 저장'으로 받은 .json 파일인지 확인하세요.")
    return w
