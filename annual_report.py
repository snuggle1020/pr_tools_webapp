# -*- coding: utf-8 -*-
"""
연간 PR리포트 데이터 처리.
  - 이 앱이 만든 월간 PDF(숨겨 둔 그달 기사·보도자료·미팅)와 월간 PR리포트 엑셀(앱 이전 달, 지난 해 12월 파일)에서 달마다 형식별 건수·보도자료·기사·기자 미팅을 읽음
  - 보도자료는 달을 넘겨 이어 실린 것까지 제목으로 합치고, 배포일은 그 달 보도자료 기사가 몰린 날짜로 추정 (앱에서 고침)
  - 매체·포털·기자 통계는 '주요 기사' 목록(보도자료·기획·인터뷰·기고) 기준. 단순언급은 엑셀에 대표 기사만 있어 형식별 건수에만 들어감
  - 포털(네이버 CP·검색제휴·다음·구글)과 기자명은 엑셀에 없어서 앱에서 버튼으로 확인 (check_portal / fetch_reporter)
  - compose_annual(): 앱에서 고친 값 -> annual_pdf.build_annual_pdf 입력
"""
import io
import re
import urllib.request
from collections import Counter
from datetime import date

import openpyxl
import requests

from article_meta import _UA, normalize_url
from media_order import MEDIA_GROUPS, media_category, media_sort_key
from monthly_pdf import assign_release, read_history, read_month_report
from monthly_report import _bigrams, history_entry, history_from_secrets
from pr_report_core import _decode_html, clean_html

# PDF 형식 5가지: (키, 이름, 엑셀 형식 이름들)
TYPES5 = [("release", "보도자료", ["보도자료"]), ("plan", "기획", ["기획(오픈피알)", "기획(외부)"]),
          ("interview", "인터뷰", ["인터뷰"]), ("contrib", "기고", ["기고"]), ("mention", "단순언급", ["단순언급"])]
TOPICS = ["제품·서비스", "제휴·MOU", "투자·실적", "공공사업", "수상·선정", "행사·전시", "기타"]
KEY_GROUPS = ("통신사", "일간지", "주요 경제지")  # '주요 매체' (media_order 정렬 그룹 이름)
PORTALS = ["네이버 CP", "네이버 검색제휴", "다음·구글"]
RELEASE_GAP_DAYS = 3  # 보도자료 기사 날짜가 이만큼 넘게 벌어지면 다른 보도자료로 봄

_OFFICIAL = {re.sub(r"\s", "", n): n for _, _, names in MEDIA_GROUPS for n in names}


def canon_media(name):
    """엑셀 매체명 표기 차이(띄어쓰기 등)를 media_order 공식 이름으로. 모르는 매체는 그대로."""
    name = (name or "").strip()
    return _OFFICIAL.get(re.sub(r"\s", "", name), name)


def media_group(name):
    """media_order 정렬 그룹 이름 (통신사·일간지·주요 경제지 …)."""
    gi = media_sort_key(canon_media(name))[0]
    return MEDIA_GROUPS[gi][0] if gi < len(MEDIA_GROUPS) else ""


def kind5(kind):
    """엑셀 기사형식(보도자료·기획·인터뷰·기고·…) -> TYPES5 키."""
    k = (kind or "").strip()
    if k.startswith("기획"):
        return "plan"
    return {"보도자료": "release", "인터뷰": "interview", "기고": "contrib", "단순언급": "mention"}.get(k)


def _norm_title(t):
    return re.sub(r"[\s,·.]", "", t or "")


# ---------------------------------------------------------------- 엑셀 읽기

def read_year_excel(raw: bytes):
    """월간 PR리포트 엑셀 하나 -> {'company','year','months': {월: read_month_report 결과}}.
    한 파일에 1월~그 달 시트가 모두 들어 있다 (지난 해는 12월 파일)."""
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    sheets = [s for s in wb.sheetnames if re.fullmatch(r"\d{1,2}월", s.strip())]
    if not sheets:
        raise ValueError("월 시트(1월, 2월 …)가 없는 엑셀입니다. 월간 PR리포트 엑셀인지 확인하세요.")
    months = {}
    for s in sheets:
        d = read_month_report(None, s, wb=wb)
        months[d["month"]] = d
    wb.close()
    first = months[min(months)]
    # 연도: Summary 시트 제목 > 기사 날짜 (read_month_report 가 이미 고름)
    year = first["year"]
    years = Counter(a["date"].year for d in months.values() for a in d["articles"] if a["date"])
    if years and year not in years:
        year = years.most_common(1)[0][0]
    return {"company": first["company"], "year": int(year), "months": months}


def month_types(months):
    """{TYPES5 키: [12]} - 월 시트의 '기사형식별 보도 수' 표."""
    out = {k: [0] * 12 for k, _, _ in TYPES5}
    for m, d in months.items():
        for k, _, names in TYPES5:
            out[k][m - 1] = sum(d["types"].get(n, 0) for n in names)
    return out


def year_releases(months):
    """달마다 보도자료 표를 제목으로 합친 목록 (처음 나온 순). [{'title','n','months':[월]}]"""
    out, index = [], {}
    for m in sorted(months):
        for r in months[m]["releases"]:
            key = _norm_title(r["title"])
            if not key:
                continue
            if key not in index:
                index[key] = len(out)
                out.append({"title": r["title"], "n": 0, "months": [], "date": None})
            rec = out[index[key]]
            rec["n"] += r["count"]
            rec["months"].append(m)
            if r.get("date") and not rec["date"]:
                rec["date"] = r["date"]
    return out


def _clusters(arts):
    """보도자료 기사를 날짜로 묶음 (RELEASE_GAP_DAYS 넘게 떨어지면 새 묶음). [[기사…]]"""
    arts = sorted((a for a in arts if a["date"]), key=lambda a: a["date"])
    out = []
    for a in arts:
        if out and (a["date"] - out[-1][-1]["date"]).days <= RELEASE_GAP_DAYS:
            out[-1].append(a)
        else:
            out.append([a])
    return out


def guess_release_dates(months, releases):
    """보도자료마다 배포일 추정: 처음 나온 달의 보도자료 기사 묶음 중 제목이 가장 비슷한 묶음의 첫 날짜.
    (엑셀 보도자료 표에는 날짜가 없음) 못 찾으면 None. releases 에 'date' 를 채운다."""
    by_month = {}
    for r in releases:
        if not r.get("date"):  # 월간 PDF에서 온 보도자료는 배포일이 이미 있음
            by_month.setdefault(r["months"][0], []).append(r)
    for m, rels in by_month.items():
        cl = _clusters([a for a in months[m]["articles"] if kind5(a["kind"]) == "release"])
        used = set()
        for r in rels:
            br = _bigrams(r["title"])
            best, best_score = None, -1
            for i, c in enumerate(cl):
                if i in used:
                    continue
                score = max(len(br & _bigrams(a["title"])) for a in c)
                if score > best_score:
                    best, best_score = i, score
            if best is not None:
                used.add(best)
                r["date"] = cl[best][0]["date"]
            else:
                r["date"] = None
    return releases


def guess_topic(title):
    t = title or ""
    rules = [("투자·실적", r"투자|유치|실적|매출|상장|법인 설립|IPO"),
             ("제휴·MOU", r"MOU|협약|협력|제휴|손잡|공동|파트너"),
             ("수상·선정", r"수상|선정|인증|대상|장관상"),
             ("행사·전시", r"참가|전시|박람회|대전|쇼|MWC|CES|컨퍼런스|기조|세미나|포럼"),
             ("공공사업", r"국책|과기부|과기정통부|정부|공공|바우처|수주|주관기관")]
    for name, pat in rules:
        if re.search(pat, t, re.I):
            return name
    return "제품·서비스"


def read_month_pdf(raw: bytes):
    """이 앱이 만든 월간 PDF에 숨겨 둔 기록 -> (회사, 연도, 월, 이번 달 데이터|None, 형식별 1~12월 건수).
    이번 달 데이터(기사·보도자료·미팅)는 2026년 10월 리포트부터 들어 있음. 기록이 없는 PDF면 ValueError."""
    h = read_history(raw)
    if not h:
        raise ValueError("이 앱의 월간 PR리포트 탭에서 만든 PDF가 아닙니다 (숨겨 둔 기록이 없음).")
    det = h.get("detail")
    month = None
    if det:
        d = lambda x: date.fromisoformat(x) if x else None
        m = h["month"]
        month = {
            "month": m, "company": h["company"], "year": h["year"],
            "types": {k: h["rows"][k][m - 1] for k in h["rows"]},
            "articles": [dict(a, date=d(a.get("date"))) for a in det.get("articles", [])],
            "releases": [{"title": r["title"], "count": int(r.get("count") or 0), "date": d(r.get("date"))}
                         for r in det.get("releases", [])],
            "meetings": [tuple((list(x) + ["", "", ""])[:3]) for x in det.get("meetings", [])],
        }
    return h["company"], h["year"], h["month"], month, h["rows"]


def load_annual(cur_raw=None, prev_raws=(), pdf_raws=()):
    """앱에 올린 파일들 -> 앱 편집 전 기본 데이터.
    cur_raw: 올해 PR리포트 엑셀(앱 전에 엑셀로 만든 달) / pdf_raws: 올해 월간 PDF들 (같은 달은 PDF가 우선)
    prev_raws: 지난 해 12월 엑셀들. 반환에 'notes' (사용자에게 보여 줄 안내) 포함."""
    months, company, year, notes = {}, None, None, []
    if cur_raw:
        cur = read_year_excel(cur_raw)
        months.update(cur["months"])
        company, year = cur["company"], cur["year"]
    rows_all = None  # 가장 늦은 달 PDF의 형식별 1~12월 건수 (기사 목록이 없는 달 건수 채우기용)
    latest = 0
    for raw in pdf_raws:
        c, y, m, month, rows = read_month_pdf(raw)
        if year is None:
            company, year = c, y
        if y != year:
            notes.append(f"{y}년 {m}월 PDF는 {year}년이 아니라 뺐습니다.")
            continue
        if m >= latest:
            latest, rows_all = m, rows
        if month:
            months[m] = month
        else:
            notes.append(f"{m}월 PDF에는 기사 목록이 없어(이전 버전으로 만든 PDF) 건수만 씁니다.")
    if year is None:
        raise ValueError("올해 월간 PDF나 PR리포트 엑셀을 올려 주세요.")
    prev = []
    for raw in prev_raws:
        y = read_year_excel(raw)
        if y["year"] != year:
            prev.append(y)
    prev.sort(key=lambda y: y["year"])
    monthly = month_types(months)
    if rows_all:  # 기사 목록은 없지만 PDF 기록에 건수가 있는 달
        for i in range(latest):
            if i + 1 not in months:
                for k, _, names in TYPES5:
                    monthly[k][i] = sum(rows_all.get(n, [0] * 12)[i] for n in names)
    last = max([latest] + list(months))
    missing = [m for m in range(1, last + 1) if m not in months]
    if missing:
        notes.append("기사 목록이 없는 달: " + ", ".join(f"{m}월" for m in missing)
                     + " (월별 건수에는 들어가지만 매체·보도자료·하이라이트 집계에서는 빠집니다)")
    releases = guess_release_dates(months, year_releases(months))
    for r in releases:
        r["topic"] = guess_topic(r["title"])
    articles = []
    for m in sorted(months):
        for a in months[m]["articles"]:
            k = kind5(a["kind"])
            if not k or k == "mention":
                continue
            articles.append({"media": canon_media(a["media"]), "title": a["title"], "url": a["url"],
                             "date": a["date"], "kind": k, "portal": None, "reporter": None})
    meetings = [{"month": m, "media": canon_media(x[0]), "reporter": x[1].strip(), "memo": x[2].strip()}
                for m in sorted(months) for x in months[m]["meetings"] if (x[0] or "").strip()]
    return {
        "company": company, "year": year, "last_month": last,
        "monthly": monthly, "releases": releases, "articles": articles, "meetings": meetings, "notes": notes,
        "prev_years": [{"year": y["year"], "monthly": month_types(y["months"]),
                        "releases": len(year_releases(y["months"]))} for y in prev],
    }


# ---------------------------------------------------------------- 포털 · 기자명 확인 (엑셀에 없음)

def check_portal(article, client_id, client_secret):
    """네이버 뉴스 검색에서 이 기사를 찾아 'n.news.naver.com' 링크가 있으면 네이버 CP, 있되 없으면 검색제휴, 못 찾으면 다음·구글."""
    title = re.sub(r"(\.\.\.|…)$", "", article["title"]).strip()
    query = re.sub(r"[\"'“”‘’\[\]]", " ", title)[:40]
    try:
        r = requests.get("https://openapi.naver.com/v1/search/news.json",
                         headers={"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret},
                         params={"query": query, "display": 30, "sort": "sim"}, timeout=10)
        r.raise_for_status()
        items = r.json().get("items", [])
    except Exception:
        return None
    target = normalize_url(article["url"])
    want = _bigrams(title)
    for it in items:
        orig = normalize_url(it.get("originallink") or "")
        same = orig == target or (orig and target and orig.split("/")[0] == target.split("/")[0]
                                  and len(want & _bigrams(clean_html(it.get("title", "")))) >= max(5, len(want) // 2))
        if same:
            return "네이버 CP" if "n.news.naver.com" in (it.get("link") or "") else "네이버 검색제휴"
    return "다음·구글"


_BYLINE_META = ["author", "article:author", "dable:author", "og:article:author", "byl"]
_NOT_NAME = {"기자", "편집", "온라인", "디지털", "뉴스", "취재", "사진", "영상", "인턴", "객원", "선임", "수석", "전문", "대표", "특파원"}


def _byline_names(text):
    """'OOO 기자' · 'OOO 특파원' 의 이름 (앞뒤가 다른 한글에 붙어 있으면 지명·단어라 뺌)."""
    return [n for n in re.findall(r"(?<![가-힣])([가-힣]{2,4})\s?(?:기자|특파원)(?![가-힣])", text) if n not in _NOT_NAME]


def fetch_reporter(url):
    """기사 페이지에서 기자 이름 (못 찾으면 ''). 메타 태그(이름만 있거나 'OOO 기자') > 본문에 가장 많이 나온 'OOO 기자'."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": _UA})
        with urllib.request.urlopen(req, timeout=8) as resp:
            page = _decode_html(resp.read(800_000))
    except Exception:
        return ""
    for key in _BYLINE_META:
        k = re.escape(key)
        m = re.search(rf'<meta[^>]+(?:property|name)\s*=\s*["\']{k}["\'][^>]*?content\s*=\s*["\']([^"\']+)', page, re.I)
        if not m:
            continue
        val = m.group(1).strip()
        names = _byline_names(val)
        if names:
            return names[0]
        if re.fullmatch(r"[가-힣]{2,4}", val) and val not in _NOT_NAME:
            return val
    text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", page, flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    cands = Counter(_byline_names(text))
    return cands.most_common(1)[0][0] if cands else ""


# ---------------------------------------------------------------- PDF 데이터

def compose_annual(base, company, releases, highlights, summary_text, wrap):
    """base: load_annual() 결과(포털·기자 확인값 포함) / releases: 앱에서 고친 [{'title','date','topic','n'}]
    highlights: 하이라이트로 고른 기사들 / wrap: {'good','bad','plan': [줄]}. 반환: annual_pdf 입력."""
    year, last = base["year"], base["last_month"]
    monthly = base["monthly"]
    types = {k: sum(monthly[k]) for k, _, _ in TYPES5}
    total = sum(types.values())
    prev = next((p for p in base["prev_years"] if p["year"] == year - 1), None)
    # 올해가 아직 12월까지 없으면 전년도 같은 기간(1월~last월)끼리 비교
    prev_types = {k: sum(prev["monthly"][k][:last]) for k, _, _ in TYPES5} if prev else None
    prev_monthly = [sum(prev["monthly"][k][i] for k, _, _ in TYPES5) for i in range(12)] if prev else None

    # 보도자료: 주요 매체 수(통신사·일간지·주요 경제지 중 실린 곳)는 배포일로 기사를 묶어 셈
    rels = sorted([dict(r) for r in releases if r["title"]], key=lambda r: (r["date"] or date.max))
    dates = [r["date"] for r in rels]
    keys = [set() for _ in rels]
    arts = base["articles"]
    for a in arts:
        if a["kind"] != "release":
            continue
        i = assign_release({"kind": "보도자료", "date": a["date"]}, [d or date.max for d in dates])
        if i is not None and media_group(a["media"]) in KEY_GROUPS:
            keys[i].add(a["media"])
    for r, k in zip(rels, keys):
        r["key"] = len(k)
    rel_avg = (sum(r["n"] for r in rels) / len(rels)) if rels else 0

    topics = {}
    for r in rels:
        topics.setdefault(r["topic"] or "기타", []).append(r["n"])
    topic_rows = sorted(((t, sum(v) / len(v), len(v)) for t, v in topics.items()), key=lambda x: -x[1])
    wd = {w: [r["n"] for r in rels if r["date"] and r["date"].weekday() == w] for w in range(5)}
    weekday_rows = [("월화수목금"[w] + "요일", (sum(v) / len(v)) if v else 0, len(v)) for w, v in wd.items()]

    cat = Counter(media_category(a["media"]) for a in arts)
    from media_order import CATEGORY_ORDER
    media_types = [(c, cat.get(c, 0)) for c in CATEGORY_ORDER if cat.get(c, 0)]
    unknown = sum(n for c, n in cat.items() if c not in CATEGORY_ORDER)
    if unknown:
        media_types.append(("기타", unknown))
    key_share = sum(1 for a in arts if media_group(a["media"]) in KEY_GROUPS)

    checked = [a for a in arts if a["portal"]]
    portals = [(p, sum(1 for a in checked if a["portal"] == p)) for p in PORTALS] if checked else None
    top_media = [(m, media_category(m), n) for m, n in Counter(a["media"] for a in arts).most_common(5)]
    named = [a for a in arts if a["reporter"]]
    top_reporters = None
    if named:
        cnt = Counter((a["reporter"], a["media"]) for a in named)
        top_reporters = [(r, m, n) for (r, m), n in cnt.most_common(5)]

    meet = None
    if base["meetings"]:
        def at(media):
            return [a for a in arts if a["media"] == media or (len(media) >= 2 and (a["media"].startswith(media) or media.startswith(a["media"])))]
        conv = conv_planned = 0
        first_month = {}
        for mt in base["meetings"]:
            first_month[mt["media"]] = min(first_month.get(mt["media"], 13), mt["month"])
            later = [a for a in at(mt["media"]) if a["date"] and a["date"] >= date(year, mt["month"], 1)]
            conv += bool(later)
            conv_planned += any(a["kind"] in ("plan", "interview", "contrib") for a in later)
        got = {}
        for media, m in first_month.items():
            for a in at(media):
                if a["date"] and a["date"] >= date(year, m, 1):
                    got[a["url"] or (a["media"], a["title"])] = a
        meet = {"meetings": len(base["meetings"]), "outlets": len(first_month), "converted": conv,
                "conv_planned": conv_planned,
                "stories": sum(1 for a in got.values() if a["kind"] != "release"),
                "releases": sum(1 for a in got.values() if a["kind"] == "release")}

    # 연도별: (연도, 총 보도, 보도자료 수, (첫 달, 끝 달), 전년 대비 %) - 올해가 12월 전이면 전년 같은 기간과 비교
    yearly = []
    for p in base["prev_years"]:
        tot_m = [sum(p["monthly"][k][i] for k, _, _ in TYPES5) for i in range(12)]
        on = [i + 1 for i, v in enumerate(tot_m) if v] or [1, 12]
        yearly.append([p["year"], sum(tot_m), p["releases"], (on[0], on[-1]), None])
    yearly.append([year, total, len(rels), (1, last), None])
    for i in range(1, len(yearly)):
        before = yearly[i - 1]
        if yearly[i][0] == year and last < 12 and prev_types:
            yearly[i][4] = pct(total, sum(prev_types.values()))
        elif before[0] == yearly[i][0] - 1:
            yearly[i][4] = pct(yearly[i][1], before[1])

    return {
        "company": company, "year": year, "last_month": last,
        "monthly": monthly, "monthly_total": [sum(monthly[k][i] for k, _, _ in TYPES5) for i in range(12)],
        "prev_monthly": prev_monthly, "types": types, "prev_types": prev_types, "total": total,
        "prev_total": sum(prev_types.values()) if prev_types else None,
        "releases": rels, "rel_avg": rel_avg, "topics": topic_rows, "weekdays": weekday_rows,
        "yearly": yearly[-6:], "highlights": sorted(highlights, key=lambda a: a["date"] or date.max)[:6],
        "media_types": media_types, "key_media": key_share, "media_total": len(arts),
        "portals": portals, "top_media": top_media, "top_reporters": top_reporters, "meet": meet,
        "summary": summary_text, "wrap": wrap,
    }


def pct(a, b):
    return round((a / b - 1) * 100) if b else None


def prev_label(d):
    """올해가 12월까지 없으면 '전년 동기'(1월~마지막 달끼리 비교), 아니면 '전년'."""
    return "전년" if d["last_month"] == 12 else "전년 동기"


def default_summary(d):
    """1쪽 '올해의 요약' 초안 (보고서 말투)."""
    s = [f"보도자료 {len(d['releases'])}건을 배포해 총 {d['total']}건이 보도됐습니다."]
    if d["prev_total"]:
        p = pct(d["total"], d["prev_total"])
        s.append(f"{prev_label(d)}({d['prev_total']}건) 대비 {abs(p)}% {'증가' if p >= 0 else '감소'}했습니다.")
    if d["releases"]:
        top = max(d["releases"], key=lambda r: r["n"])
        s.append(f"가장 많이 실린 보도자료는 '{top['title']}'({top['n']}건)입니다.")
    planned = d["types"]["interview"] + d["types"]["plan"] + d["types"]["contrib"]
    if planned:
        s.append(f"인터뷰 {d['types']['interview']}건·기획 {d['types']['plan']}건·기고 {d['types']['contrib']}건이 실렸습니다.")
    return " ".join(s)


def default_wrap(d):
    """총 정리 초안: 수치로 쓸 수 있는 줄만 (나머지는 담당자가 씀)."""
    good, bad, plan = [], [], []
    if d["prev_total"]:
        p = pct(d["total"], d["prev_total"])
        (good if p >= 0 else bad).append(f"총 보도 {d['total']}건, {prev_label(d)} 대비 {'+' if p >= 0 else ''}{p}%")
    if d["releases"]:
        top = max(d["releases"], key=lambda r: r["n"])
        good.append(f"'{top['title']}' 보도자료가 {top['n']}건으로 가장 많이 실림")
    if d["topics"]:
        good.append(f"주제별로는 {d['topics'][0][0]} 소식의 반응이 가장 큼 (평균 {d['topics'][0][1]:.1f}건)")
    months = d["monthly_total"][:d["last_month"]]
    if months:
        low = sorted(range(len(months)), key=lambda i: months[i])[:2]
        bad.append("보도가 적었던 달: " + ", ".join(f"{i + 1}월({months[i]}건)" for i in sorted(low)))
    days = [w for w in d["weekdays"] if w[2] >= 2]
    if days:
        best = max(days, key=lambda w: w[1])
        plan.append(f"{best[0]} 배포를 기본으로 (평균 {best[1]:.1f}건으로 가장 많이 실림)")
    return {"good": good, "bad": bad, "plan": plan}


# ---------------------------------------------------------------- 하이라이트 후보

def highlight_candidates(articles, releases):
    """하이라이트로 고를 수 있는 기사: 인터뷰·기획·기고 전부 + 보도자료마다 대표 기사 1건(매체 영향력이 가장 큰 기사). 게재일 순."""
    rels = sorted([r for r in releases if r.get("title")], key=lambda r: (r["date"] or date.max))
    dates = [r["date"] or date.max for r in rels]
    leads = {}
    for a in articles:
        if a["kind"] != "release":
            continue
        i = assign_release({"kind": "보도자료", "date": a["date"]}, dates)
        if i is not None and (i not in leads or media_sort_key(a["media"]) < media_sort_key(leads[i]["media"])):
            leads[i] = a
    out = [a for a in articles if a["kind"] in ("plan", "interview", "contrib")]
    out += [dict(a, release=rels[i]["title"]) for i, a in leads.items()]
    return sorted(out, key=lambda a: (a["date"] or date.max, media_sort_key(a["media"])))


def default_highlights(cands, releases, n=6):
    """기본 하이라이트: 가장 많이 실린 보도자료의 대표 기사 + 인터뷰 > 기고 > 기획 (매체 영향력 순)으로 n건."""
    picked = []
    if releases:
        top = max(releases, key=lambda r: r["n"])["title"]
        picked += [a for a in cands if a.get("release") == top][:1]
    rank = {"interview": 0, "contrib": 1, "plan": 2}
    rest = sorted((a for a in cands if a["kind"] in rank), key=lambda a: (rank[a["kind"]], media_sort_key(a["media"])))
    seen = {_norm_title(a["title"])[:20] for a in picked}
    for a in rest:  # 같은 기사를 계열 매체가 함께 실은 경우(제목이 같음)는 한 번만
        if len(picked) >= n:
            break
        if _norm_title(a["title"])[:20] in seen:
            continue
        seen.add(_norm_title(a["title"])[:20])
        picked.append(a)
    return {a["url"] or a["title"] for a in picked}


# ---------------------------------------------------------------- 저장해 둔 지난 해 기록 (Secrets)

def add_saved_years(base, secrets, back=6):
    """올리지 않은 지난 해는 Secrets [monthly_history."고객사"."연도"] 에 저장해 둔 기록으로 채운다.
    (형식별 1~12월 건수 + '보도자료 배포' 수). 채운 연도 목록을 돌려준다."""
    have = {p["year"] for p in base["prev_years"]}
    added = []
    for y in range(base["year"] - back, base["year"]):
        if y in have:
            continue
        rows = history_from_secrets(secrets, base["company"], y)
        if not rows or not any(sum(v) for v in rows.values()):
            continue
        src = history_entry(secrets, base["company"], y) or {}
        try:
            rel = int(src.get("보도자료 배포", 0))
        except Exception:
            rel = 0
        base["prev_years"].append({"year": y, "releases": rel, "saved": True,
                                   "monthly": {k: [sum(rows[n][i] for n in names) for i in range(12)] for k, _, names in TYPES5}})
        added.append(y)
    base["prev_years"].sort(key=lambda p: p["year"])
    return added
