# -*- coding: utf-8 -*-
"""
결과보고서 작업 파일 (.json) 저장 / 불러오기.

Streamlit Cloud는 서버에 파일을 계속 보관할 수 없어서, 담당자가 작업 파일을 내려받아 두었다가
다음 날 다시 올리면 그동안의 선택을 그대로 이어서 쓸 수 있게 한다.
저장 내용: 보도자료 제목·배포일 / '포함' 해제한 기사 / 기사별 다음·구글 노출 / 다음·구글에서만 찾은 기사.
기사는 URL로 맞춰 보므로, 다시 검색해서 새로 생긴 기사는 그대로 추가되고 저장한 선택만 덮어씌워진다.
"""
import json
from datetime import date, datetime

from article_meta import normalize_url

VERSION = 1


def dump_work(title, dist_date, rows, included, extra, manual, company="") -> bytes:
    """rows: 이번에 수집한 네이버 기사 전체 / included: 그중 '포함'된 기사 /
    extra: {URL: {"다음","구글"}} / manual: 다음·구글에서만 찾은 기사 목록."""
    included_urls = {r["URL"] for r in included}
    data = {
        "version": VERSION,
        "title": title,
        "company": company,  # 기업명을 직접 입력했을 때만 값이 있음
        "distribution_date": dist_date.isoformat(),
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "excluded_urls": [r["URL"] for r in rows if r["URL"] not in included_urls],
        "naver_urls": [r["URL"] for r in rows],  # 저장 당시 확인한 네이버 기사 (포함/해제를 직접 정한 기사)
        "portals": {url: sorted(ps) for url, ps in extra.items() if ps},
        "manual": [
            {"매체명": m["매체명"], "제목": m["제목"], "URL": m["URL"],
             "게재일": m["게재일자"].date().isoformat(), "포털": sorted(m["포털"])}
            for m in manual
        ],
    }
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def load_work(raw: bytes) -> dict:
    try:
        data = json.loads(raw.decode("utf-8"))
        data["distribution_date"] = date.fromisoformat(data["distribution_date"])
        data["title"] = str(data["title"])
    except Exception:
        raise ValueError("작업 파일을 읽을 수 없습니다. '작업 저장'으로 받은 .json 파일인지 확인하세요.")
    data.setdefault("company", "")
    data.setdefault("excluded_urls", [])
    data.setdefault("naver_urls", [])
    data.setdefault("portals", {})
    data.setdefault("manual", [])
    return data


def restore_choices(saved, rows):
    """저장한 선택을 이번 검색 결과(rows)에 맞춰 돌려준다.
    반환: (제외할 URL 집합, {URL: {"다음","구글"}}, 다음·구글 전용 기사 목록, 저장 당시 있던 URL 집합)
    저장 당시 다음·구글에서만 찾은 기사가 이번에는 네이버에서도 잡혔으면 중복 추가하지 않고
    그 네이버 기사의 다음·구글 노출로 합친다."""
    by_norm = {normalize_url(r["URL"]): r["URL"] for r in rows}
    excluded = {by_norm[n] for n in map(normalize_url, saved["excluded_urls"]) if n in by_norm}
    known = {by_norm[n] for n in map(normalize_url, saved["naver_urls"] + saved["excluded_urls"]) if n in by_norm}
    portals = {}
    for url, ps in saved["portals"].items():
        if normalize_url(url) in by_norm:
            portals.setdefault(by_norm[normalize_url(url)], set()).update(ps)
    manual = []
    for m in saved["manual"]:
        naver_url = by_norm.get(normalize_url(m["URL"]))
        if naver_url:
            portals.setdefault(naver_url, set()).update(m["포털"])
            continue
        manual.append({
            "매체명": m["매체명"], "제목": m["제목"], "URL": m["URL"],
            "게재일자": datetime.combine(date.fromisoformat(m["게재일"]), datetime.min.time()),
            "포털": set(m["포털"]), "note": "",
        })
    return excluded, portals, manual, known
