# -*- coding: utf-8 -*-
"""
결과보고서 누적 엑셀(보도자료마다 시트가 쌓인 파일)을 읽어 '최근 보도 성과 비교' 수치를 만든다.

시트 형식 (report_sheet.build_sheet 와 같음):
  C3 제목 / C5 배포일 / C6 주제(직접 고친 짧은 주제, 없으면 비어 있음) / 8행부터 기사 (C 매체명, D 제목, G 게재 포털)
예전 시트에는 다른 회사 기사가 섞인 경우가 있어서, 기사 제목에 회사명이 있는 기사만 센다.
"""
import re
from datetime import date, datetime

TOPIC_ROW = 6  # 주제를 저장하는 행 (B6 '주제' / C6 값)
RECENT_COUNT = 6  # 비교 카드에 보여줄 최근 보도자료 수 (이번 건 제외)
REGULAR_RATIO = 0.5  # 올해 보도자료의 절반 이상에 실은 매체 = 꾸준히 싣는 매체


def _as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.search(r"(20\d\d)[-./](\d{1,2})[-./](\d{1,2})", str(v or ""))
    return date(int(m[1]), int(m[2]), int(m[3])) if m else None


def read_history(wb, company):
    """워크북의 보도자료 시트들 -> [{sheet, date, title, topic, total, media(set)}], 배포일 순."""
    out = []
    for ws in wb.worksheets:
        if ws["B3"].value != "제목":
            continue
        d = _as_date(ws["C5"].value)
        if not d:
            continue
        own = [r for r in ws.iter_rows(min_row=8, values_only=True)
               if len(r) > 3 and r[1] is not None and r[2] and company in str(r[3] or "")]
        topic = ws.cell(row=TOPIC_ROW, column=3).value if ws.cell(row=TOPIC_ROW, column=2).value == "주제" else None
        out.append({
            "sheet": ws.title, "date": d, "title": str(ws["C3"].value or ""),
            "topic": str(topic).strip() if topic else "",
            "total": len(own), "media": {r[2] for r in own},
        })
    out.sort(key=lambda h: h["date"])
    return out


def build_comparison(history, dist_date, total, naver_media, today):
    """history: read_history() 결과 / dist_date: 이번 배포일 / total: 이번 총 게재 기사 수 /
    naver_media: 이번 보고서 표 순서대로 네이버에 노출된 매체명 목록 / today: 보고서 만드는 날.
    비교할 기록이 없으면 None."""
    past = [h for h in history if h["date"] != dist_date]  # 같은 배포일 시트(다시 만드는 경우)는 제외
    if not past:
        return None
    year = [h for h in past if h["date"].year == today.year and h["date"] <= today]
    before = [h for h in past if h["date"] < dist_date]
    prev = before[-1] if before else None

    avg = sum(h["total"] for h in year) / len(year) if year else None
    freq_year = {}
    for h in year:
        for m in h["media"]:
            freq_year[m] = freq_year.get(m, 0) + 1
    seen_ever = set().union(*(h["media"] for h in past))
    return {
        "period": (date(today.year, 1, 1), today),
        "avg": avg,
        "vs_avg": (total - avg) / avg * 100 if avg else None,
        "prev_total": prev["total"] if prev else None,
        "vs_prev": (total - prev["total"]) / prev["total"] * 100 if prev and prev["total"] else None,
        "rank": 1 + sum(h["total"] > total for h in year),
        "recent": before[-RECENT_COUNT:],
        "first_media": [m for m in naver_media if m not in seen_ever],
        "regular_media": [m for m in naver_media if year and freq_year.get(m, 0) / len(year) >= REGULAR_RATIO],
    }
