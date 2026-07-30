# -*- coding: utf-8 -*-
"""
네이버 뉴스 검색 API 호출 모듈 (결과보고서용: 단일 배포일 기준 전후 N일 검색)

인증 정보는 환경변수 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 에서 읽는다.
(app.py가 Streamlit secrets를 이 환경변수로 옮겨놓음)
"""

import os
import re
import html
import time
from datetime import datetime

import requests


def load_config():
    client_id = os.environ.get("NAVER_CLIENT_ID")
    client_secret = os.environ.get("NAVER_CLIENT_SECRET")
    if not client_id or not client_secret:
        raise RuntimeError(
            "NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이(가) 설정되어 있지 않습니다. "
            "Streamlit Cloud라면 App settings > Secrets에 등록하세요."
        )
    return {"naver_client_id": client_id, "naver_client_secret": client_secret}


def _strip_tags(text: str) -> str:
    """네이버 API 응답의 <b>강조태그</b>와 HTML 엔티티 제거"""
    text = re.sub(r"</?b>", "", text)
    return html.unescape(text)


def _parse_pubdate(pubdate_str: str) -> datetime:
    # 예: 'Thu, 16 Jul 2026 10:04:00 +0900'
    return datetime.strptime(pubdate_str, "%a, %d %b %Y %H:%M:%S %z")


def search_news(query: str, display: int = 100, start: int = 1, sort: str = "date"):
    """네이버 뉴스 검색 API 단일 호출 (최대 display=100)"""
    cfg = load_config()
    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id": cfg["naver_client_id"],
        "X-Naver-Client-Secret": cfg["naver_client_secret"],
    }
    params = {"query": query, "display": display, "start": start, "sort": sort}
    r = requests.get(url, headers=headers, params=params, timeout=10)
    r.raise_for_status()
    return r.json()


def search_all(query: str, max_results: int = 300, sort: str = "date"):
    """페이지네이션 처리해서 여러 페이지 결과를 이어붙여 반환"""
    results = []
    start = 1
    while start <= min(max_results, 1000):
        data = search_news(query, display=100, start=start, sort=sort)
        items = data.get("items", [])
        if not items:
            break
        results.extend(items)
        if len(items) < 100:
            break
        start += 100
        time.sleep(0.2)  # 과도한 연속 호출 방지
    return results


def collect_coverage(query: str, distribution_date: str, window_days: int = 3):
    """
    보도자료 배포일(distribution_date, 'YYYY-MM-DD') 기준으로
    전후 window_days일 이내에 게재된 관련 기사만 걸러서 반환.

    반환 형식: [{"매체명": ..., "제목": ..., "URL": ..., "게재일자": datetime, "게재포털": "네이버"}, ...]
    URL 기준 중복 제거, 게재일자 오름차순 정렬.
    """
    dist_dt = datetime.strptime(distribution_date, "%Y-%m-%d")
    raw_items = search_all(query)

    seen_urls = set()
    coverage = []
    for item in raw_items:
        try:
            pub_dt = _parse_pubdate(item["pubDate"]).replace(tzinfo=None)
        except (KeyError, ValueError):
            continue

        delta_days = (pub_dt.date() - dist_dt.date()).days
        if abs(delta_days) > window_days:
            continue

        url = item.get("originallink") or item.get("link")
        if not url or url in seen_urls:
            continue
        seen_urls.add(url)

        title = _strip_tags(item.get("title", ""))

        coverage.append(
            {
                "매체명": "",  # app.py에서 도메인 매핑으로 채움
                "제목": title,
                "URL": url,
                "게재일자": pub_dt,
                "게재포털": "네이버",  # API 출처 표시. 구글/다음 게재 여부는 수동 확인 필요
            }
        )

    coverage.sort(key=lambda x: x["게재일자"])
    return coverage
