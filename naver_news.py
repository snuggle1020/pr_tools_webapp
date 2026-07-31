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
import urllib.request
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


# 한국어는 명사 뒤에 조사가 띄어쓰기 없이 바로 붙는다("딥파인은", "딥파인이" 등).
# 그래서 순수 \b(단어경계) 정규식은 조사 결합을 인식 못 해서 본문 속 진짜 언급을
# 다 놓친다. 반대로 순수 부분일치(in)는 "에티버스이피에이"처럼 회사명을 포함한
# 다른 브랜드명까지 걸린다. 그래서 "회사명 + (흔한 조사 하나, 있으면) + 그 다음은
# 한글/영문/숫자가 이어지지 않음"을 하나의 패턴으로 검사한다.
_KOREAN_PARTICLES = sorted(
    [
        "이라고", "라고", "이라는", "라는", "이란", "란",
        "에서", "에게", "한테", "부터", "까지", "처럼", "보다", "께서", "으로", "이나",
        "은", "는", "이", "가", "을", "를", "의", "와", "과", "도", "만", "에", "로", "나",
    ],
    key=len,
    reverse=True,
)
_PARTICLE_PATTERN = "|".join(re.escape(p) for p in _KOREAN_PARTICLES)


def _word_pattern(term: str):
    return re.compile(rf"{re.escape(term)}(?:{_PARTICLE_PATTERN})?(?![가-힣A-Za-z0-9])")


def _contains_word(text: str, term: str) -> bool:
    """단순 부분일치(in)는 '에티버스이피에이'처럼 회사명을 포함한 다른 계열사/브랜드
    이름까지 걸려서 무관한 보도자료가 섞여 들어옴. 독립된 단어(조사 결합 포함)로
    등장할 때만 인정."""
    return _word_pattern(term).search(text) is not None


def _fetch_article_text(url, timeout=8):
    """기사 페이지를 받아서 태그를 제거한 순수 텍스트로 변환 (best-effort)"""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (compatible; ReportBot/1.0)"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw_html = resp.read(500_000).decode("utf-8", errors="ignore")
    except Exception:
        return None
    raw_html = re.sub(r"<script.*?</script>", " ", raw_html, flags=re.DOTALL | re.IGNORECASE)
    raw_html = re.sub(r"<style.*?</style>", " ", raw_html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", raw_html)
    text = re.sub(r"&nbsp;|&quot;|&amp;|&lt;|&gt;|&apos;", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _looks_like_press_release_mention(snippet: str, company: str) -> bool:
    """회사가 문장의 주어로 등장하며 '~라고 (며칠) 밝혔다/말했다/전했다' 같은
    보도자료 특유의 문장 패턴이 있는지 확인 (딥파인 PR리포트의 classify_articles.py와 동일 로직)."""
    kw = re.escape(company)
    patterns = [
        rf"{kw}\s*(은|는|이|가)\s*[^.]{{0,150}}(라고|다고)\s*\d{{0,2}}일?\s*(밝혔다|말했다|전했다)",
        rf"{kw}\s*[가-힣]{{0,10}}\s*(대표|CTO|CEO|본부장)[^.]{{0,20}}(는|은)?[^.]{{0,80}}(말했다|밝혔다|전했다)",
    ]
    return any(re.search(p, snippet) for p in patterns)


def _extract_snippets(text: str, term: str, window: int = 700, max_occurrences: int = 6):
    """term이 독립된 단어로 등장하는 모든 위치 주변 텍스트를 잘라내기 (단순 substring이
    아니라 단어경계 기준 - "에티버스ePA"/"에티버스이피에이" 같은 다른 브랜드명 안에
    끼어있는 경우는 무시). 첫 등장 위치가 메뉴/내비게이션이고 실제 문장은 더 뒤에 있는
    경우가 있어서(예: 언론사 사이트 상단 메뉴에 제목이 반복 노출), 첫 등장 위치 하나만
    보지 않고 여러 등장 위치를 순서대로 반환한다 (조사 결합도 인식하는 _word_pattern 사용
    - classify_articles.py의 extract_relevant_paragraphs와 동일 로직 + 조사 인식 매칭)."""
    pattern = _word_pattern(term)
    snippets = []
    for m in pattern.finditer(text):
        if len(snippets) >= max_occurrences:
            break
        start = max(0, m.start() - 50)
        end = min(len(text), m.start() + window)
        snippets.append(text[start:end])
    return snippets


def _relevance_status(title: str, url: str, company: str, description: str) -> str:
    """require_text 정밀 검증. 'excluded' / 'ambiguous' / 'confirmed' 중 하나를 반환.

    - 제목이 회사명으로 시작 -> 'confirmed' (관례상 확정 보도자료, 네트워크 요청 없이 빠름)
    - 그 외에는 기사 본문을 직접 열어서 확인한다. 네이버가 주는 짧은 요약(description)만
      보고 "요약에 없으니 무관"으로 판단하면, [게시판]/[Tech & Now] 같은 여러 기업을
      한꺼번에 다루는 다이제스트형 기사에서 실제로는 회사명이 본문에 나오는데 요약에는
      빠져있어서 잘못 제외되는 문제가 있었다 (요약은 보통 첫 문단만 담고, 다이제스트
      기사는 회사 언급이 뒤쪽 항목에 있는 경우가 많음). 그래서 본문 전체를 확인해서:
        - 회사명이 본문에 독립된 단어로 아예 없으면 'excluded' (다른 계열사/브랜드명
          예: "에티버스이피에이"의 무관한 기사가 섞이는 것도 여기서 걸러짐)
        - 있으면, 등장하는 위치마다(첫 등장 위치만 보면 메뉴 등에 걸려 놓칠 수 있어서)
          보도자료 특유의 문장 패턴("~라고 밝혔다" 등)이 있는지 확인해서 있으면 'confirmed'
        - 언급은 있지만 그 패턴을 못 찾으면 'ambiguous'

    'ambiguous'는 실제로 관련 있는 기사인데 문장 패턴이 다르게 쓰인 경우(예: "체결했다")도
    있을 수 있어서, 제외하지 않고 결과에는 포함하되 표시만 해서 사람이 확인하게 한다
    (여기서 확신 없다고 제외하면 진짜 기사를 놓칠 위험이 있음)."""
    if title.strip().startswith(company):
        return "confirmed"

    text = _fetch_article_text(url)
    if text is None:
        # 본문을 못 가져온 경우: 제목/요약에 단어로도 없으면 무관, 있으면 사람 확인
        if _contains_word(title, company) or _contains_word(description, company):
            return "ambiguous"
        return "excluded"

    snippets = _extract_snippets(text, company)
    if not snippets:
        return "excluded"  # 본문에도 독립된 단어로 전혀 없음 (다른 브랜드명 등) -> 확실히 무관

    for snippet in snippets:
        if _looks_like_press_release_mention(snippet, company):
            return "confirmed"
    return "ambiguous"


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


def collect_coverage(
    query: str,
    distribution_date: str,
    window_days: int = 3,
    require_text: str = None,
    verify_relevance: bool = True,
):
    """
    보도자료 배포일(distribution_date, 'YYYY-MM-DD') 기준으로
    배포일 당일부터 이후 window_days일 이내에 게재된 관련 기사만 걸러서 반환
    (배포일 이전 기사는 제외).

    require_text: 지정하면(보통 기업명) 제목/요약 어디에도 독립된 단어로 없는 기사는
    제외한다 (기업명을 포함한 다른 계열사/브랜드명, 예: "에티버스이피에이"의 무관한
    기사가 섞이는 것을 막음).

    verify_relevance: True(기본값)면 한 단계 더 나아가 "진짜 이 기업에 관한 보도자료가
    맞는지"까지 검증한다 - 제목이 기업명으로 시작하면 확정 통과, 애매하면(본문에
    언급은 있지만 "~라고 밝혔다" 같은 보도자료 문장 패턴을 못 찾은 경우) 실제로 관련
    기사인데 표현만 다를 수 있어서(예: "체결했다") 제외하지 않고 제목 앞에
    "[확인필요] "를 붙여 포함한다. 이건 "이 특정 보도자료의 커버리지만 뽑고 싶다"는
    결과보고서/PR리포트용 정밀 검증이라 느리다(기사 본문을 열어봄).
    False면 require_text가 단어로 등장하기만 하면 다 포함한다(빠름, 본문 조회 없음)
    - "이 키워드 걸린 기사를 그냥 다 모아줘"라는 네이버 기사 수집 탭에 맞는 동작.

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
        if delta_days < 0 or delta_days > window_days:
            continue

        url = item.get("originallink") or item.get("link")
        if not url or url in seen_urls:
            continue

        title = _strip_tags(item.get("title", ""))
        description = _strip_tags(item.get("description", ""))
        if require_text and not verify_relevance:
            if not _contains_word(title, require_text) and not _contains_word(description, require_text):
                continue
        elif require_text:
            status = _relevance_status(title, url, require_text, description)
            if status == "excluded":
                continue
            if status == "ambiguous":
                title = f"[확인필요] {title}"

        seen_urls.add(url)

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
