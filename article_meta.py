# -*- coding: utf-8 -*-
"""
기사 URL 하나로 매체명 / 제목 / 게재일을 읽어오는 모듈.
결과보고서 탭에서 다음·구글에서 찾은 기사 링크를 붙여넣으면 이걸로 표를 자동으로 채운다.

- 매체명: domain_to_media.json -> pr_report_core.MEDIA_MAP -> 페이지의 og:site_name -> 도메인
- 제목:   og:title -> twitter:title -> <title> (끝에 붙은 " - 매체명" 등은 떼어냄)
- 게재일: article:published_time 등 발행일 메타태그 -> JSON-LD datePublished
페이지를 못 열거나 값을 못 찾은 칸은 비워두고 note 에 이유를 적는다 (앱에서 직접 고치도록).
"""
import html as html_lib
import re
import urllib.request
from datetime import datetime

from pr_report_core import MEDIA_MAP, _decode_html
from report_sheet import domain_of, load_domain_map

_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

_DATE_KEYS = [
    "article:published_time", "og:article:published_time", "og:published_time", "datePublished",
    "pubdate", "publishdate", "publish-date", "article:published", "date", "dc.date", "dable:published_time",
]


def _unescape(text):
    """일부 매체는 &amp;quot; 처럼 두 번 인코딩해 둬서 두 번 푼다."""
    return html_lib.unescape(html_lib.unescape(text))


def _meta(page, key):
    k = re.escape(key)
    for pat in (
        rf'<meta[^>]+(?:property|name|itemprop)\s*=\s*["\']{k}["\'][^>]*?content\s*=\s*(["\'])(.*?)\1',
        rf'<meta[^>]+content\s*=\s*(["\'])(.*?)\1[^>]*?(?:property|name|itemprop)\s*=\s*["\']{k}["\']',
    ):
        m = re.search(pat, page, re.I | re.S)
        if m and m.group(2).strip():
            return _unescape(re.sub(r"<.*?>", "", m.group(2))).strip()
    return None


def _parse_date(text):
    m = re.search(r"(20\d{2})[-./년]\s*(\d{1,2})[-./월]\s*(\d{1,2})", text or "")
    if not m:
        return None
    try:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _media_name(url, page, domain_map):
    domain = domain_of(url)
    bare = re.sub(r"^[^.]+\.", "", domain)
    for d in (domain, bare):
        if d in domain_map:
            return domain_map[d]
        if d in MEDIA_MAP:
            return MEDIA_MAP[d]
    return (page and _meta(page, "og:site_name")) or domain


def _clean_title(title, media):
    """'기사 제목 - 매체명', '[매체명] 기사 제목' 처럼 앞뒤에 붙은 매체명을 떼어냄."""
    m = re.match(r"\[([^\]]+)\]\s*", title)
    if m and media and media in m.group(1):
        title = title[m.end():]
    for sep in (" - ", " | ", " :: ", " : ", " – "):
        if sep in title:
            head, tail = title.rsplit(sep, 1)
            if media and (media in tail or tail in media):
                return head.strip()
    return title.strip()


def normalize_url(url):
    """같은 기사인지 비교용: 스킴·www·끝 슬래시 제거."""
    u = re.sub(r"^https?://", "", (url or "").strip().lower())
    return re.sub(r"^www\.", "", u).rstrip("/")


def fetch_article_meta(url, domain_map=None):
    """반환: {"URL", "매체명", "제목", "게재일자"(datetime|None), "note"}"""
    domain_map = domain_map if domain_map is not None else load_domain_map()
    page, note = None, []
    try:
        req = urllib.request.Request(url, headers={"User-Agent": _UA})
        with urllib.request.urlopen(req, timeout=8) as resp:
            page = _decode_html(resp.read(600_000))
    except Exception:
        note.append("페이지를 열 수 없음")

    media = _media_name(url, page, domain_map)
    title, pub = "", None
    if page:
        title = _meta(page, "og:title") or _meta(page, "twitter:title") or ""
        if not title:
            m = re.search(r"<title[^>]*>(.*?)</title>", page, re.I | re.S)
            title = _unescape(re.sub(r"\s+", " ", m.group(1))).strip() if m else ""
        title = _clean_title(title, media)
        for key in _DATE_KEYS:
            pub = _parse_date(_meta(page, key))
            if pub:
                break
        if not pub:
            m = re.search(r'"datePublished"\s*:\s*"([^"]+)"', page)
            pub = _parse_date(m.group(1)) if m else None
        if not title:
            note.append("제목 못 찾음")
        if not pub:
            note.append("게재일 못 찾음")
    return {"URL": url.strip(), "매체명": media, "제목": title, "게재일자": pub, "note": ", ".join(note)}


def _plain(text):
    return re.sub(r"[^0-9a-z가-힣]", "", (text or "").lower())


def complete_truncated_titles(rows, domain_map=None):
    """네이버 검색 API는 긴 제목을 '...'로 잘라서 준다. 끝이 잘린 제목은 기사 원문에서 전체 제목을 읽어 바꾼다.
    원문 제목이 잘린 제목으로 시작할 때만 바꿔서 엉뚱한 제목으로 덮어쓰지 않는다. 바꾼 건수 반환."""
    domain_map = domain_map if domain_map is not None else load_domain_map()
    fixed = 0
    for r in rows:
        title = r["제목"].rstrip()
        if not title.endswith(("...", "…")):
            continue
        head = _plain(title.rstrip(".…"))
        full = fetch_article_meta(r["URL"], domain_map)["제목"]
        if full and len(_plain(full)) > len(head) and _plain(full).startswith(head):
            r["제목"] = full
            fixed += 1
    return fixed
