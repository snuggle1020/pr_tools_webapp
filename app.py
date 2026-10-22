# -*- coding: utf-8 -*-
"""
PR 리포트 도구 웹앱.
- 탭1: 결과보고서 (보도자료 1건 배포 결과 - 네이버 뉴스 수집 후 엑셀 시트 + PDF 보고서 생성)
- 탭2: 월간 PR리포트 (한 달 기사 수집·자동분류 -> 앱에서 확인·수정 -> PDF, 월별 기록은 PDF에 숨겨 다음 달로 이어감)

네이버 API 키는 Streamlit Cloud의 App settings > Secrets 에
NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 로 등록해두면 자동으로 읽힘.
"""
import io
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone

import openpyxl
import pandas as pd
import streamlit as st

from naver_news import collect_coverage
from report_sheet import (
    EXTRA_PORTALS, MIN_SHARED_CHUNKS, apply_manual_portals, build_sheet, derive_company, derive_keyword,
    load_domain_map, resolve_coverage, title_overlap, write_topics,
)
from report_pdf import auto_topic, build_pdf
from report_history import build_comparison, read_history
from article_meta import complete_truncated_titles, fetch_article_meta, normalize_url
from work_file import dump_work, load_work, restore_choices
from pr_report_core import fetch_month_articles, classify_row, clean_html
from daily_collect_sheet import build_daily_sheet
from monthly_pdf import TYPE_ORDER, assign_release, build_monthly_pdf, default_summary, read_history
from monthly_report import (
    KIND_OPTIONS, auto_group_mentions, auto_mark_releases, compose, guess_releases, history_from_secrets, initial_kind, sort_articles, trend_rows_from_xlsx,
    dump_work as dump_month_work, load_work as load_month_work,
)

REPORT_WINDOW_DAYS = 3  # 배포일 이후 며칠까지 검색할지 (고정)

@st.cache_data(show_spinner=False, max_entries=8)
def _pdf_page_images(pdf_bytes: bytes):
    """PDF 각 쪽을 PNG 이미지로 (앱 안 미리보기용)."""
    import pypdfium2 as pdfium

    pages = []
    for page in pdfium.PdfDocument(pdf_bytes):
        img = io.BytesIO()
        page.render(scale=1.6).to_pil().save(img, "PNG")
        pages.append(img.getvalue())
    return pages


st.set_page_config(page_title="오픈피알_PR 보고서", page_icon="📰", layout="wide")


def _load_naver_credentials():
    cid = os.environ.get("NAVER_CLIENT_ID")
    csecret = os.environ.get("NAVER_CLIENT_SECRET")
    if not cid or not csecret:
        try:
            cid = st.secrets["NAVER_CLIENT_ID"]
            csecret = st.secrets["NAVER_CLIENT_SECRET"]
            os.environ["NAVER_CLIENT_ID"] = cid
            os.environ["NAVER_CLIENT_SECRET"] = csecret
        except Exception:
            cid, csecret = None, None
    return cid, csecret


NAVER_CLIENT_ID, NAVER_CLIENT_SECRET = _load_naver_credentials()


st.title("📰 PR 보고서 생성기")

if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
    st.error(
        "네이버 API 키가 설정되어 있지 않습니다. 관리자에게 문의하거나, "
        "로컬 실행 시 `.streamlit/secrets.toml`에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 을 등록하세요."
    )

tab1, tab2, tab3 = st.tabs(["📄 보도자료 결과보고서", "📅 월간 PR리포트", "🔎 네이버 기사 수집"])

# ------------------------------------------------------------------
# 탭 1: 결과보고서
# ------------------------------------------------------------------
with tab1:
    st.caption(
        "보도자료 제목만 넣으면 기업명/검색어는 자동으로 뽑아냅니다. "
        f"배포일 당일부터 {REPORT_WINDOW_DAYS}일 뒤까지 게재된 기사만 모읍니다."
    )

    work_file = st.file_uploader(
        "이전 작업 불러오기 (아래 '💾 작업 저장'으로 받아둔 .json 파일을 올리면 이어서 작업할 수 있습니다)",
        type=["json"],
        key="work_file",
    )
    if work_file is not None and st.session_state.get("work_file_id") != work_file.file_id:
        try:
            saved_work = load_work(work_file.getvalue())
        except ValueError as e:
            st.error(str(e))
        else:
            st.session_state["work_file_id"] = work_file.file_id
            st.session_state["saved_work"] = saved_work
            st.session_state["report_title_input"] = saved_work["title"]
            st.session_state["report_date_input"] = saved_work["distribution_date"]
            st.session_state["report_company_input"] = saved_work.get("company", "")
    if work_file is None and st.session_state.pop("work_file_id", None):  # 올렸던 파일을 뺐으면 불러온 작업도 해제
        st.session_state.pop("saved_work", None)
    if st.session_state.get("saved_work"):
        st.caption(
            "불러온 작업이 아래에 채워졌습니다. **리포트 생성**을 누르면 네이버를 다시 검색하고, "
            "저장해 둔 포함 해제·다음·구글 노출·직접 추가한 기사를 그대로 적용합니다."
        )

    st.session_state.setdefault("report_date_input", date.today())
    with st.form("report_form"):
        title = st.text_input("보도자료 제목", key="report_title_input",
                              placeholder="예: 딥파인, 국방·산업 현장용 스마트글래스 기반 AI 에이전트 개발 나선다")
        distribution_date = st.date_input("배포일", key="report_date_input")
        company_input = st.text_input(
            "기업명 (자동 인식이 틀릴 때만 입력)", key="report_company_input",
            placeholder="비워두면 보도자료 제목에서 자동으로 찾습니다. 이 이름으로 네이버를 검색합니다.",
        )
        existing_file = st.file_uploader(
            "기존 결과보고서 파일 (파일을 추가하면 맨 왼쪽에 새 시트를 추가합니다 / 없으면 새 파일을 생성합니다)",
            type=["xlsx"],
            key="report_existing_wb",
        )
        submitted = st.form_submit_button("리포트 생성", type="primary")

    if submitted:
        if not title:
            st.error("보도자료 제목을 입력하세요.")
        elif not NAVER_CLIENT_ID:
            st.error("네이버 API 키가 없어 진행할 수 없습니다.")
        else:
            company = company_input.strip() or derive_company(title)
            keyword = company_input.strip() or derive_keyword(title)
            dist_str = distribution_date.strftime("%Y-%m-%d")
            with st.spinner(f"네이버에서 '{keyword}' 검색 중..."):
                try:
                    coverage = collect_coverage(
                        keyword, dist_str, window_days=REPORT_WINDOW_DAYS, require_text=company
                    )
                except Exception as e:
                    st.error(f"뉴스 수집 실패: {e}")
                    coverage = None

            if coverage is not None:
                with st.spinner("고정 목록에 없는 매체는 기사 페이지에서 매체명 자동 인식 중..."):
                    rows, unmapped, auto_detected = resolve_coverage(coverage)
                with st.spinner("네이버에서 잘려서 온 기사 제목을 원문에서 확인 중..."):
                    complete_truncated_titles(rows)
                # 다음·구글 수동 입력으로 화면이 다시 그려져도(rerun) 결과가 유지되도록 세션에 보관
                st.session_state["report"] = {
                    "title": title,
                    "company_input": company_input.strip(),
                    "company": company,
                    "keyword": keyword,
                    "dist_date": distribution_date,
                    "rows": rows,
                    "unmapped": unmapped,
                    "auto_detected": auto_detected,
                    "existing": existing_file.getvalue() if existing_file is not None else None,
                    # 기존 결과보고서 엑셀이 있으면 이전 보도자료 기록 (최근 보도 성과 비교용)
                    "history": read_history(openpyxl.load_workbook(io.BytesIO(existing_file.getvalue()), data_only=True), company)
                    if existing_file is not None else [],
                }
                run = st.session_state.get("report_run", 0) + 1
                st.session_state["report_run"] = run
                saved_work = st.session_state.get("saved_work")
                if saved_work and saved_work["title"] == title:
                    excluded, portals, manual_saved, known = restore_choices(saved_work, rows)
                    st.session_state["report"]["restore"] = {"excluded": excluded, "portals": portals, "known": known}
                    st.session_state["report"]["saved_topics"] = saved_work.get("topics", {})
                    st.session_state[f"fetched_{run}"] = {"matched": {}, "fetched": manual_saved, "n": 0}

    report = st.session_state.get("report")
    if report:
        title, company = report["title"], report["company"]
        dist_date = report["dist_date"]
        dist_str = dist_date.strftime("%Y-%m-%d")
        rows = report["rows"]
        run = st.session_state["report_run"]  # 새로 생성할 때마다 입력표 초기화

        st.caption(f"자동 인식: 기업명 **{company}** / 검색어 **{report['keyword']}**")
        st.success(f"완료! 네이버 기사 {len(rows)}건 수집됨 (시트: {dist_date:%y%m%d}).")
        if report["auto_detected"]:
            lines = "\n".join(f"- {d} → {n}" for d, n in sorted(report["auto_detected"].items()))
            st.info(
                "다음 매체는 고정 목록에 없어서 기사 페이지에서 자동으로 인식했습니다 "
                "(맞는지 확인해보고, domain_to_media.json에 추가해두면 다음부턴 더 빨라져요):\n"
                + lines
            )
        if report["unmapped"]:
            st.warning(
                "매체명을 자동으로도 못 찾은 도메인: " + ", ".join(sorted(report["unmapped"]))
                + " — 직접 확인해서 domain_to_media.json에 추가해주세요."
            )

        st.markdown("##### 기사 확인 및 다음·구글 반영")
        st.caption(
            "**⚠ 확인** 표시는 제목이 보도자료와 많이 달라 인터뷰·다른 소식일 수 있는 기사입니다. "
            "**⚠ 본문 확인 필요** 기사는 처음에 **포함**이 해제돼 있으니, 넣을 기사만 체크하세요. "
            "보고서에서 뺄 기사는 **포함** 체크를 해제하세요. "
            "다음·구글은 자동 수집이 안 되니 직접 확인한 노출을 체크하면 PDF·엑셀에 바로 반영됩니다."
        )

        def _check_note(r):
            if r.get("확인필요"):
                return "⚠ 본문 확인 필요"
            if title_overlap(title, company, r["제목"]) < MIN_SHARED_CHUNKS:
                return "⚠ 제목이 보도자료와 다름"
            return ""

        notes = [_check_note(r) for r in rows]
        restore = report.get("restore") or {"excluded": set(), "portals": {}, "known": set()}

        def _default_include(r):
            # 작업 파일에서 이미 포함/해제를 정한 기사는 그 선택대로, 처음 보는 기사는 '본문 확인 필요'면 해제 상태로 시작
            if r["URL"] in restore.get("known", set()):
                return r["URL"] not in restore["excluded"]
            return not r.get("확인필요")
        if report.get("restore"):
            st.info("작업 파일의 선택(포함 해제·다음·구글 노출·직접 추가한 기사)을 적용했습니다. 새로 검색된 기사는 그대로 추가됐습니다.")
        order = range(len(rows))  # 보고서와 같은 순서 (포털 > 매체 영향력)
        flags = st.data_editor(
            pd.DataFrame(
                [{"포함": _default_include(rows[i]), "확인": notes[i],
                  "매체명": rows[i]["매체명"], "제목": rows[i]["제목"], "게재일": rows[i]["게재일자"].date(),
                  "다음": "다음" in restore["portals"].get(rows[i]["URL"], ()),
                  "구글": "구글" in restore["portals"].get(rows[i]["URL"], ())} for i in order],
                columns=["포함", "확인", "매체명", "제목", "게재일", "다음", "구글"],
            ),
            disabled=["확인", "매체명", "제목", "게재일"],
            hide_index=True,
            width="stretch",
            key=f"portal_flags_{run}",
            column_config={
                "포함": st.column_config.CheckboxColumn("포함", width="small"),
                "확인": st.column_config.TextColumn("확인", width="medium"),
                "다음": st.column_config.CheckboxColumn("다음에도 노출", width="small"),
                "구글": st.column_config.CheckboxColumn("구글에도 노출", width="small"),
            },
        )
        included, extra = [], {}
        for pos, i in enumerate(order):
            if not bool(flags.iloc[pos]["포함"]):
                continue
            included.append(rows[i])
            extra[rows[i]["URL"]] = {p for p in EXTRA_PORTALS if bool(flags.iloc[pos][p])}
        n_excluded = len(rows) - len(included)
        if n_excluded:
            st.caption(f"포함 해제: {n_excluded}건은 PDF·엑셀에서 빠집니다.")

        st.markdown("###### 다음·구글에서만 찾은 기사")
        st.caption(
            "기사 링크를 한 줄에 하나씩 붙여넣고 **기사 정보 불러오기**를 누르면 매체명·제목·게재일을 자동으로 채웁니다. "
            "네이버 목록에 이미 있는 기사 링크면 그 기사에 다음·구글 노출로 반영됩니다."
        )
        u1, u2 = st.columns(2)
        daum_text = u1.text_area("다음에서 찾은 기사 URL", key=f"daum_urls_{run}", height=110,
                                 placeholder="https://www.example.co.kr/news/...")
        google_text = u2.text_area("구글에서 찾은 기사 URL", key=f"google_urls_{run}", height=110,
                                   placeholder="https://www.example.co.kr/news/...")
        fetched_key = f"fetched_{run}"
        if st.button("기사 정보 불러오기", key=f"fetch_btn_{run}"):
            portals_by_url = {}
            for portal, text in (("다음", daum_text), ("구글", google_text)):
                for u in re.findall(r"https?://[^\s<>\"']+", text or ""):
                    portals_by_url.setdefault(u.rstrip(".,)"), set()).add(portal)
            naver_by_url = {normalize_url(r["URL"]): r["URL"] for r in rows}
            prev = st.session_state.get(fetched_key, {"matched": {}, "fetched": [], "n": 0})
            # 지금 표에 있는 기사(불러온 작업·이전에 불러온 기사·직접 고친 내용)는 그대로 두고 새 기사만 아래에 추가
            fetched = [dict(m) for m in st.session_state.get(f"manual_table_{run}", prev["fetched"])]
            by_url = {normalize_url(m["URL"]): m for m in fetched if m["URL"]}
            matched = {u: set(ps) for u, ps in prev["matched"].items()}
            with st.spinner(f"기사 {len(portals_by_url)}건 정보 불러오는 중..."):
                domain_map = load_domain_map()
                for u, ps in portals_by_url.items():
                    naver_url = naver_by_url.get(normalize_url(u))
                    if naver_url:
                        matched.setdefault(naver_url, set()).update(ps)
                        continue
                    if normalize_url(u) in by_url:  # 이미 표에 있는 기사면 포털만 더함
                        by_url[normalize_url(u)]["포털"] = set(by_url[normalize_url(u)]["포털"]) | ps
                        continue
                    meta = fetch_article_meta(u, domain_map)
                    meta["포털"] = ps
                    fetched.append(meta)
                    by_url[normalize_url(u)] = meta
            st.session_state[fetched_key] = {"matched": matched, "fetched": fetched, "n": prev["n"] + 1}
        fetched_state = st.session_state.get(fetched_key, {"matched": {}, "fetched": [], "n": 0})

        for url, ps in fetched_state["matched"].items():
            if url in extra:  # 포함 해제한 기사는 그대로 제외
                extra[url] |= ps
        if fetched_state["matched"]:
            st.caption(f"네이버 목록에 있는 기사 {len(fetched_state['matched'])}건은 다음·구글 노출로 반영했습니다.")

        manual_cols = ["매체명", "제목", "게재일", "다음", "구글", "URL", "비고"]
        manual_src = pd.DataFrame(
            [{"매체명": m["매체명"], "제목": m["제목"],
              "게재일": m["게재일자"].date() if m["게재일자"] else dist_date,
              "다음": "다음" in m["포털"], "구글": "구글" in m["포털"], "URL": m["URL"],
              "비고": m["note"] + (" → 게재일은 배포일로 넣음" if not m["게재일자"] else "")}
             for m in fetched_state["fetched"]],
            columns=manual_cols,
        ).astype({"매체명": "str", "제목": "str", "다음": "bool", "구글": "bool", "URL": "str", "비고": "str"})
        if fetched_state["fetched"]:
            st.caption("불러온 내용이 틀리면 표에서 바로 고치세요. 뺄 기사는 행을 선택해 삭제하면 됩니다.")
        manual_df = st.data_editor(
            manual_src,
            num_rows="dynamic",
            hide_index=True,
            width="stretch",
            disabled=["비고"],
            key=f"manual_articles_{run}_{fetched_state['n']}",
            column_config={
                "게재일": st.column_config.DateColumn("게재일", default=dist_date, format="YYYY-MM-DD"),
                "다음": st.column_config.CheckboxColumn("다음", default=False, width="small"),
                "구글": st.column_config.CheckboxColumn("구글", default=True, width="small"),
                "비고": st.column_config.TextColumn("비고", width="medium"),
            },
        )
        manual, skipped, table_now = [], 0, []
        for _, m in manual_df.iterrows():
            media, art_title, url = (
                "" if pd.isna(m.get(k)) else str(m.get(k)).strip() for k in ("매체명", "제목", "URL")
            )
            portals = [p for p in EXTRA_PORTALS if m.get(p) == True]  # noqa: E712 (NaN/None 은 False)
            if not (media or art_title or url):
                continue  # 빈 행
            pub = m.get("게재일")
            pub = pub if isinstance(pub, date) and not pd.isna(pub) else dist_date
            note = m.get("비고")
            table_now.append({
                "매체명": media, "제목": art_title, "URL": url,
                "게재일자": datetime.combine(pub, datetime.min.time()), "포털": set(portals),
                "note": "" if note is None or pd.isna(note) else str(note),
            })
            if not (media and art_title and url and portals):
                skipped += 1
                continue
            manual.append({
                "매체명": media, "제목": art_title, "URL": url,
                "게재일자": datetime.combine(pub, datetime.min.time()), "포털": portals,
            })
        # 다음에 '기사 정보 불러오기'를 누를 때 지금 표 내용을 그대로 이어가기 위해 보관
        st.session_state[f"manual_table_{run}"] = table_now
        if skipped:
            st.warning(f"매체명·제목·URL·포털(다음/구글) 중 빈 칸이 있는 {skipped}개 행은 보고서에서 빠졌습니다.")

        final_rows = apply_manual_portals(included, extra, manual)

        # 최근 보도 성과 비교 (기존 결과보고서 엑셀을 올렸을 때만): 수치는 이번 보고서 기준으로 매번 새로 계산
        today_kst = datetime.now(timezone(timedelta(hours=9))).date()
        naver_media = list(dict.fromkeys(r["매체명"] for r in final_rows if "네이버" in r.get("포털", ["네이버"])))
        comparison = build_comparison(report.get("history", []), dist_date, len(final_rows), naver_media, today_kst)
        topics = {}
        if comparison:
            saved_topics = report.get("saved_topics", {})
            cmp_rows = [{"date": h["date"], "topic": saved_topics.get(h["date"]) or h["topic"] or auto_topic(h["title"], company),
                         "total": h["total"]} for h in comparison["recent"]]
            cmp_rows.append({"date": dist_date, "topic": saved_topics.get(dist_date) or auto_topic(title, company),
                             "total": len(final_rows), "now": True})
            st.markdown("##### 최근 보도 성과 비교 (PDF 마지막 쪽)")
            st.caption(
                "올린 결과보고서 엑셀의 이전 보도자료와 비교합니다. 주제가 길면 PDF에서 두 줄로 나옵니다. "
                "주제를 고치면 PDF와 엑셀(각 시트 6행 '주제')에 함께 저장돼 다음 보고서에서도 그대로 쓰입니다."
            )
            edited = st.data_editor(
                pd.DataFrame([{"배포일": r["date"], "주제": r["topic"], "게재 수": r["total"]} for r in cmp_rows]),
                disabled=["배포일", "게재 수"],
                hide_index=True,
                width="stretch",
                key=f"cmp_topics_{run}",
                column_config={
                    "배포일": st.column_config.DateColumn("배포일", format="YYYY-MM-DD", width="small"),
                    "주제": st.column_config.TextColumn("주제", width="large"),
                    "게재 수": st.column_config.NumberColumn("게재 수", format="%d건", width="small"),
                },
            )
            for r, topic in zip(cmp_rows, edited["주제"]):
                r["topic"] = str(topic).strip() if topic is not None and str(topic).strip() else r["topic"]
                topics[r["date"]] = r["topic"]
            comparison["rows"] = cmp_rows

        if report["existing"] is not None:
            wb = openpyxl.load_workbook(io.BytesIO(report["existing"]))
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        build_sheet(wb, dist_date.strftime("%y%m%d"), title, company, dist_str, final_rows)
        if topics:
            write_topics(wb, topics)
        wb.active = 0  # 새로 추가된(맨 왼쪽) 시트가 열리도록
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        pdf_bytes = build_pdf(title, company, dist_str, final_rows, window_days=REPORT_WINDOW_DAYS,
                              generated_on=today_kst, comparison=comparison)

        n_extra = sum(len(v) > 0 for v in extra.values()) + len(manual)
        if n_extra:
            st.caption(f"다음·구글 반영: {n_extra}건 (보고서 총 {len(final_rows)}건)")

        filename = f"결과보고서_{title}_{date.today():%Y%m%d}.xlsx"
        # on_click="ignore": 한쪽을 받아도 화면이 다시 그려지지 않아서 나머지 버튼이 유지됨
        d1, d2, d3, d4 = st.columns(4)
        d1.download_button(
            "📑 PDF 보고서 다운로드",
            data=pdf_bytes,
            file_name=filename.removesuffix(".xlsx") + ".pdf",
            mime="application/pdf",
            type="primary",
            width="stretch",
            on_click="ignore",
        )
        d2.download_button(
            "⬇️ 엑셀 다운로드",
            data=buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            on_click="ignore",
        )
        d3.download_button(
            "💾 작업 저장",
            data=dump_work(title, dist_date, rows, included, extra, manual, company=report.get("company_input", ""), topics=topics),
            file_name=f"작업_{company}_{dist_date:%y%m%d}.json",
            mime="application/json",
            width="stretch",
            on_click="ignore",
            help="내일 이어서 하려면 이 파일을 받아두었다가 맨 위 '이전 작업 불러오기'에 올리세요.",
        )
        preview_open = st.session_state.get("pdf_preview_open", False)
        if d4.button("🙈 미리보기 닫기" if preview_open else "👁 PDF 미리보기", width="stretch", key="pdf_preview_btn"):
            st.session_state["pdf_preview_open"] = not preview_open
            st.rerun()
        if preview_open:
            with st.container(border=True):
                with st.spinner("미리보기 만드는 중..."):
                    pages = _pdf_page_images(pdf_bytes)
                st.caption(f"PDF 미리보기 · {len(pages)}쪽 — 위에서 체크·입력을 바꾸면 바로 반영됩니다.")
                _, mid, _ = st.columns([1, 6, 1])
                for img in pages:
                    mid.image(img, width="stretch")

# ------------------------------------------------------------------
# 탭 2: 월간 PR리포트
# ------------------------------------------------------------------
with tab2:
    st.caption(
        "네이버에서 한 달 기사를 모아 앱에서 확인·수정하고, 고객사에 보낼 PDF를 바로 만듭니다. "
        "숫자(형식별·보도자료별·관련 기사 수)는 아래 표에서 자동으로 계산됩니다."
    )

    m_work = st.file_uploader(
        "이전 작업 불러오기 (아래 '💾 작업 저장'으로 받아둔 .json 파일을 올리면 수집한 기사와 수정 내용이 그대로 돌아옵니다)",
        type=["json"],
        key="m_work_file",
    )
    if m_work is not None and st.session_state.get("m_work_id") != m_work.file_id:
        try:
            w = load_month_work(m_work.getvalue())
        except ValueError as e:
            st.error(str(e))
        else:
            st.session_state["m_work_id"] = m_work.file_id
            st.session_state["monthly"] = {
                **w["meta"], "articles": w["articles"], "releases": w["releases"], "meetings": w["meetings"],
                "plans": w["plans"], "summary": w["summary"], "trend_rows": w["trend_rows"],
            }
            st.session_state["m_run"] = st.session_state.get("m_run", 0) + 1
            st.session_state["m_cid"] = st.session_state.get("m_cid", 0) + 1

    m_prev = st.file_uploader(
        "지난달 PDF 리포트 (월별 추이·전월 비교용) · 이 앱으로 만든 PDF에는 월별 건수가 들어 있어서 올리기만 하면 이어집니다. "
        "처음 한 번은 기존 PR리포트 엑셀을 올려도 됩니다",
        type=["pdf", "xlsx"],
        key="m_prev",
    )

    _last_month = date.today().replace(day=1) - timedelta(days=1)
    with st.form("monthly_form"):
        c1, c2, c3 = st.columns([2, 1, 1])
        m_keyword = c1.text_input("키워드 (기업명)", value="딥파인", key="m_keyword")
        m_year = c2.number_input("연도", min_value=2000, max_value=2100, value=_last_month.year, step=1, key="m_year")
        m_month = c3.number_input("월", min_value=1, max_value=12, value=_last_month.month, step=1, key="m_month")
        m_auto = st.checkbox("보도자료 자동 분류 (기사 본문을 확인해서 1~2분 걸립니다)", value=True, key="m_auto")
        m_go = st.form_submit_button("기사 수집", type="primary")

    if m_go:
        if not m_keyword:
            st.error("키워드를 입력하세요.")
        elif not NAVER_CLIENT_ID:
            st.error("네이버 API 키가 없어 진행할 수 없습니다.")
        else:
            with st.spinner(f"네이버에서 '{m_keyword}' {int(m_month)}월 기사 검색 중..."):
                try:
                    items = fetch_month_articles(m_keyword, int(m_year), int(m_month), NAVER_CLIENT_ID, NAVER_CLIENT_SECRET)
                except Exception as e:
                    st.error(f"뉴스 수집 실패: {e}")
                    items = None
            if items is not None:
                auto = {}
                if m_auto and items:
                    bar = st.progress(0.0, text="보도자료 자동 분류 중...")
                    with ThreadPoolExecutor(max_workers=6) as pool:
                        futs = {pool.submit(classify_row, clean_html(it["title"]), it.get("originallink") or it.get("link"),
                                            m_keyword): it for it in items}
                        for n, fut in enumerate(as_completed(futs), 1):
                            it = futs[fut]
                            try:
                                auto[id(it)] = fut.result()[0]
                            except Exception:
                                auto[id(it)] = None
                            bar.progress(n / len(items), text=f"보도자료 자동 분류 [{n}/{len(items)}]")
                    bar.empty()
                arts = []
                for it in items:
                    t = clean_html(it["title"])
                    arts.append({
                        "url": it.get("originallink") or it.get("link"), "media": it["_media"], "title": t,
                        "date": it["_dt"].date(), "kind": initial_kind(t, m_keyword, auto.get(id(it))),
                        "include": True, "group": "",
                    })
                m_rels = guess_releases(arts, m_keyword)
                n_rel = auto_mark_releases(arts, m_rels, m_keyword)  # 제목에 키워드 없는 우리 보도자료 기사
                n_grp = auto_group_mentions(arts)  # 단순언급 중 같은 소식끼리 묶음
                st.session_state["monthly"] = {
                    "company": m_keyword, "keyword": m_keyword, "year": int(m_year), "month": int(m_month),
                    "articles": arts, "releases": m_rels, "meetings": [], "plans": [],
                    "summary": None, "trend_rows": history_from_secrets(st.secrets, m_keyword, int(m_year)),
                    "auto_note": (n_rel, n_grp),
                }
                st.session_state["m_run"] = st.session_state.get("m_run", 0) + 1
                st.session_state["m_cid"] = st.session_state.get("m_cid", 0) + 1

    M = st.session_state.get("monthly")
    if M:
        run = st.session_state["m_run"]
        # 지난달 PDF(숨은 월별 기록) 또는 처음 한 번 PR리포트 엑셀(Summary)로 월별 건수 채우기 - 수집·불러오기마다 한 번
        prev_key = (m_prev.file_id, st.session_state.get("m_cid", 0)) if m_prev is not None else None
        if prev_key and st.session_state.get("m_prev_applied") != prev_key:
            st.session_state["m_prev_applied"] = prev_key
            raw = m_prev.getvalue()
            if m_prev.name.lower().endswith(".pdf"):
                h = read_history(raw)
                year, rows = (h["year"], h["rows"]) if h else (None, None)
                if not h:
                    msg = ("warning", "이 PDF에는 월별 기록이 없습니다 (이 앱으로 만든 PDF가 아님). "
                                      "아래 ⑤ 월별 건수 표에 직접 넣거나 기존 PR리포트 엑셀을 한 번 올려주세요.")
            else:
                year, rows = trend_rows_from_xlsx(raw)
                if not rows:
                    msg = ("warning", "엑셀에서 Summary 시트(월별 커버리지)를 찾지 못했습니다.")
            if rows:
                if year not in (None, M["year"]):
                    msg = ("warning", f"올린 파일은 {year}년 기록이라 {M['year']}년 리포트에는 쓰지 않았습니다.")
                else:
                    M["trend_rows"] = rows
                    st.session_state["m_hist_ver"] = st.session_state.get("m_hist_ver", 0) + 1
                    msg = ("info", f"'{m_prev.name}'에서 월별 건수를 불러왔습니다. 아래 ⑤ 표에서 확인할 수 있습니다.")
            st.session_state["m_prev_msg"] = msg
        if m_prev is None:
            st.session_state.pop("m_prev_msg", None)
        kind_, text_ = st.session_state.get("m_prev_msg", (None, None))
        if kind_:
            getattr(st, kind_)(text_)
        elif M["trend_rows"] and M["trend_rows"] == history_from_secrets(st.secrets, M["company"], M["year"]):
            st.info("앱에 저장해 둔 월별 건수(Secrets)를 불러왔습니다. 지난달 PDF를 올리면 그 기록으로 바뀝니다.")
        st.success(f"{M['company']} {M['year']}년 {M['month']}월 · 기사 {len(M['articles'])}건")
        n_rel, n_grp = M.get("auto_note") or (0, 0)
        if n_rel or n_grp:
            st.caption(
                "자동 정리: " + " · ".join(x for x in (
                    f"제목에 기업명이 없지만 보도자료 기사와 같은 소식인 {n_rel}건을 보도자료로" if n_rel else "",
                    f"단순언급 중 같은 소식인 {n_grp}건에 같은 묶음 이름을" if n_grp else "",
                ) if x) + " 넣었습니다. 아래 표에서 확인하세요."
            )
        company_name = st.text_input("PDF에 표시할 고객사명", value=M["company"], key=f"m_company_{run}")

        # ① 보도자료
        st.markdown("##### ① 이번 달 배포한 보도자료")
        st.caption("보도자료 기사가 몰린 날짜로 미리 채웠습니다. 제목·배포일을 확인하고 고치세요. 빠진 건은 추가, 틀린 건은 삭제하세요. "
                   "형식이 '보도자료'인 기사는 배포일 기준으로 각 보도자료에 자동으로 묶입니다.")
        # 표(data_editor) 칸에서는 한글 입력이 Enter 때 날아가는 일이 있어서, 보도자료는 줄마다 일반 입력칸으로 받음
        rows_key = f"m_rel_rows_{run}"
        if rows_key not in st.session_state:
            st.session_state[rows_key] = [{"id": i, "title": r["title"], "date": r["date"]}
                                          for i, r in enumerate(M["releases"])]
        rel_rows = st.session_state[rows_key]
        for r in rel_rows:
            c1, c2, c3 = st.columns([7, 2, 1], vertical_alignment="bottom")
            r["title"] = c1.text_input("보도자료 제목", value=r["title"], key=f"m_rel_t_{run}_{r['id']}")
            r["date"] = c2.date_input("배포일", value=r["date"], format="YYYY-MM-DD", key=f"m_rel_d_{run}_{r['id']}")
            if c3.button("삭제", key=f"m_rel_x_{run}_{r['id']}", width="stretch"):
                st.session_state[rows_key] = [x for x in rel_rows if x["id"] != r["id"]]
                st.rerun()
        if st.button("➕ 보도자료 추가", key=f"m_rel_add_{run}"):
            rel_rows.append({"id": max((x["id"] for x in rel_rows), default=-1) + 1, "title": "", "date": None})
            st.rerun()
        releases = [{"title": r["title"].strip(), "date": r["date"]} for r in rel_rows if r["title"].strip()]
        releases.sort(key=lambda r: r["date"] or date.max)
        if any(r["date"] is None for r in releases):
            st.warning("배포일이 비어 있는 보도자료가 있습니다. 배포일을 넣어야 그 보도자료 기사가 묶여서 게재 수가 셉니다.")

        # ② 기사
        st.markdown("##### ② 기사 확인")
        st.caption(
            "**형식**을 확인하세요. 자동 분류는 보도자료만 잡고(제목이 보도자료 기사와 같은 소식이면 보도자료), "
            "제목에 키워드가 없으면 단순언급으로 넣었습니다. 형식이 빈 기사는 PDF에서 빠집니다. "
            "단순언급은 제목이 비슷한 같은 소식끼리 **묶음** 이름을 자동으로 넣었고, 같은 이름끼리 하나로 묶여 "
            "대표 기사 1건 + '관련 기사 N건'으로 나옵니다. 묶음은 직접 고치거나 지워도 됩니다. "
            "보고서에서 뺄 기사는 **포함**을 해제하세요."
        )
        art_df = st.data_editor(
            pd.DataFrame(
                [{"포함": a["include"], "형식": a["kind"] or None, "묶음": a.get("group") or "",
                  "매체": a["media"], "제목": a["title"], "게재일": a["date"], "URL": a["url"]} for a in M["articles"]],
                columns=["포함", "형식", "묶음", "매체", "제목", "게재일", "URL"],
            ).astype({"묶음": "str", "매체": "str", "제목": "str"}),
            hide_index=True, width="stretch", height=min(38 + 35 * max(len(M["articles"]), 1), 560),
            disabled=["URL"], key=f"m_arts_{run}",
            column_config={
                "포함": st.column_config.CheckboxColumn("포함", width="small"),
                "형식": st.column_config.SelectboxColumn("형식", options=KIND_OPTIONS, width="small"),
                "묶음": st.column_config.TextColumn("묶음 (단순언급)", width="small"),
                "매체": st.column_config.TextColumn("매체", width="small"),
                "제목": st.column_config.TextColumn("제목", width="large"),
                "게재일": st.column_config.DateColumn("게재일", format="MM-DD", width="small"),
                "URL": st.column_config.LinkColumn("원문", display_text="열기", width="small"),
            },
        )
        articles = []
        for (_, r), a in zip(art_df.iterrows(), M["articles"]):
            text = lambda k, fb: fb if pd.isna(r[k]) else str(r[k]).strip()
            d = r["게재일"]
            articles.append({
                "url": a["url"], "include": bool(r["포함"]),
                "kind": "" if pd.isna(r["형식"]) or r["형식"] is None else str(r["형식"]),
                "group": text("묶음", ""), "media": text("매체", a["media"]), "title": text("제목", a["title"]),
                "date": d if isinstance(d, date) and not pd.isna(d) else a["date"],
            })

        # ②-1 네이버에 없는 기사 (접지 않고 항상 보이게)
        st.markdown("##### ②-1 네이버에 없는 기사 추가")
        st.caption("다음·구글 등에서 찾은 기사 링크를 한 줄에 하나씩 붙여 넣고 버튼을 누르면 매체·제목·게재일을 읽어 위 표에 추가합니다. "
                   "형식·단순언급 묶음도 같은 기준으로 자동으로 정리됩니다.")
        add_text = st.text_area("기사 링크 (한 줄에 하나씩)", key=f"m_add_urls_{run}", height=100,
                                placeholder="https://www.example.co.kr/news/...")
        add_go = st.button("➕ 기사 정보 불러와서 추가", key=f"m_add_btn_{run}")

        def _row_inputs(name, init_rows, fields, add_label):
            """줄마다 일반 입력칸 + 삭제 버튼, 아래 추가 버튼 (표 칸은 한글 입력이 저장되지 않는 경우가 있어서).
            fields: [(이름, 너비, 예시)]. 반환: 내용이 있는 줄의 (값, ...) 목록."""
            key = f"{name}_{run}"
            if key not in st.session_state:
                st.session_state[key] = [{"id": i, "v": [str(x) for x in r]} for i, r in enumerate(init_rows)]
            rows = st.session_state[key]
            for n, r in enumerate(rows):
                cols = st.columns([w for _, w, _ in fields] + [1], vertical_alignment="bottom")
                for j, (label, _, hint) in enumerate(fields):
                    r["v"][j] = cols[j].text_input(label, value=r["v"][j], placeholder=hint,
                                                   key=f"{key}_{r['id']}_{j}",
                                                   label_visibility="visible" if n == 0 else "collapsed")
                if cols[-1].button("삭제", key=f"{key}_{r['id']}_x", width="stretch"):
                    st.session_state[key] = [x for x in rows if x["id"] != r["id"]]
                    st.rerun()
            if st.button(add_label, key=f"{key}_add"):
                rows.append({"id": max((x["id"] for x in rows), default=-1) + 1, "v": [""] * len(fields)})
                st.rerun()
            return [tuple(v.strip() for v in r["v"]) for r in rows if any(v.strip() for v in r["v"])]

        # ③ 기자 미팅
        st.markdown("##### ③ 기자 미팅")
        meetings = _row_inputs("m_meet", M["meetings"], [("매체", 2, "바이라인네트워크"), ("기자", 1.5, "홍길동"),
                                                         ("주요 내용", 5, "회사 소개, 인터뷰 피칭")], "➕ 미팅 추가")
        # ④ 다음 달 목표
        nm = M["month"] % 12 + 1
        st.markdown(f"##### ④ {nm}월 PR 목표")
        plans = _row_inputs("m_plan", M["plans"], [("형식", 1.5, "보도자료"), ("내용", 5, "한국관광공사 오픈 이노베이션"),
                                                   ("진행 상황", 2, "10/12 배포 예정")], "➕ 목표 추가")

        # ⑤ 월별 건수 (지난달까지) - 지난달 PDF/엑셀에서 불러온 값, 직접 고칠 수 있음. 이번 달은 기사 표에서 셈
        st.markdown("##### ⑤ 월별 건수 (월별 추이·전월 비교)")
        hist_rows = {k: list((M["trend_rows"] or {}).get(k) or [0] * 12) for k in TYPE_ORDER}
        if M["month"] > 1:
            with st.expander(
                f"1~{M['month'] - 1}월 형식별 건수 보기·고치기"
                + ("" if M["trend_rows"] else " · 지난달 PDF를 올리거나 여기에 직접 넣으면 PDF에 추이·전월 비교가 나옵니다"),
                expanded=not M["trend_rows"],
            ):
                hist_df = st.data_editor(
                    pd.DataFrame([{"월": f"{i + 1}월", **{k: int(hist_rows[k][i]) for k in TYPE_ORDER}}
                                  for i in range(M["month"] - 1)]),
                    hide_index=True, width="stretch", disabled=["월"],
                    key=f"m_hist_{run}_{st.session_state.get('m_hist_ver', 0)}",
                    column_config={k: st.column_config.NumberColumn(k, min_value=0, step=1) for k in TYPE_ORDER},
                )
                st.caption(f"{M['month']}월 숫자는 위 기사 표에서 자동으로 셉니다. "
                           "이 숫자들은 PDF에 함께 저장돼서, 다음 달엔 이번 PDF만 올리면 이어집니다.")
            for i, (_, r) in enumerate(hist_df.iterrows()):
                for k in TYPE_ORDER:
                    hist_rows[k][i] = 0 if pd.isna(r[k]) else int(r[k])
        has_hist = any(hist_rows[k][i] for k in TYPE_ORDER for i in range(M["month"] - 1))
        trend_input = hist_rows if has_hist else None

        meta = {"company": company_name.strip() or M["company"], "year": M["year"], "month": M["month"]}
        mdata, rel_dates, n_blank = compose(meta, articles, releases, meetings, plans, trend_input)

        # ⑤ 요약
        st.markdown("##### ⑥ 이달의 요약")
        s_key, auto_key = f"m_summary_{run}", f"m_summary_auto_{run}"
        auto_text = default_summary(mdata, rel_dates)
        # 아직 손대지 않았으면(자동 문장 그대로면) 표를 고칠 때마다 자동 문장도 새로 맞춤. 직접 고친 뒤에는 건드리지 않음
        if s_key not in st.session_state:
            st.session_state[s_key] = M["summary"] if M["summary"] is not None else auto_text
        elif st.session_state[s_key] == st.session_state.get(auto_key):
            st.session_state[s_key] = auto_text
        st.session_state[auto_key] = auto_text
        if st.button("↺ 자동 문장으로 다시 채우기", key=f"m_sum_reset_{run}",
                     help="표 내용(보도자료·매체 유형·단순 언급 묶음·전월 대비)으로 만든 문장으로 되돌립니다."):
            st.session_state[s_key] = auto_text
            st.rerun()
        summary_text = st.text_area(
            "PDF 1쪽에 들어갑니다 · 표 내용으로 자동으로 쓴 초안이니 자유롭게 고치세요 (비우면 빠짐)", height=130, key=s_key)

        if add_go:  # 지금까지 고친 내용을 저장한 뒤 기사를 붙이고 표를 다시 그림
            urls = [u.rstrip(".,)") for u in re.findall(r"https?://[^\s<>\"']+", add_text or "")]
            have = {normalize_url(a["url"]) for a in articles}
            new = []
            with st.spinner(f"기사 {len(urls)}건 정보 불러오는 중..."):
                domain_map = load_domain_map()
                for u in urls:
                    if normalize_url(u) in have:
                        continue
                    meta_a = fetch_article_meta(u, domain_map)
                    pub = meta_a["게재일자"].date() if meta_a["게재일자"] else date(M["year"], M["month"], 1)
                    t_new = meta_a["제목"] or u
                    new.append({"url": meta_a["URL"], "media": meta_a["매체명"], "title": t_new, "date": pub,
                                "kind": initial_kind(t_new, M.get("keyword") or M["company"], None),
                                "include": True, "group": ""})
                    have.add(normalize_url(u))
            merged = articles + new
            # 새로 넣은 기사도 같은 기준으로: 우리 보도자료와 같은 소식이면 보도자료, 단순언급은 같은 소식끼리 묶음
            n_rel = auto_mark_releases(merged, releases, M.get("keyword") or M["company"], only={a["url"] for a in new})
            n_grp = auto_group_mentions(merged)
            M.update(articles=sort_articles(merged), releases=releases, meetings=meetings, plans=plans,
                     summary=summary_text, company=meta["company"], trend_rows=trend_input, auto_note=(n_rel, n_grp))
            st.session_state["m_run"] = run + 1
            st.rerun()

        # 집계 미리보기
        t = mdata["types"]
        st.caption(
            "집계: " + " · ".join(f"{k} {t[k]}" for k in t if t[k]) + f" · 합계 **{sum(t.values())}건**"
            + (" | 보도자료별: " + " · ".join(f"{i + 1}) {r['count']}건" for i, r in enumerate(mdata["releases"]))
               if mdata["releases"] else "")
            + f" | 단순 언급 주제 {len(mdata['mentions'])}개"
        )
        if n_blank:
            st.warning(f"형식이 비어 있는 기사 {n_blank}건은 PDF에서 빠집니다. 표에서 형식을 골라주세요.")
        no_release = sum(1 for a in articles if a["include"] and a["kind"] == "보도자료"
                         and assign_release(a, rel_dates) is None)
        if no_release:
            why = "① 보도자료 목록이 비어 있어" if not releases else "배포일보다 먼저 실린 기사라"
            st.warning(f"형식이 보도자료인데 {why} 어느 보도자료에도 묶이지 않은 기사 {no_release}건이 있습니다 "
                       "(PDF에서 '기획·기타'로 나갑니다). 보도자료 목록이나 기사 형식을 확인하세요.")

        try:
            m_pdf = build_monthly_pdf(mdata, rel_dates, summary_text)
        except Exception as e:
            st.error(f"PDF를 만들지 못했습니다: {e}")
            m_pdf = None
        fname = f"{M['year']}년_{M['month']}월_{meta['company']}_PR리포트"
        b1, b2, b3 = st.columns(3)
        if m_pdf:
            b1.download_button("📑 PDF 리포트 다운로드", data=m_pdf, file_name=f"{fname}.pdf", mime="application/pdf",
                               type="primary", width="stretch", on_click="ignore", key="m_dl_pdf",
                               help="다음 달에 이 PDF를 '지난달 PDF 리포트'에 올리면 월별 추이·전월 비교가 이어집니다.")
        b2.download_button(
            "💾 작업 저장 (중간 저장용)",
            data=dump_month_work(meta, articles, releases, meetings, plans, summary_text, trend_input),
            file_name=f"월간작업_{meta['company']}_{M['year']}{M['month']:02d}.json", mime="application/json",
            width="stretch", on_click="ignore", key="m_dl_work",
            help="작업을 나중에 이어서 하려면 이 파일을 받아두었다가 맨 위 '이전 작업 불러오기'에 올리세요.",
        )
        m_preview_open = st.session_state.get("m_pdf_preview_open", False)
        if b3.button("🙈 미리보기 닫기" if m_preview_open else "👁 PDF 미리보기", width="stretch",
                     key="m_pdf_preview_btn", disabled=not m_pdf):
            st.session_state["m_pdf_preview_open"] = not m_preview_open
            st.rerun()
        if m_preview_open and m_pdf:
            with st.container(border=True):
                with st.spinner("미리보기 만드는 중..."):
                    m_pages = _pdf_page_images(m_pdf)
                st.caption(f"PDF 미리보기 · {len(m_pages)}쪽 — 위에서 표·입력을 바꾸면 바로 반영됩니다.")
                _, m_mid, _ = st.columns([1, 6, 1])
                for img in m_pages:
                    m_mid.image(img, width="stretch")

# ------------------------------------------------------------------
# 탭 3: 네이버 기사 수집
# ------------------------------------------------------------------
with tab3:
    st.caption("지정한 기간에 게재된, 키워드가 포함된 네이버 기사를 모두 모아 엑셀로 만듭니다. 리포트 양식 없이 목록만 뽑습니다.")

    with st.form("daily_form"):
        date_range = st.date_input(
            "기간",
            value=(date.today(), date.today()),
            key="daily_date_range",
        )
        keywords_text = st.text_area(
            "키워드 (한 단어로 입력)",
            placeholder="예)딥파인",
            key="daily_keywords",
        )
        submitted3 = st.form_submit_button("기사 수집", type="primary")

    if submitted3:
        keywords = [k.strip() for k in keywords_text.splitlines() if k.strip()]
        if isinstance(date_range, (list, tuple)):
            start_date = date_range[0]
            end_date = date_range[1] if len(date_range) > 1 else date_range[0]
        else:
            start_date = end_date = date_range

        if not keywords:
            st.error("키워드를 최소 1개 입력하세요.")
        elif start_date > end_date:
            st.error("시작일이 종료일보다 늦을 수 없습니다.")
        elif not NAVER_CLIENT_ID:
            st.error("네이버 API 키가 없어 진행할 수 없습니다.")
        else:
            start_str = start_date.strftime("%Y-%m-%d")
            window_days = (end_date - start_date).days
            keyword_coverage = []
            with st.spinner(f"네이버에서 {len(keywords)}개 키워드 검색 중..."):
                for kw in keywords:
                    try:
                        cov = collect_coverage(
                            kw, start_str, window_days=window_days, require_text=kw, verify_relevance=False
                        )
                    except Exception as e:
                        st.error(f"'{kw}' 검색 실패: {e}")
                        cov = []
                    keyword_coverage.append((kw, cov))

            total = sum(len(cov) for _, cov in keyword_coverage)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            if start_date == end_date:
                sheet_name = start_date.strftime("%y%m%d")
                date_str = start_date.strftime("%Y%m%d")
            else:
                sheet_name = f"{start_date.strftime('%y%m%d')}-{end_date.strftime('%y%m%d')}"
                date_str = f"{start_date.strftime('%Y%m%d')}-{end_date.strftime('%Y%m%d')}"

            with st.spinner("고정 목록에 없는 매체는 기사 페이지에서 매체명 자동 인식 중..."):
                unmapped, auto_detected = build_daily_sheet(wb, sheet_name, keyword_coverage)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            filename = f"네이버기사수집_{date_str}.xlsx"

            st.success(f"완료! 총 {total}건 수집됨.")
            for kw, cov in keyword_coverage:
                st.caption(f"- {kw}: {len(cov)}건")
            if auto_detected:
                lines = "\n".join(f"- {d} → {n}" for d, n in sorted(auto_detected.items()))
                st.info(
                    "다음 매체는 고정 목록에 없어서 기사 페이지에서 자동으로 인식했습니다 "
                    "(맞는지 확인해보고, domain_to_media.json에 추가해두면 다음부턴 더 빨라져요):\n"
                    + lines
                )
            if unmapped:
                st.warning(
                    "매체명을 자동으로도 못 찾은 도메인: " + ", ".join(sorted(unmapped))
                    + " — 직접 확인해서 domain_to_media.json에 추가해주세요."
                )
            st.download_button(
                "⬇️ 엑셀 다운로드",
                data=buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="daily_download",
            )
