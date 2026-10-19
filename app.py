# -*- coding: utf-8 -*-
"""
PR 리포트 도구 웹앱.
- 탭1: 결과보고서 (보도자료 1건 배포 결과 - 네이버 뉴스 수집 후 엑셀 시트 + PDF 보고서 생성)
- 탭2: 월간 PR리포트 (기존 워크북에 새 달 시트 추가 + 기사형식 자동분류, 완성된 엑셀 -> PDF)

네이버 API 키는 Streamlit Cloud의 App settings > Secrets 에
NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 로 등록해두면 자동으로 읽힘.
"""
import io
import os
import re
from datetime import date, datetime

import openpyxl
import pandas as pd
import streamlit as st

from naver_news import collect_coverage
from report_sheet import (
    EXTRA_PORTALS, LOW_RELEVANCE, apply_manual_portals, build_sheet, derive_company, derive_keyword,
    load_domain_map, resolve_coverage, title_relevance,
)
from report_pdf import build_pdf
from article_meta import fetch_article_meta, normalize_url
from work_file import dump_work, load_work, restore_choices
from pr_report_core import fetch_month_articles, build_month_tab, classify_report
from daily_collect_sheet import build_daily_sheet
from monthly_pdf import month_sheets, read_month_report, guess_release_dates, default_summary, build_monthly_pdf

REPORT_WINDOW_DAYS = 3  # 배포일 이후 며칠까지 검색할지 (고정)

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
            company = derive_company(title)
            keyword = derive_keyword(title)
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
                # 다음·구글 수동 입력으로 화면이 다시 그려져도(rerun) 결과가 유지되도록 세션에 보관
                st.session_state["report"] = {
                    "title": title,
                    "company": company,
                    "keyword": keyword,
                    "dist_date": distribution_date,
                    "rows": rows,
                    "unmapped": unmapped,
                    "auto_detected": auto_detected,
                    "existing": existing_file.getvalue() if existing_file is not None else None,
                }
                run = st.session_state.get("report_run", 0) + 1
                st.session_state["report_run"] = run
                saved_work = st.session_state.get("saved_work")
                if saved_work and saved_work["title"] == title:
                    excluded, portals, manual_saved = restore_choices(saved_work, rows)
                    st.session_state["report"]["restore"] = {"excluded": excluded, "portals": portals}
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
            "보고서에서 뺄 기사는 **포함** 체크를 해제하세요. "
            "다음·구글은 자동 수집이 안 되니 직접 확인한 노출을 체크하면 PDF·엑셀에 바로 반영됩니다."
        )

        def _check_note(r):
            if r.get("확인필요"):
                return "⚠ 본문 확인 필요"
            if title_relevance(title, company, r["제목"]) < LOW_RELEVANCE:
                return "⚠ 제목이 보도자료와 다름"
            return ""

        notes = [_check_note(r) for r in rows]
        restore = report.get("restore") or {"excluded": set(), "portals": {}}
        if report.get("restore"):
            st.info("작업 파일의 선택(포함 해제·다음·구글 노출·직접 추가한 기사)을 적용했습니다. 새로 검색된 기사는 그대로 추가됐습니다.")
        order = range(len(rows))  # 보고서와 같은 순서 (포털 > 매체 영향력)
        flags = st.data_editor(
            pd.DataFrame(
                [{"포함": rows[i]["URL"] not in restore["excluded"], "확인": notes[i],
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
            matched, fetched = {}, []
            with st.spinner(f"기사 {len(portals_by_url)}건 정보 불러오는 중..."):
                domain_map = load_domain_map()
                for u, ps in portals_by_url.items():
                    naver_url = naver_by_url.get(normalize_url(u))
                    if naver_url:
                        matched.setdefault(naver_url, set()).update(ps)
                        continue
                    meta = fetch_article_meta(u, domain_map)
                    meta["포털"] = ps
                    fetched.append(meta)
            prev_n = st.session_state.get(fetched_key, {}).get("n", 0)
            st.session_state[fetched_key] = {"matched": matched, "fetched": fetched, "n": prev_n + 1}
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
        manual, skipped = [], 0
        for _, m in manual_df.iterrows():
            media, art_title, url = (
                "" if pd.isna(m.get(k)) else str(m.get(k)).strip() for k in ("매체명", "제목", "URL")
            )
            portals = [p for p in EXTRA_PORTALS if m.get(p) == True]  # noqa: E712 (NaN/None 은 False)
            if not (media or art_title or url):
                continue  # 빈 행
            if not (media and art_title and url and portals):
                skipped += 1
                continue
            pub = m.get("게재일")
            pub = pub if isinstance(pub, date) and not pd.isna(pub) else dist_date
            manual.append({
                "매체명": media, "제목": art_title, "URL": url,
                "게재일자": datetime.combine(pub, datetime.min.time()), "포털": portals,
            })
        if skipped:
            st.warning(f"매체명·제목·URL·포털(다음/구글) 중 빈 칸이 있는 {skipped}개 행은 보고서에서 빠졌습니다.")

        final_rows = apply_manual_portals(included, extra, manual)

        if report["existing"] is not None:
            wb = openpyxl.load_workbook(io.BytesIO(report["existing"]))
        else:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        build_sheet(wb, dist_date.strftime("%y%m%d"), title, company, dist_str, final_rows)
        wb.active = 0  # 새로 추가된(맨 왼쪽) 시트가 열리도록
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        pdf_bytes = build_pdf(title, company, dist_str, final_rows, window_days=REPORT_WINDOW_DAYS)

        n_extra = sum(len(v) > 0 for v in extra.values()) + len(manual)
        if n_extra:
            st.caption(f"다음·구글 반영: {n_extra}건 (보고서 총 {len(final_rows)}건)")

        filename = f"결과보고서_{title}_{date.today():%Y%m%d}.xlsx"
        # on_click="ignore": 한쪽을 받아도 화면이 다시 그려지지 않아서 나머지 버튼이 유지됨
        d1, d2, d3 = st.columns(3)
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
            data=dump_work(title, dist_date, rows, included, extra, manual),
            file_name=f"작업_{company}_{dist_date:%y%m%d}.json",
            mime="application/json",
            width="stretch",
            on_click="ignore",
            help="내일 이어서 하려면 이 파일을 받아두었다가 맨 위 '이전 작업 불러오기'에 올리세요.",
        )

# ------------------------------------------------------------------
# 탭 2: 월간 PR리포트
# ------------------------------------------------------------------
with tab2:
    st.caption("기존 파일에 새 달 시트를 추가하고, 기사형식이 보도자료인 것만 자동분류합니다.")

    template_file = st.file_uploader(
        "기존 PR리포트 파일 업로드", type=["xlsx"], key="pr_template"
    )

    sheet_options = []
    if template_file is not None:
        try:
            template_file.seek(0)
            _preview_wb = openpyxl.load_workbook(template_file, read_only=True)
            sheet_options = [s for s in _preview_wb.sheetnames if s.endswith("월")]
            _preview_wb.close()
        except Exception as e:
            st.error(f"워크북을 열 수 없습니다: {e}")

    keyword2 = st.text_input("키워드", value="딥파인", key="pr_keyword")
    c1, c2 = st.columns(2)
    year = c1.number_input("연도", min_value=2000, max_value=2100, value=date.today().year, step=1, key="pr_year")
    month = c2.number_input("새로 만들 달", min_value=1, max_value=12, value=date.today().month, step=1, key="pr_month")

    if sheet_options:
        template_sheet_name = st.selectbox(
            "서식을 복제할 기존 월 시트 (가장 최근 달이 기본 선택됨)", sheet_options, index=0
        )
    else:
        template_sheet_name = st.text_input("서식을 복제할 기존 월 시트 이름 (예: 6월)", key="pr_sheet_manual")

    do_classify = st.checkbox("기사형식 자동 분류도 실행", value=True, key="pr_classify")

    if st.button("PR리포트 생성", type="primary", key="pr_submit"):
        if template_file is None:
            st.error("기존 워크북을 업로드하세요.")
        elif not template_sheet_name:
            st.error("서식을 복제할 월 시트 이름을 지정하세요.")
        elif not keyword2:
            st.error("키워드를 입력하세요.")
        elif not NAVER_CLIENT_ID:
            st.error("네이버 API 키가 없어 진행할 수 없습니다.")
        else:
            template_file.seek(0)
            wb = openpyxl.load_workbook(template_file)

            if template_sheet_name not in wb.sheetnames:
                st.error(f"'{template_sheet_name}' 시트를 찾을 수 없습니다. 현재 시트: {wb.sheetnames}")
            else:
                with st.spinner(f"네이버에서 '{keyword2}' 검색 중..."):
                    try:
                        articles = fetch_month_articles(
                            keyword2, int(year), int(month), NAVER_CLIENT_ID, NAVER_CLIENT_SECRET
                        )
                    except Exception as e:
                        st.error(f"뉴스 수집 실패: {e}")
                        articles = None

                if articles is not None:
                    ws, new_count = build_month_tab(wb, template_sheet_name, int(month), int(year), keyword2, articles)
                    sheet_name = f"{int(month)}월"
                    st.success(f"'{sheet_name}' 시트 생성 완료. 3번 섹션에 기사 {new_count}건 채움.")

                    if do_classify and new_count > 0:
                        progress_bar = st.progress(0.0, text="기사형식 분류 준비 중...")

                        def _cb(done, total, t, result):
                            frac = done / total if total else 1.0
                            progress_bar.progress(frac, text=f"[{done}/{total}] {result} - {t[:30]}")

                        filled, blanked = classify_report(wb, sheet_name, keyword2, progress_callback=_cb)
                        progress_bar.progress(1.0, text="분류 완료")
                        st.info(
                            f"확정 보도자료로 채운 행: {filled}건 / 공란으로 남은 행: {blanked}건 "
                            "(공란은 인터뷰/기고/기획(오픈피알)/기획(외부)/단순언급 중 직접 확인해서 채워주세요)"
                        )

                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    filename = f"{int(year)}년_{int(month)}월_{keyword2}_PR리포트.xlsx"

                    st.download_button(
                        "⬇️ 엑셀 다운로드",
                        data=buf,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.caption(
                        "1/2/4/5/6번 섹션은 지난달 데이터가 지워진 빈 템플릿 상태이니 직접 채워주세요. "
                        "다음 달 작업을 위해 이번에 만든 파일을 템플릿으로 보관해두세요."
                    )

    # --- 다 채운 월간 리포트 엑셀 -> 고객사 전달용 PDF ---
    st.divider()
    st.subheader("📄 월간 리포트 PDF 만들기")
    st.caption("1~6번 섹션을 다 채운 PR리포트 엑셀을 올리면, 고객사에 보낼 PDF로 만들어 줍니다.")

    report_file = st.file_uploader("완성된 PR리포트 파일 업로드", type=["xlsx"], key="mpdf_file")
    if report_file is not None:
        try:
            mpdf_sheets = month_sheets(report_file)
        except Exception as e:
            st.error(f"워크북을 열 수 없습니다: {e}")
            mpdf_sheets = []
        if not mpdf_sheets:
            st.error("'8월'처럼 'N월' 이름의 시트를 찾지 못했습니다.")
        else:
            mpdf_sheet = st.selectbox("PDF로 만들 달", mpdf_sheets, index=0, key="mpdf_sheet")
            try:
                mdata = read_month_report(report_file, mpdf_sheet)
            except Exception as e:
                st.error(f"'{mpdf_sheet}' 시트를 읽지 못했습니다: {e}")
                mdata = None

            if mdata is not None:
                st.caption(
                    f"{mdata['company']} · 총 {sum(mdata['types'].values())}건 · "
                    f"보도자료 {len(mdata['releases'])}건 · 주요 기사 {len(mdata['articles'])}건 · "
                    f"단순 언급 주제 {len(mdata['mentions'])}개"
                )
                release_dates = []
                if mdata["releases"]:
                    st.markdown("**보도자료 배포일** · 기사 목록을 보도자료별로 묶는 데 씁니다 (기사가 몰린 날짜로 미리 채움)")
                    guessed = guess_release_dates(mdata)
                    for i, (rel, d0) in enumerate(zip(mdata["releases"], guessed)):
                        release_dates.append(st.date_input(
                            f"{i + 1}. {rel['title']} ({rel['count']}건)", value=d0,
                            key=f"mpdf_date_{mpdf_sheet}_{i}",
                        ))
                summary_text = st.text_area(
                    "이달의 요약 (PDF 1쪽에 들어갑니다. 비우면 빠짐)",
                    value=default_summary(mdata), height=110, key=f"mpdf_summary_{mpdf_sheet}",
                )
                try:
                    mpdf_bytes = build_monthly_pdf(mdata, release_dates, summary_text)
                except Exception as e:
                    st.error(f"PDF를 만들지 못했습니다: {e}")
                    mpdf_bytes = None
                if mpdf_bytes:
                    st.download_button(
                        "⬇️ 월간 리포트 PDF 다운로드",
                        data=mpdf_bytes,
                        file_name=f"{mdata['year']}년_{mdata['month']}월_{mdata['company']}_PR리포트.pdf",
                        mime="application/pdf",
                        type="primary",
                        key="mpdf_download",
                        on_click="ignore",
                    )

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
