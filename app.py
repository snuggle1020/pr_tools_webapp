# -*- coding: utf-8 -*-
"""
PR 리포트 도구 웹앱.
- 탭1: 결과보고서 (보도자료 1건 배포 결과 - 네이버 뉴스 수집 후 엑셀 시트 생성)
- 탭2: 월간 PR리포트 (기존 워크북에 새 달 시트 추가 + 기사형식 자동분류)

네이버 API 키는 Streamlit Cloud의 App settings > Secrets 에
NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 로 등록해두면 자동으로 읽힘.
"""
import io
import os
from datetime import date

import openpyxl
import streamlit as st

from naver_news import collect_coverage
from report_sheet import build_sheet
from pr_report_core import fetch_month_articles, build_month_tab, classify_report

st.set_page_config(page_title="PR 리포트 도구", page_icon="📰", layout="wide")


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

st.title("📰 PR 리포트 도구")

if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
    st.error(
        "네이버 API 키가 설정되어 있지 않습니다. 관리자에게 문의하거나, "
        "로컬 실행 시 `.streamlit/secrets.toml`에 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 을 등록하세요."
    )

tab1, tab2 = st.tabs(["📄 결과보고서 (보도자료 1건)", "📅 월간 PR리포트"])

# ------------------------------------------------------------------
# 탭 1: 결과보고서
# ------------------------------------------------------------------
with tab1:
    st.caption("보도자료 1건을 배포한 뒤, 배포일 전후 기사를 모아 결과보고서 시트를 만듭니다.")

    with st.form("report_form"):
        title = st.text_input("보도자료 제목", placeholder="예: 딥파인, 국방·산업 현장용 스마트글래스 기반 AI 에이전트 개발 나선다")
        company = st.text_input("기업명", placeholder="예: 딥파인")
        keyword = st.text_input("네이버 뉴스 검색 키워드", placeholder="예: 딥파인 스마트글래스 AI 에이전트")
        distribution_date = st.date_input("배포일", value=date.today())
        window_days = st.number_input("배포일 전후 며칠까지 검색할지", min_value=0, max_value=30, value=3)
        existing_file = st.file_uploader(
            "기존 결과보고서 워크북 (있으면 새 시트를 맨 왼쪽에 추가 / 없으면 새 파일 생성)",
            type=["xlsx"],
            key="report_existing_wb",
        )
        submitted = st.form_submit_button("리포트 생성", type="primary")

    if submitted:
        if not title or not company or not keyword:
            st.error("제목 / 기업명 / 검색 키워드를 모두 입력하세요.")
        elif not NAVER_CLIENT_ID:
            st.error("네이버 API 키가 없어 진행할 수 없습니다.")
        else:
            dist_str = distribution_date.strftime("%Y-%m-%d")
            with st.spinner(f"네이버에서 '{keyword}' 검색 중..."):
                try:
                    coverage = collect_coverage(keyword, dist_str, window_days=int(window_days))
                except Exception as e:
                    st.error(f"뉴스 수집 실패: {e}")
                    coverage = None

            if coverage is not None:
                sheet_name = distribution_date.strftime("%y%m%d")
                if existing_file is not None:
                    wb = openpyxl.load_workbook(existing_file)
                else:
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)

                unmapped = build_sheet(wb, sheet_name, title, company, dist_str, coverage)
                wb.active = 0  # 새로 추가된(맨 왼쪽) 시트가 열리도록

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                request_date_str = date.today().strftime("%Y%m%d")
                filename = f"결과보고서_{title}_{request_date_str}.xlsx"

                st.success(f"완료! 기사 {len(coverage)}건 수집됨 (시트: {sheet_name}).")
                if unmapped:
                    st.warning(
                        "매체명을 자동으로 못 찾은 도메인: " + ", ".join(sorted(unmapped))
                        + " — domain_to_media.json에 추가하면 다음부터 자동 인식됩니다."
                    )
                st.download_button(
                    "⬇️ 엑셀 다운로드",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ------------------------------------------------------------------
# 탭 2: 월간 PR리포트
# ------------------------------------------------------------------
with tab2:
    st.caption("기존 워크북(지난달까지 있는 파일)에 새 달 시트를 추가하고, 기사형식(K열) 자동분류까지 실행합니다.")

    template_file = st.file_uploader(
        "기존 PR리포트 워크북 업로드 (지난달까지 있는 파일)", type=["xlsx"], key="pr_template"
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

    do_classify = st.checkbox("기사형식(K열) 자동 분류도 실행 (기사 수에 따라 다소 시간 걸림)", value=True, key="pr_classify")

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
