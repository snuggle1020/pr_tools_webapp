# PR 리포트 도구 (웹앱)

팀원들이 링크 하나로 바로 쓸 수 있는 결과보고서 / 월간 PR리포트 생성 웹앱.
`결과보고서`, `딥파인_PR리포트` 두 프로젝트의 로직을 그대로 옮겨 Streamlit 웹 폼으로 만든 것.

## 로컬에서 먼저 확인해보기

```
pip install -r requirements.txt
streamlit run app.py
```

`.streamlit/secrets.toml`에 이미 테스트용 네이버 API 키가 들어있어서 바로 실행됨.
(이 파일은 `.gitignore`에 있어서 git에 올라가지 않음)

## 팀원들이 쓸 수 있게 배포하기 (Streamlit Community Cloud, 무료)

1. **GitHub 저장소 만들기**
   - github.com에서 새 저장소 생성 (사내용이면 Private 추천)
   - 이 폴더(`pr_tools_webapp`) 전체를 그 저장소에 push
     - `.streamlit/secrets.toml`은 `.gitignore`로 제외되어 있어서 실수로 올라가지 않음. 확인 필수.

2. **Streamlit Community Cloud 가입/연결**
   - https://share.streamlit.io 접속 → GitHub 계정으로 로그인
   - "New app" → 방금 만든 저장소 / 브랜치 선택 → Main file path에 `app.py` 입력 → Deploy

3. **네이버 API 키 등록 (배포된 앱에서)**
   - 앱 관리 화면 → Settings → Secrets
   - 아래처럼 입력 후 저장 (이 화면 자체가 비공개 저장소라 팀원에게 노출 안 됨)
     ```
     NAVER_CLIENT_ID = "발급받은 Client ID"
     NAVER_CLIENT_SECRET = "발급받은 Client Secret"
     ```
   - 네이버 API 키는 developers.naver.com/apps 에서 무료 발급 (검색 API 신청)

4. **완료** — 앱이 뜨면 `https://xxxx.streamlit.app` 형태의 링크가 생김.
   이 링크를 팀원들에게 공유하면 각자 브라우저에서 바로 사용 가능 (설치 불필요).

## 앱 업데이트하는 법

로직을 고치고 싶으면 이 폴더의 `.py` 파일을 수정하고 GitHub에 다시 push하면
Streamlit Cloud가 자동으로 재배포함 (몇 분 내 반영).

## 폴더 구성

- `app.py` — Streamlit 화면 (탭1: 결과보고서, 탭2: 월간 PR리포트)
- `naver_news.py` — 결과보고서용 네이버 뉴스 검색 (배포일 전후 N일)
- `report_sheet.py` — 결과보고서 엑셀 시트 조립 로직
- `pr_report_core.py` — 월간 PR리포트 엑셀 조립 + 기사형식(K열) 자동분류 로직
- `media_order.py` — 매체명 정렬 우선순위 (통신사>일간지>경제지>IT전문지>산업지>온라인매체)
- `domain_to_media.json` — 결과보고서용 도메인→매체명 매핑
- `.streamlit/secrets.toml` — 로컬 테스트용 API 키 (git에 올라가지 않음)

## 참고

- 네이버 API는 하루 호출 한도가 있음 (개발자센터에서 확인). 팀 전체가 같은 키를 쓰므로
  사용량이 많아지면 한도 초과 가능 — 그럴 땐 네이버 개발자센터에서 한도 상향 신청.
- 월간 PR리포트 탭의 "기사형식 자동분류"는 기사 본문을 하나씩 웹에서 가져와 확인하기 때문에
  기사 수에 비례해서 시간이 걸림 (기사당 약 0.5~1초).
