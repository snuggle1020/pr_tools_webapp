# -*- coding: utf-8 -*-
"""
월간 PR리포트 '이달의 요약'을 Claude로 쓰기.
기사를 하나씩 나열하지 않고 이달 보도의 특징(무엇이 보도를 이끌었는지, 어떤 성격의 매체에 실렸는지,
단순 언급이 어떤 맥락에서 나왔는지, 전월과 무엇이 달라졌는지)을 3~4문장으로 잡아 쓰게 한다.
API 키는 Streamlit Secrets 의 ANTHROPIC_API_KEY (app.py 가 환경변수로 옮겨 둠).
"""
from collections import Counter

import anthropic

from media_order import media_category

MODEL = "claude-opus-5-5"

SYSTEM = """당신은 PR 대행사 오픈피알의 시니어 컨설턴트입니다. 고객사에 보내는 월간 PR 리포트 1쪽의 '이달의 요약'을 씁니다.

요약은 이달 언론 보도의 특징을 짚는 글입니다. 기사나 매체를 하나씩 나열하지 말고, 데이터에서 보이는 흐름을 해석하세요.
다룰 만한 관점 (데이터가 뒷받침하는 것만 고르세요):
- 이달 보도를 이끈 것: 어떤 보도자료·이슈가 보도량의 중심이었는지, 배포 당일에 얼마나 집중됐는지
- 게재 매체의 성격: 종합지·경제지·IT 전문지·산업지 등 어떤 유형에 주로 실렸는지, 그게 고객사에 어떤 의미인지
- 단순 언급의 맥락: 고객사가 어떤 산업 이슈·행사·다른 기업 소식 속에서 함께 거론됐는지 (주제 묶음으로 보고, 기사 하나하나가 아니라 흐름으로)
- 전월·연간 흐름과 비교해 달라진 점, 기획·인터뷰처럼 부족했던 형식이 있다면 다음 달 방향

작성 규칙:
- 3~4문장, 한 문단. 존댓말 평서문('~했습니다').
- 숫자는 꼭 필요한 것만 (예: 보도자료 1건으로 15건 게재). 숫자 나열 금지.
- 매체명은 특징을 보여 주는 대표 예시로만 1~2개까지.
- 주어진 데이터에 없는 사실·평가·전망을 지어내지 마세요. 확실하지 않으면 쓰지 마세요.
- 제목·머리말·글머리표 없이 요약 문단만 출력하세요."""


def _fmt_types(types):
    return ", ".join(f"{k} {v}건" for k, v in types.items() if v) or "없음"


def build_prompt(data, release_dates, articles_all):
    """compose() 결과(data)와 앱 표의 기사 전체로 요약에 쓸 자료를 글로 정리."""
    m, tr = data["month"], data["trend"]
    lines = [f"고객사: {data['company']}", f"리포트: {data['year']}년 {m}월",
             f"이번 달 형식별 보도: {_fmt_types(data['types'])} (합계 {sum(data['types'].values())}건)"]
    if tr and m > 1:
        tot = [tr["release"][i] + tr["planned"][i] + tr["mention"][i] for i in range(m)]
        lines.append("월별 합계(1월부터): " + ", ".join(f"{i + 1}월 {t}건" for i, t in enumerate(tot)))

    lines.append("\n[이번 달 배포한 보도자료]")
    if data["releases"]:
        for r, d in zip(data["releases"], release_dates):
            day = f"{d.month}/{d.day} 배포" if d else "배포일 미상"
            lines.append(f"- {r['title']} ({day}, 게재 {r['count']}건)")
    else:
        lines.append("- 없음")

    main = data["articles"]
    lines.append(f"\n[주요 기사 {len(main)}건] (게재일 | 매체 | 매체 유형 | 형식 | 제목)")
    for a in main:
        d = f"{a['date'].month}/{a['date'].day}" if a["date"] else "-"
        lines.append(f"- {d} | {a['media']} | {media_category(a['media'])} | {a['kind']} | {a['title']}")
    cats = Counter(media_category(a["media"]) for a in main)
    if cats:
        lines.append("주요 기사 매체 유형 분포: " + ", ".join(f"{c} {n}건" for c, n in cats.most_common()))

    lines.append(f"\n[단순 언급 기사 - 같은 소식끼리 묶은 주제 {len(data['mentions'])}개]")
    for mn in data["mentions"]:
        lines.append(f"- {mn['title']} (대표 {mn['media']}, 관련 기사 {mn['count']}건)")

    if data["meetings"]:
        lines.append("\n[기자 미팅]")
        lines += [f"- {a} {b}: {c}" for a, b, c in data["meetings"]]
    if data["plans"]:
        lines.append(f"\n[{data['plan_title']}]")
        lines += [f"- {a}: {b} ({c})" for a, b, c in data["plans"]]
    return "\n".join(lines)


def write_summary(data, release_dates, articles_all):
    """요약 문단을 돌려준다. 실패하면 RuntimeError (메시지는 그대로 화면에 보여 줄 수 있게)."""
    client = anthropic.Anthropic()
    try:
        resp = client.beta.messages.create(
            model=MODEL,
            max_tokens=4000,
            output_config={"effort": "medium"},
            # 안전 분류기가 거절하면 서버가 다른 모델로 자동 재시도
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            system=SYSTEM,
            messages=[{"role": "user", "content": build_prompt(data, release_dates, articles_all)
                       + "\n\n위 자료로 '이달의 요약' 문단을 써 주세요."}],
        )
    except anthropic.AuthenticationError:
        raise RuntimeError("Claude API 키가 없거나 틀렸습니다. Streamlit Secrets의 ANTHROPIC_API_KEY를 확인하세요.")
    except anthropic.RateLimitError:
        raise RuntimeError("요청이 많아 잠시 막혔습니다. 1분쯤 뒤에 다시 눌러 주세요.")
    except anthropic.APIStatusError as e:
        raise RuntimeError(f"Claude API 오류 ({e.status_code}): {e.message}")
    except anthropic.APIConnectionError:
        raise RuntimeError("Claude API에 연결하지 못했습니다. 잠시 뒤 다시 시도하세요.")
    if resp.stop_reason == "refusal":
        raise RuntimeError("요약을 만들지 못했습니다(요청 거절). 직접 작성해 주세요.")
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    if not text:
        raise RuntimeError("요약 결과가 비어 있습니다. 다시 눌러 주세요.")
    return text
