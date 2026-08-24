"""
우수발송 대시보드 (주말 우수발송건)

기존 [일일 PUSH 발송성과 대시보드]에 추가되는 페이지입니다.

설치 방법
---------
daily_push_perf_dashboard.py에서 이 모듈의 render_premium_dashboard(df)를 import해서
호출하는 방식으로 붙였습니다 (이 레포에는 서로 다른 여러 Streamlit 앱이 같은 루트를
공유하고 있어서, pages/ 폴더 방식은 다른 앱에도 이 페이지가 새어 들어가 버립니다).

render_premium_dashboard(raw)에 넘기는 raw는 아래 EXPECTED_COLS와 같은 한글 컬럼명을
가진 DataFrame이어야 합니다. daily_push_perf_dashboard.py 쪽에서 이미 누적된 데이터의
컬럼명을 이 이름들로 매핑해서 넘겨줍니다.

단독 실행(`streamlit run premium_push_dashboard.py`)도 여전히 가능합니다 — 이 경우
DATA_DIR에서 주차별 엑셀을 자동으로 찾거나, 없으면 업로더를 보여줍니다.
"""

import os
import re
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------------------------------------------- 설정

SHEET_NAME = "소재별 실적(당주)"
FILE_GLOB = "*PUSH실적.xlsx"
DATA_DIR = Path(__file__).resolve().parent / "data"

EXPECTED_COLS = [
    "날짜", "요일", "시간대", "발송유형", "BPU", "우선순위",
    "카테고리", "속성", "브랜드", "기획전",
    "발송", "UV", "VISIT", "고객수", "주문건수", "주문금액",
]
SUM_COLS = ["발송", "UV", "VISIT", "고객수", "주문건수", "주문금액"]
WEEKDAY_ORDER = ["월", "화", "수", "목", "금", "토", "일"]
WEEKEND = ["토", "일"]

BLUE = "blues"


def _st_version():
    parts = []
    for p in str(getattr(st, "__version__", "0.0")).split(".")[:2]:
        parts.append(int(p) if p.isdigit() else 0)
    return tuple(parts)


# Streamlit 1.49부터 use_container_width가 width="stretch"로 바뀌었습니다.
FILL = {"width": "stretch"} if _st_version() >= (1, 49) else {"use_container_width": True}

# ---------------------------------------------------------------- 데이터


@st.cache_data(show_spinner=False)
def load_raw(paths: tuple) -> pd.DataFrame:
    """주차별 엑셀에서 소재 단위 원장을 읽어 하나로 합칩니다."""
    frames = []
    for p in paths:
        name = Path(p).name
        df = pd.read_excel(p, sheet_name=SHEET_NAME)
        df["주차"] = name.split("_")[0]
        df["파일"] = name
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=EXPECTED_COLS + ["주차", "파일"])
    return pd.concat(frames, ignore_index=True)


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """우수발송 행만 남기고 파생 컬럼을 붙입니다."""
    out = df.copy()

    out["발송유형"] = out["발송유형"].astype(str).str.strip()
    out = out[out["발송유형"].str.contains("우수", na=False)].copy()
    if out.empty:
        return out

    # '우수발송 1' / '우수발송1' 표기 흔들림 정규화
    seq = out["발송유형"].str.extract(r"(\d+)")[0]
    out["회차명"] = "우수발송 " + seq.fillna("-")

    for c in SUM_COLS:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0)

    out["우선순위"] = pd.to_numeric(out["우선순위"], errors="coerce").astype("Int64")
    out["요일"] = out["요일"].astype(str).str.strip()

    hhmm = pd.to_numeric(out["시간대"], errors="coerce").fillna(0).astype(int)
    out["슬롯"] = hhmm.astype(str).str.zfill(4).str[:2] + ":00"

    out["날짜"] = out["날짜"].astype(str).str.strip()
    out["회차키"] = out["날짜"] + " " + out["슬롯"]

    out["BPU"] = out["BPU"].astype(str).str.strip()

    # "주차"는 daily_push_perf_dashboard.py에서 넘어온 라벨 문자열(예: "8/17~8/23")이라
    # 그대로 정렬하면 "8/3~8/9"보다 "8/17~8/23"이 사전식으로 먼저 와버린다(문자 '1'<'3').
    # 그래서 날짜에서 직접 월요일 기준 주 시작일을 다시 계산해 정렬 전용 키로 쓴다.
    dts = pd.to_datetime(out["날짜"], format="%Y%m%d", errors="coerce")
    out["주차_정렬"] = (dts - pd.to_timedelta(dts.dt.dayofweek, unit="D")).dt.date
    return out


def summarize(df: pd.DataFrame, keys=None) -> pd.DataFrame:
    """지정한 축으로 집계하고 효율·전환율을 계산합니다."""
    agg = {"회차": ("회차키", "nunique"), "소재수": ("발송", "size")}
    agg.update({c: (c, "sum") for c in SUM_COLS})

    if keys:
        t = df.groupby(keys, dropna=False).agg(**agg).reset_index()
    else:
        t = pd.DataFrame([{
            "회차": df["회차키"].nunique(),
            "소재수": len(df),
            **{c: df[c].sum() for c in SUM_COLS},
        }])

    def _safe_div(num, den):
        num = num.astype(float)
        den = den.astype(float)
        return np.where(den > 0, num / den, np.nan)

    t["효율"] = _safe_div(t["주문금액"], t["발송"])
    t["유입전환율"] = _safe_div(t["UV"], t["발송"])
    t["주문전환율"] = _safe_div(t["고객수"], t["UV"])
    t["객단가"] = _safe_div(t["주문금액"], t["고객수"])
    return t


# ---------------------------------------------------------------- 포맷


def f_won(v):
    return "-" if pd.isna(v) else f"{v / 10_000:,.0f}만원"


def f_cnt(v):
    return "-" if pd.isna(v) else f"{v / 10_000:,.1f}만건"


def f_eff(v):
    return "-" if pd.isna(v) else f"{v:,.1f}원"


def f_pct(v):
    return "-" if pd.isna(v) else f"{v * 100:.2f}%"


def as_table(t: pd.DataFrame, label_cols) -> pd.DataFrame:
    """화면에 뿌릴 표시용 표로 변환합니다."""
    view = t[label_cols].copy()
    view["회차"] = t["회차"]
    view["소재수"] = t["소재수"]
    view["발송"] = t["발송"].map(f_cnt)
    view["UV"] = t["UV"].map(lambda v: f"{v:,.0f}")
    view["주문고객"] = t["고객수"].map(lambda v: f"{v:,.0f}")
    view["거래액"] = t["주문금액"].map(f_won)
    view["효율"] = t["효율"].map(f_eff)
    view["유입전환율"] = t["유입전환율"].map(f_pct)
    view["주문전환율"] = t["주문전환율"].map(f_pct)
    view["객단가"] = t["객단가"].map(lambda v: "-" if pd.isna(v) else f"{v:,.0f}원")
    return view


# ---------------------------------------------------------------- 차트


def heatmap(t, x_field, y_field, y_order, x_title, y_title, height=180, x_order=None):
    """효율(발송건당 거래액) 히트맵. 셀 안에 수치를 같이 찍습니다.

    가독성을 위해: y축 제목은 생략(섹션 제목과 중복 + 세로로 돌아가며 라벨과 겹치는
    문제가 있었음), 셀 사이 흰 테두리로 경계를 분명히 하고, 값/보조텍스트 글자를
    키우고 간격을 넓혔다. 글자색 대비 기준도 데이터 범위 중앙값으로 바꿔 중간톤
    셀에서도 잘 읽히게 했다."""
    # Vega 표현식 파서에 한글 필드명이 들어가지 않도록 ASCII 사본을 씁니다.
    t = t.copy()
    t["eff"] = t["효율"].astype(float)
    lo, hi = t["eff"].min(), t["eff"].max()
    cut = (lo + hi) / 2 if pd.notna(hi) else float("inf")

    base = alt.Chart(t).encode(
        x=alt.X(f"{x_field}:O", title=x_title, sort=x_order,
                axis=alt.Axis(labelAngle=0, labelFontSize=12, titleFontSize=13)),
        y=alt.Y(f"{y_field}:O", title=None, sort=y_order,
                axis=alt.Axis(labelFontSize=13, labelFontWeight="bold", labelPadding=8)),
    )
    rect = base.mark_rect(stroke="white", strokeWidth=3).encode(
        color=alt.Color(
            "효율:Q",
            title="발송건당 거래액(원)",
            scale=alt.Scale(scheme=BLUE),
            legend=alt.Legend(orient="right"),
        ),
        tooltip=[
            alt.Tooltip(f"{y_field}:N", title=y_title),
            alt.Tooltip(f"{x_field}:N", title=x_title),
            alt.Tooltip("효율:Q", title="효율(원)", format=",.1f"),
            alt.Tooltip("회차:Q", title="회차"),
            alt.Tooltip("발송:Q", title="발송", format=","),
            alt.Tooltip("주문금액:Q", title="거래액", format=","),
            alt.Tooltip("유입전환율:Q", title="유입전환율", format=".2%"),
            alt.Tooltip("주문전환율:Q", title="주문전환율", format=".2%"),
        ],
    )
    label = base.mark_text(fontSize=19, fontWeight="bold", dy=-9).encode(
        text=alt.Text("효율:Q", format=",.1f"),
        color=alt.condition(
            alt.datum.eff > cut, alt.value("white"), alt.value("#1f2933"),
        ),
    )
    sub = base.mark_text(fontSize=12, dy=13).encode(
        text=alt.Text("라벨:N"),
        color=alt.condition(
            alt.datum.eff > cut, alt.value("#e8f1fc"), alt.value("#52514e"),
        ),
    )
    return (rect + label + sub).properties(height=height)


def trend_chart(t):
    """회차별 효율 추이."""
    bar = alt.Chart(t).mark_bar(size=26, color="#2a78d6").encode(
        x=alt.X("회차라벨:N", title=None, sort=None,
                axis=alt.Axis(labelAngle=-40)),
        y=alt.Y("효율:Q", title="발송건당 거래액(원)"),
        tooltip=[
            alt.Tooltip("회차라벨:N", title="회차"),
            alt.Tooltip("효율:Q", title="효율(원)", format=",.1f"),
            alt.Tooltip("발송:Q", title="발송", format=","),
            alt.Tooltip("주문금액:Q", title="거래액", format=","),
        ],
    )
    mean = alt.Chart(t).mark_rule(
        strokeDash=[6, 4], color="#898781"
    ).encode(y="mean(효율):Q")
    return (bar + mean).properties(height=260)


def share_chart(t):
    """우선순위별 거래액 비중."""
    t = t.copy()
    t["비중"] = t["주문금액"] / t["주문금액"].sum()
    t["순위라벨"] = t["우선순위"].astype(str) + "순위"
    return alt.Chart(t).mark_bar(size=34).encode(
        x=alt.X("비중:Q", title="거래액 비중", axis=alt.Axis(format=".0%"),
                stack="zero"),
        color=alt.Color("순위라벨:N", title="우선순위",
                        scale=alt.Scale(range=["#2a78d6", "#85B7EB", "#D8E7F8"])),
        order=alt.Order("우선순위:Q"),
        tooltip=[
            alt.Tooltip("순위라벨:N", title="우선순위"),
            alt.Tooltip("비중:Q", title="비중", format=".1%"),
            alt.Tooltip("주문금액:Q", title="거래액", format=","),
            alt.Tooltip("효율:Q", title="효율(원)", format=",.1f"),
        ],
    ).properties(height=90)


# ---------------------------------------------------------------- 화면


# ---------------------------------------------------------------- AI 대화 (Gemini)
# daily_push_perf_dashboard.py의 "UV 유입 분석" 페이지에 붙인 AI 채팅과 동일한 방식.
# 이 모듈은 daily_push_perf_dashboard.py에서 import돼 쓰이기도 하고 단독 실행도
# 되므로, 공용 함수를 거기서 가져다 쓰지 않고 이 파일 안에 그대로 둔다.
AI_MODELS = {
    "Gemini 2.5 Flash (균형)": "gemini-2.5-flash",
    "Gemini 3.1 Pro Preview (최고 품질·무료 사용량 없을 수 있음)": "gemini-3.1-pro-preview",
    "Gemini 2.5 Flash-Lite (빠름·저렴)": "gemini-2.5-flash-lite",
}
AI_FREE_TIER_SAFE_MODEL = "gemini-2.5-flash"


def gemini_key():
    for k in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        try:
            if k in st.secrets:
                return st.secrets[k]
        except Exception:
            pass
        v = os.environ.get(k)
        if v:
            return v
    return None


def ai_chat_generate(system, history, model):
    """history: [(role, text), ...] role은 'user' 또는 'assistant'.
    반환값: (텍스트, 에러메시지, 실제로 사용된 모델명)."""
    key = gemini_key()
    if not key:
        return None, "GEMINI_API_KEY 미설정 — Streamlit Secrets 또는 환경변수에 추가하세요.", model
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        return None, "google-genai 패키지가 없습니다. requirements.txt 반영 후 재배포하세요.", model

    client = genai.Client(api_key=key)
    contents = [
        types.Content(role=("user" if role == "user" else "model"), parts=[types.Part(text=text)])
        for role, text in history
    ]

    def _call(m):
        resp = client.models.generate_content(
            model=m, contents=contents,
            config=types.GenerateContentConfig(system_instruction=system, max_output_tokens=4000),
        )
        return (resp.text or "").strip()

    def _friendly_error(msg):
        if "RESOURCE_EXHAUSTED" in msg or "429" in msg:
            wait = re.search(r"retry in ([\d.]+)s", msg)
            wait_txt = f" 약 {round(float(wait.group(1)))}초 후 다시 시도해보세요." if wait else " 잠시 후 다시 시도해보세요."
            return (f"⏳ 이 모델의 무료 사용량 한도를 초과했어요.{wait_txt} 계속 막히면 더 가벼운 모델"
                   "(Flash/Flash-Lite)로 바꾸거나, Google AI Studio에서 결제를 연결해보세요.")
        return f"생성 오류: {msg}"

    try:
        text = _call(model)
        return (text or None), (None if text else "빈 응답"), model
    except Exception as e:
        msg = str(e)
        if ("RESOURCE_EXHAUSTED" in msg or "429" in msg) and model != AI_FREE_TIER_SAFE_MODEL:
            try:
                text = _call(AI_FREE_TIER_SAFE_MODEL)
                return (text or None), (None if text else "빈 응답"), AI_FREE_TIER_SAFE_MODEL
            except Exception as e2:
                return None, _friendly_error(str(e2)), model
        if "NOT_FOUND" in msg or "404" in msg:
            alts = re.findall(r"models/([\w.\-]+)", msg)
            alt_model = alts[-1] if alts else None
            if alt_model and alt_model != model:
                try:
                    text = _call(alt_model)
                    return (text or None), (None if text else "빈 응답"), alt_model
                except Exception as e2:
                    return None, _friendly_error(str(e2)), model
        return None, _friendly_error(msg), model


def render_premium_dashboard(raw: pd.DataFrame):
    df = prepare(raw)

    if df.empty:
        st.warning("우수발송 데이터가 없습니다. 발송유형 컬럼에 '우수발송'이 포함된 행이 있는지 확인해 주세요.")
        return

    # ------- 필터
    with st.sidebar:
        st.markdown("### 우수발송 필터")
        weeks = sorted(df["주차"].dropna().unique().tolist())
        sel_week = st.multiselect("주차", weeks, default=weeks)

        weekend_only = st.checkbox("주말(토·일)만 보기", value=True)

        days = [d for d in WEEKDAY_ORDER if d in set(df["요일"])]
        default_days = [d for d in days if d in WEEKEND] if weekend_only else days
        sel_day = st.multiselect("요일", days, default=default_days)

        slots = sorted(df["슬롯"].unique().tolist())
        sel_slot = st.multiselect("시간대", slots, default=slots)

    d = df[
        df["주차"].isin(sel_week)
        & df["요일"].isin(sel_day)
        & df["슬롯"].isin(sel_slot)
    ].copy()

    st.title("우수발송 대시보드")
    if weekend_only:
        st.caption(
            f"주말 우수발송 · {len(sel_week)}개 주차 · "
            f"{d['회차키'].nunique()}회차 / {len(d)}소재"
        )
    else:
        st.caption(f"{len(sel_week)}개 주차 · {d['회차키'].nunique()}회차 / {len(d)}소재")

    if d.empty:
        st.info("선택한 조건에 해당하는 발송이 없습니다.")
        return

    # ------- KPI
    tot = summarize(d).iloc[0]
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 발송", f_cnt(tot["발송"]))
    k2.metric("거래액", f_won(tot["주문금액"]))
    k3.metric("평균 효율", f_eff(tot["효율"]))
    k4.metric("유입 / 주문 전환율", f"{f_pct(tot['유입전환율'])} / {f_pct(tot['주문전환율'])}")

    st.divider()

    # ------- 요일 x 시간대
    st.subheader("요일 × 시간대별 효율")

    dt = summarize(d, ["요일", "슬롯"])
    dt["라벨"] = dt["회차"].astype(str) + "회 · " + dt["발송"].map(f_cnt)
    y_order = [x for x in WEEKDAY_ORDER if x in set(dt["요일"])]
    st.altair_chart(
        heatmap(dt, "슬롯", "요일", y_order, "시간대", "요일",
                height=max(170, 110 * len(y_order))),
        **FILL,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**시간대별 합계**")
        s = summarize(d, ["슬롯"]).sort_values("효율", ascending=False)
        st.dataframe(as_table(s, ["슬롯"]), hide_index=True, **FILL)
    with c2:
        st.markdown("**요일별 합계**")
        s = summarize(d, ["요일"]).sort_values("효율", ascending=False)
        st.dataframe(as_table(s, ["요일"]), hide_index=True, **FILL)

    best = dt.loc[dt["효율"].idxmax()]
    worst = dt.loc[dt["효율"].idxmin()]
    st.info(
        f"최고 조합은 **{best['요일']} {best['슬롯']}** ({f_eff(best['효율'])}, "
        f"{int(best['회차'])}회차), 최저는 **{worst['요일']} {worst['슬롯']}** "
        f"({f_eff(worst['효율'])}) 입니다. "
        "회차 수가 적은 칸은 단일 회차 성과에 크게 흔들리니 함께 보세요."
    )

    st.divider()

    # ------- BPU x 우선순위
    st.subheader("BPU × 우선순위별 실적")

    m = summarize(d, ["BPU", "우선순위"])
    m = m[m["우선순위"].notna()].copy()
    m["우선순위"] = m["우선순위"].astype(int)
    m["순위라벨"] = m["우선순위"].astype(str) + "순위"
    m["라벨"] = m["주문금액"].map(f_won) + " · " + m["회차"].astype(str) + "회"

    bpu_order = (
        summarize(d, ["BPU"]).sort_values("주문금액", ascending=False)["BPU"].tolist()
    )
    st.altair_chart(
        heatmap(m, "순위라벨", "BPU", bpu_order, "우선순위", "BPU",
                height=max(170, 110 * len(bpu_order))),
        **FILL,
    )

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**BPU별 실적**")
        s = summarize(d, ["BPU"]).sort_values("주문금액", ascending=False)
        st.dataframe(as_table(s, ["BPU"]), hide_index=True, **FILL)
    with c4:
        st.markdown("**우선순위별 실적**")
        s = summarize(d, ["우선순위"]).dropna(subset=["우선순위"]).sort_values("우선순위")
        s["우선순위"] = s["우선순위"].astype(int).astype(str) + "순위"
        st.dataframe(as_table(s, ["우선순위"]), hide_index=True, **FILL)

    st.markdown("**우선순위별 거래액 비중**")
    p = summarize(d, ["우선순위"]).dropna(subset=["우선순위"])
    p["우선순위"] = p["우선순위"].astype(int)
    st.altair_chart(share_chart(p), **FILL)

    if len(p) > 1:
        p = p.sort_values("우선순위")
        top, second = p.iloc[0], p.iloc[1]
        if pd.notna(second["효율"]) and second["효율"] > 0:
            drop = top["효율"] / second["효율"]
            share = top["주문금액"] / p["주문금액"].sum()
            st.info(
                f"1순위 효율이 2순위의 **{drop:,.1f}배**이고, "
                f"거래액의 **{share:.1%}**를 1순위가 만듭니다. "
                "다만 1순위 자리에 특정 카테고리·속성 소재가 몰려 있으면 "
                "순위 효과와 소재 효과가 섞이니 아래 소재별 상세를 함께 확인하세요."
            )

    st.divider()

    # ------- 속성 / 카테고리
    st.subheader("속성 · 카테고리별 실적")
    c5, c6 = st.columns(2)
    with c5:
        s = summarize(d, ["속성"]).sort_values("효율", ascending=False)
        st.dataframe(as_table(s, ["속성"]), hide_index=True, **FILL)
    with c6:
        s = summarize(d, ["카테고리"]).sort_values("효율", ascending=False)
        st.dataframe(as_table(s, ["카테고리"]), hide_index=True, **FILL)

    st.markdown("**카테고리별 요일 × 시간대 효율**")
    st.caption("카테고리마다 어느 요일·시간대가 잘 나오는지 나란히 비교해요. 진할수록 효율이 좋아요.")
    cat_order = summarize(d, ["카테고리"]).sort_values("주문금액", ascending=False)["카테고리"].tolist()
    cat_order = [c for c in cat_order if str(c).strip()]
    if not cat_order:
        st.info("카테고리 정보가 있는 데이터가 없어요.")
    else:
        cat_cols = st.columns(2)
        for i, cat in enumerate(cat_order):
            cat_d = d[d["카테고리"] == cat]
            ct = summarize(cat_d, ["요일", "슬롯"])
            if ct.empty:
                continue
            ct["라벨"] = ct["회차"].astype(str) + "회 · " + ct["발송"].map(f_cnt)
            y_order = [x for x in WEEKDAY_ORDER if x in set(ct["요일"])]
            tot_cat = summarize(cat_d).iloc[0]
            with cat_cols[i % 2]:
                st.markdown(f"**{cat}** · {f_cnt(tot_cat['발송'])} 발송 · 평균 효율 {f_eff(tot_cat['효율'])}")
                st.altair_chart(
                    heatmap(ct, "슬롯", "요일", y_order, "시간대", "요일",
                            height=max(130, 85 * len(y_order))),
                    **FILL,
                )

    st.divider()

    # ------- 주차별 추이
    st.subheader("주차별 효율 추이")
    st.caption("우수발송(주말) 실적을 월~일 기준 주차로 묶어서 봐요.")
    w = summarize(d, ["주차_정렬", "주차"]).sort_values("주차_정렬").reset_index(drop=True)
    if len(w) < 2:
        st.info("추이를 보려면 최소 2개 주차 이상의 데이터가 필요해요.")
    else:
        w["회차라벨"] = w["주차"]
        st.altair_chart(trend_chart(w), **FILL)
        st.caption("점선은 선택 구간의 평균 효율입니다.")

        st.markdown("##### 전주 대비 비교")
        week_options = w["주차"].tolist()
        sel_week = st.selectbox("기준 주차", week_options, index=len(week_options) - 1, key="premium_wk_pick")
        idx = week_options.index(sel_week)
        cur_row = w.iloc[idx]
        prev_row = w.iloc[idx - 1] if idx > 0 else None

        def _delta_txt(cv, pv):
            if pv is None or pd.isna(pv) or pv == 0 or cv is None or pd.isna(cv):
                return "-"
            pct = (cv - pv) / pv * 100
            arrow = "▲" if pct > 0 else ("▽" if pct < 0 else "-")
            return f"{arrow}{abs(pct):.1f}%"

        wk_metrics = [
            ("발송", "발송", f_cnt), ("UV", "UV", lambda v: f"{v:,.0f}"),
            ("고객수", "주문고객", lambda v: f"{v:,.0f}"), ("주문금액", "거래액", f_won),
            ("효율", "효율", f_eff), ("유입전환율", "유입전환율", f_pct),
            ("주문전환율", "주문전환율", f_pct),
            ("객단가", "객단가", lambda v: "-" if pd.isna(v) else f"{v:,.0f}원"),
        ]
        wk_rows = []
        for key, label, fmt in wk_metrics:
            cv = cur_row[key]
            pv = prev_row[key] if prev_row is not None else None
            wk_rows.append({
                "지표": label,
                "전주": (fmt(pv) if pv is not None and pd.notna(pv) else "-"),
                "기준주": fmt(cv),
                "증감%": _delta_txt(cv, pv),
            })
        st.dataframe(pd.DataFrame(wk_rows), hide_index=True, **FILL)
        if prev_row is None:
            st.caption("이 주차는 비교할 전주 데이터가 없어요(가장 이른 주차).")

        st.markdown("##### 요일 × 주차 효율 히트맵")
        st.caption("같은 주차 안에서도 토요일과 일요일 중 어느 쪽이 더 좋았는지 비교해요.")
        wd = summarize(d, ["주차_정렬", "주차", "요일"]).sort_values("주차_정렬")
        if wd.empty:
            st.info("표시할 데이터가 없어요.")
        else:
            wd["라벨"] = wd["회차"].astype(str) + "회 · " + wd["발송"].map(f_cnt)
            y_order = wd.drop_duplicates("주차")["주차"].tolist()
            x_order = [x for x in WEEKDAY_ORDER if x in set(wd["요일"])]
            st.altair_chart(
                heatmap(wd, "요일", "주차", y_order, "요일", "주차",
                        height=max(170, 70 * len(y_order)), x_order=x_order),
                **FILL,
            )

    st.divider()

    # ------- 회차별 추이
    st.subheader("회차별 효율 추이")
    r = summarize(d, ["주차", "날짜", "요일", "슬롯", "회차명"])
    r = r.sort_values(["날짜", "슬롯"])
    r["회차라벨"] = (
        r["주차"] + " " + r["요일"] + " " + r["슬롯"] + " (" + r["회차명"] + ")"
    )
    st.altair_chart(trend_chart(r), **FILL)
    st.caption("점선은 선택 구간의 평균 효율입니다.")

    with st.expander("회차별 상세"):
        st.dataframe(
            as_table(r, ["주차", "날짜", "요일", "슬롯", "회차명"]),
            hide_index=True, **FILL,
        )

    with st.expander("소재별 원장"):
        cols = [
            "주차", "날짜", "요일", "슬롯", "회차명", "BPU", "우선순위",
            "카테고리", "속성", "브랜드", "기획전",
            "발송", "UV", "고객수", "주문건수", "주문금액",
        ]
        raw_view = d[[c for c in cols if c in d.columns]].copy()
        raw_view["효율"] = (raw_view["주문금액"] / raw_view["발송"]).round(2)
        st.dataframe(
            raw_view.sort_values(["날짜", "슬롯", "우선순위"]),
            hide_index=True, **FILL,
        )
        st.download_button(
            "CSV 내려받기",
            raw_view.to_csv(index=False).encode("utf-8-sig"),
            file_name="우수발송_소재별_원장.csv",
            mime="text/csv",
        )

    st.divider()
    st.markdown("#### 🤖 AI와 지표 논의하기")
    st.caption("위에서 본 우수발송 지표들을 근거로 AI에게 물어보세요. 예: '이번 주 효율이 떨어진 이유가 뭐야?', "
              "'토요일이랑 일요일 중 어디에 더 집중해야 해?', '어떤 BPU를 줄이는 게 좋을까?'")

    def _build_premium_ai_facts():
        tot = summarize(d).iloc[0]
        lines = [
            f"전체(필터 적용): 발송 {tot['발송']:,.0f} · UV {tot['UV']:,.0f} · 주문금액 {tot['주문금액']:,.0f}원 · "
            f"효율 {tot['효율']:.2f} · 유입전환율 {tot['유입전환율']:.2%} · 주문전환율 {tot['주문전환율']:.2%}"
        ]

        wk = summarize(d, ["주차_정렬", "주차"]).sort_values("주차_정렬")
        if not wk.empty:
            lines.append("\n주차별(월~일) 효율 추이:")
            for _, r in wk.iterrows():
                lines.append(f"- {r['주차']}: 발송 {r['발송']:,.0f} · 효율 {r['효율']:.2f} · "
                             f"주문금액 {r['주문금액']:,.0f}원")

        dow = summarize(d, ["요일"])
        if not dow.empty:
            dow_order = [x for x in WEEKDAY_ORDER if x in set(dow["요일"])]
            dow = dow.set_index("요일").loc[dow_order].reset_index()
            lines.append("\n요일별 효율:")
            for _, r in dow.iterrows():
                lines.append(f"- {r['요일']}: 발송 {r['발송']:,.0f} · 효율 {r['효율']:.2f} · "
                             f"주문전환율 {r['주문전환율']:.2%}")

        for dim in ["BPU", "카테고리", "속성", "슬롯"]:
            sg = summarize(d, [dim]) if dim in d.columns else pd.DataFrame()
            sg = sg[sg[dim].astype(str).str.strip() != ""] if not sg.empty else sg
            if sg.empty:
                continue
            sg = sg.sort_values("효율", ascending=False)
            lines.append(f"\n{dim}별 효율 (효율 내림차순):")
            for _, r in sg.iterrows():
                lines.append(f"- {r[dim]}: 발송 {r['발송']:,.0f} · 효율 {r['효율']:.2f} · "
                             f"주문금액 {r['주문금액']:,.0f}원")
        return "\n".join(lines)

    ai_hist_key = "premium_ai_chat_history"
    if ai_hist_key not in st.session_state:
        st.session_state[ai_hist_key] = []

    ac1, ac2 = st.columns([3, 1])
    ai_model_name = ac1.selectbox("AI 모델", list(AI_MODELS), key="premium_ai_model",
                                  index=list(AI_MODELS).index("Gemini 2.5 Flash (균형)"))
    nominal_model = AI_MODELS[ai_model_name]
    ai_model_map = st.session_state.setdefault("premium_ai_working_model", {})
    ai_model = ai_model_map.get(nominal_model, nominal_model)
    if ac2.button("🗑️ 대화 초기화", use_container_width=True, key="premium_ai_reset"):
        st.session_state[ai_hist_key] = []
        st.rerun()

    for role, text in st.session_state[ai_hist_key]:
        with st.chat_message("user" if role == "user" else "assistant"):
            st.markdown(text)

    user_q = st.chat_input("지표에 대해 질문해보세요", key="premium_ai_input")
    if user_q:
        st.session_state[ai_hist_key].append(("user", user_q))
        with st.chat_message("user"):
            st.markdown(user_q)
        system = ("당신은 LF몰 CRM 우수발송(주말) PUSH 데이터 분석가입니다. 아래 [데이터]에 있는 수치만 근거로 "
                  "사용자와 한국어로 대화하세요. 데이터에 없는 값은 지어내지 말고 모른다고 답하세요. "
                  "실무자가 바로 실행할 수 있도록 구체적이고 간결하게 답변하세요.\n\n"
                  f"[데이터]\n{_build_premium_ai_facts()}")
        with st.chat_message("assistant"):
            with st.spinner("생각 중…"):
                txt, err, used_model = ai_chat_generate(system, st.session_state[ai_hist_key], ai_model)
            if err:
                st.warning(err)
                st.session_state[ai_hist_key].pop()
            else:
                if used_model != nominal_model:
                    ai_model_map[nominal_model] = used_model
                    st.caption(f"ℹ️ '{ai_model_name}' 모델은 더 이상 지원되지 않아 "
                              f"`{used_model}`로 자동 전환해서 답했어요.")
                st.markdown(txt)
                st.session_state[ai_hist_key].append(("assistant", txt))


# ---------------------------------------------------------------- 진입점 (단독 실행용)


def _main():
    st.set_page_config(page_title="우수발송 대시보드", page_icon="📈", layout="wide")

    paths = sorted(str(p) for p in DATA_DIR.glob(FILE_GLOB)) if DATA_DIR.exists() else []

    if not paths:
        st.title("우수발송 대시보드")
        st.info(
            f"`{DATA_DIR}` 에서 `{FILE_GLOB}` 파일을 찾지 못했습니다. "
            "아래에 주차별 엑셀을 올리거나, DATA_DIR 경로를 수정해 주세요."
        )
        ups = st.file_uploader(
            "주차별 PUSH 실적 엑셀", type=["xlsx"], accept_multiple_files=True
        )
        if not ups:
            st.stop()
        frames = []
        for f in ups:
            t = pd.read_excel(f, sheet_name=SHEET_NAME)
            t["주차"] = f.name.split("_")[0]
            t["파일"] = f.name
            frames.append(t)
        raw = pd.concat(frames, ignore_index=True)
    else:
        raw = load_raw(tuple(paths))

    render_premium_dashboard(raw)


if __name__ == "__main__":
    _main()
