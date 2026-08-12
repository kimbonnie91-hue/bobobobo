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

from pathlib import Path

import altair as alt
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

    발송 = t["발송"].replace(0, pd.NA)
    uv = t["UV"].replace(0, pd.NA)
    고객 = t["고객수"].replace(0, pd.NA)

    t["효율"] = (t["주문금액"] / 발송).astype(float)
    t["유입전환율"] = (t["UV"] / 발송).astype(float)
    t["주문전환율"] = (t["고객수"] / uv).astype(float)
    t["객단가"] = (t["주문금액"] / 고객).astype(float)
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


def heatmap(t, x_field, y_field, y_order, x_title, y_title, height=180):
    """효율(발송건당 거래액) 히트맵. 셀 안에 수치를 같이 찍습니다."""
    # Vega 표현식 파서에 한글 필드명이 들어가지 않도록 ASCII 사본을 씁니다.
    t = t.copy()
    t["eff"] = t["효율"].astype(float)
    hi = t["eff"].max()
    cut = hi * 0.6 if pd.notna(hi) else float("inf")

    base = alt.Chart(t).encode(
        x=alt.X(f"{x_field}:O", title=x_title, sort=None,
                axis=alt.Axis(labelAngle=0)),
        y=alt.Y(f"{y_field}:O", title=y_title, sort=y_order),
    )
    rect = base.mark_rect().encode(
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
    label = base.mark_text(fontSize=15, fontWeight="bold").encode(
        text=alt.Text("효율:Q", format=",.1f"),
        color=alt.condition(
            alt.datum.eff > cut, alt.value("white"), alt.value("#1f2933"),
        ),
    )
    sub = base.mark_text(fontSize=11, dy=16).encode(
        text=alt.Text("라벨:N"),
        color=alt.condition(
            alt.datum.eff > cut, alt.value("white"), alt.value("#52514e"),
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
                height=max(120, 70 * len(y_order))),
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
                height=max(120, 70 * len(bpu_order))),
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
