import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="앱 푸시 성과 대시보드", layout="wide")


# ──────────────────────────────────────────────────────────────
# Mock Data
# ──────────────────────────────────────────────────────────────
@st.cache_data
def make_mock_data(seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)

    copies = [
        "오늘만 30% 할인 받아가세요",
        "장바구니에 담긴 상품, 품절 임박!",
        "회원님만을 위한 추천 상품",
        "지금 구매하면 무료배송",
        "놓치면 후회하는 마감 세일",
        "신상 입고! 가장 먼저 만나보세요",
        "포인트 소멸 D-3, 지금 사용하세요",
        "재입고 알림, 지금 확인하세요",
    ]
    creatives = ["A_제품컷", "B_모델컷", "C_텍스트강조", "D_라이프스타일", "E_할인배지"]
    ab_groups = ["A그룹", "B그룹"]

    dates = pd.date_range("2025-05-01", "2025-06-28", freq="D")

    rows = []
    for d in dates:
        n_per_day = rng.integers(3, 7)
        for _ in range(n_per_day):
            copy = rng.choice(copies)
            creative = rng.choice(creatives)
            ab = rng.choice(ab_groups)
            hour = int(rng.choice([8, 9, 11, 13, 15, 18, 20, 21, 22]))
            body_len = len(copy) + int(rng.integers(-5, 15))

            sends = int(rng.integers(8000, 60000))

            # 시간대/문구 길이/소재에 따른 효율 차등 (스윗스팟 시뮬레이션)
            hour_factor = 1.3 if hour in (8, 21) else (1.1 if hour in (13, 20) else 0.9)
            len_factor = 1.2 if 15 <= body_len <= 25 else 0.85
            creative_factor = {"A_제품컷": 1.0, "B_모델컷": 1.15, "C_텍스트강조": 0.9,
                                "D_라이프스타일": 1.1, "E_할인배지": 1.05}[creative]
            ab_factor = 1.08 if ab == "B그룹" else 1.0

            base_ctr = 0.045 * hour_factor * len_factor * creative_factor * ab_factor
            ctr = float(np.clip(rng.normal(base_ctr, 0.006), 0.005, 0.25))
            clicks = int(sends * ctr)

            base_cr = 0.12 * hour_factor * len_factor * ab_factor
            cr = float(np.clip(rng.normal(base_cr, 0.02), 0.0, 0.6))
            orders = int(clicks * cr)

            aov = float(rng.normal(45000, 8000))
            revenue = max(orders * aov, 0)
            cost = sends * 1.2  # 발송 단가(원) 가정
            roas = revenue / cost if cost > 0 else 0

            rows.append({
                "발송일자": d, "푸시문구": copy, "소재그룹": creative, "AB그룹": ab,
                "발송시간대": hour, "문구길이": body_len, "발송수": sends,
                "클릭수": clicks, "전환수": orders, "매출": round(revenue),
                "비용": round(cost), "CTR": ctr, "CR": cr, "ROAS": roas,
            })

    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────
# Tabs
# ──────────────────────────────────────────────────────────────
st.title("LF몰 앱 푸시 성과 대시보드")

tab_brand, tab_bpu = st.tabs(["브랜드", "BPU"])

with tab_brand:
    st.info("브랜드 탭은 기존 대시보드 영역입니다. (이 파일에서는 BPU 탭만 신규 구현)")

with tab_bpu:
    df = make_mock_data()

    # ── 필터 ──
    st.sidebar.header("BPU 필터")
    min_d, max_d = df["발송일자"].min().date(), df["발송일자"].max().date()
    date_range = st.sidebar.date_input("기간", value=(min_d, max_d), min_value=min_d, max_value=max_d)
    creative_opts = ["전체"] + sorted(df["소재그룹"].unique().tolist())
    sel_creative = st.sidebar.selectbox("소재 타입", creative_opts)
    ab_opts = ["전체"] + sorted(df["AB그룹"].unique().tolist())
    sel_ab = st.sidebar.selectbox("A/B 그룹", ab_opts)

    fdf = df.copy()
    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        lo, hi = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
        fdf = fdf[(fdf["발송일자"] >= lo) & (fdf["발송일자"] <= hi)]
    if sel_creative != "전체":
        fdf = fdf[fdf["소재그룹"] == sel_creative]
    if sel_ab != "전체":
        fdf = fdf[fdf["AB그룹"] == sel_ab]

    if len(fdf) == 0:
        st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
        st.stop()

    # ── 가. 상단 KPI 요약 ──
    st.subheader("📌 핵심 지표 요약")
    avg_ctr = fdf["클릭수"].sum() / fdf["발송수"].sum()
    avg_cr = fdf["전환수"].sum() / fdf["클릭수"].sum() if fdf["클릭수"].sum() else 0
    avg_roas = fdf["매출"].sum() / fdf["비용"].sum() if fdf["비용"].sum() else 0

    best = (fdf.groupby("푸시문구")
            .apply(lambda g: g["클릭수"].sum() / g["발송수"].sum())
            .sort_values(ascending=False))
    best_copy = best.index[0]
    best_ctr = best.iloc[0]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("평균 CTR", f"{avg_ctr*100:.2f}%")
    k2.metric("평균 CR", f"{avg_cr*100:.2f}%")
    k3.metric("평균 ROAS", f"{avg_roas*100:.0f}%")
    k4.metric("베스트 푸시 문구", best_copy[:14] + ("…" if len(best_copy) > 14 else ""),
              delta=f"CTR {best_ctr*100:.2f}%")

    st.divider()

    # ── 나. 문구/소재별 성과 비교 ──
    st.subheader("📊 문구·소재별 CTR / CR 비교")
    compare_dim = st.radio("비교 기준", ["푸시문구", "소재그룹"], horizontal=True, key="bpu_compare_dim")
    g = fdf.groupby(compare_dim).agg(발송수=("발송수", "sum"), 클릭수=("클릭수", "sum"),
                                      전환수=("전환수", "sum")).reset_index()
    g["CTR"] = g["클릭수"] / g["발송수"]
    g["CR"] = g["전환수"] / g["클릭수"].replace(0, np.nan)
    g = g.sort_values("CTR", ascending=False)

    fig_bar = go.Figure()
    fig_bar.add_trace(go.Bar(x=g[compare_dim], y=g["CTR"] * 100, name="CTR(%)"))
    fig_bar.add_trace(go.Bar(x=g[compare_dim], y=g["CR"] * 100, name="CR(%)"))
    fig_bar.update_layout(barmode="group", height=420, xaxis_tickangle=-30,
                           yaxis_title="%", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()

    # ── 다. 성과 추이 분석 ──
    st.subheader("📈 발송 일자별 CTR / CR 추이 (A/B 비교)")
    trend = fdf.groupby(["발송일자", "AB그룹"]).agg(
        발송수=("발송수", "sum"), 클릭수=("클릭수", "sum"), 전환수=("전환수", "sum")).reset_index()
    trend["CTR"] = trend["클릭수"] / trend["발송수"]
    trend["CR"] = trend["전환수"] / trend["클릭수"].replace(0, np.nan)

    metric_choice = st.radio("지표", ["CTR", "CR"], horizontal=True, key="bpu_trend_metric")
    fig_line = px.line(trend, x="발송일자", y=metric_choice, color="AB그룹", markers=True)
    fig_line.update_layout(height=400, yaxis_tickformat=".1%")
    st.plotly_chart(fig_line, use_container_width=True)

    st.divider()

    # ── 라. 발송 조건(시간대/문구 길이)에 따른 효율 ──
    st.subheader("🔥 발송 시간대 × 문구 길이 효율 히트맵")
    heat_metric = st.radio("효율 지표", ["CTR", "CR"], horizontal=True, key="bpu_heat_metric")

    fdf["문구길이대"] = pd.cut(fdf["문구길이"], bins=[0, 10, 15, 20, 25, 30, 100],
                            labels=["~10", "11-15", "16-20", "21-25", "26-30", "31+"])
    hm = fdf.groupby(["발송시간대", "문구길이대"]).agg(
        발송수=("발송수", "sum"), 클릭수=("클릭수", "sum"), 전환수=("전환수", "sum")).reset_index()
    hm["CTR"] = hm["클릭수"] / hm["발송수"]
    hm["CR"] = hm["전환수"] / hm["클릭수"].replace(0, np.nan)
    pivot = hm.pivot(index="발송시간대", columns="문구길이대", values=heat_metric).sort_index()

    fig_heat = go.Figure(go.Heatmap(
        z=pivot.values, x=[str(c) for c in pivot.columns], y=[str(i) + "시" for i in pivot.index],
        colorscale="YlOrRd", colorbar=dict(title=f"{heat_metric}")))
    fig_heat.update_layout(height=420, xaxis_title="문구 길이(자)", yaxis_title="발송 시간대")
    st.plotly_chart(fig_heat, use_container_width=True)

    st.divider()

    # ── 마. 상세 데이터 테이블 ──
    st.subheader("📋 상세 데이터")
    show_cols = ["발송일자", "푸시문구", "소재그룹", "AB그룹", "발송시간대", "문구길이",
                 "발송수", "클릭수", "전환수", "매출", "CTR", "CR", "ROAS"]
    table = fdf[show_cols].sort_values("발송일자", ascending=False).copy()
    table["CTR"] = (table["CTR"] * 100).round(2)
    table["CR"] = (table["CR"] * 100).round(2)
    table["ROAS"] = (table["ROAS"] * 100).round(1)
    st.dataframe(table, use_container_width=True, hide_index=True, height=420)
