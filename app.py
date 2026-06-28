import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="LF몰 세션 대시보드", layout="wide")
st.title("LF몰 세션 채널별 대시보드")

uploaded = st.file_uploader("xlsx 파일 업로드", type=["xlsx"])
if uploaded is None:
    st.info("xlsx 파일을 업로드하세요.")
    st.stop()

@st.cache_data
def load_raw(file_bytes):
    import io
    xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name="세션", header=None)
    row0 = xl.iloc[0].tolist()
    row1 = xl.iloc[1].tolist()
    row2 = xl.iloc[2].tolist()

    # 컬럼 메타 정리
    col_meta = {}
    for i in range(1, len(row0)):
        if pd.notna(row0[i]):
            col_meta[i] = {
                "cat": str(row0[i]),
                "sub": str(row1[i]) if pd.notna(row1[i]) else "기타",
                "media": str(row2[i]) if pd.notna(row2[i]) else "기타",
            }

    data = xl.iloc[4:].copy()
    data[0] = data[0].apply(
        lambda x: str(int(x)) if pd.notna(x) and str(x).replace(".0","").isdigit() else None
    )
    data = data.dropna(subset=[0])
    data = data[data[0].str.match(r"^\d{8}$", na=False)]
    data["ym"] = pd.to_datetime(data[0], format="%Y%m%d").dt.to_period("M").astype(str)

    for col in col_meta:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    # 월별로 집계
    rows = []
    for ym, grp in data.groupby("ym"):
        row = {"ym": ym}
        for i, meta in col_meta.items():
            if i in grp.columns:
                key = f"{meta['cat']}||{meta['sub']}||{meta['media']}"
                row[key] = row.get(key, 0) + float(grp[i].sum())
        rows.append(row)

    df = pd.DataFrame(rows).sort_values("ym").reset_index(drop=True)
    df = df.fillna(0)
    return df, col_meta

file_bytes = uploaded.read()
df, col_meta = load_raw(file_bytes)

# 카테고리/서브채널 목록
CATS = ["직접", "광고", "PUSH", "EP", "미디어커머스", "브랜드광고", "제휴"]
COLORS = {
    "직접": "#185FA5", "광고": "#639922", "PUSH": "#BA7517",
    "EP": "#3C3489", "미디어커머스": "#D85A30", "브랜드광고": "#0F6E56", "제휴": "#888780"
}
SUB_COLORS = ["#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd","#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf"]

def get_cols(df, cat, sub=None, media=None):
    cols = []
    for c in df.columns:
        if "||" not in c:
            continue
        parts = c.split("||")
        if parts[0] != cat:
            continue
        if sub and parts[1] != sub:
            continue
        if media and parts[2] != media:
            continue
        cols.append(c)
    return cols

def sum_cols(df, cols):
    if not cols:
        return pd.Series([0]*len(df), index=df.index)
    return df[cols].sum(axis=1)

def get_subs(cat):
    subs = set()
    for meta in col_meta.values():
        if meta["cat"] == cat:
            subs.add(meta["sub"])
    return sorted(subs)

def get_medias(cat, sub):
    medias = set()
    for meta in col_meta.values():
        if meta["cat"] == cat and meta["sub"] == sub:
            medias.add(meta["media"])
    return sorted(medias)

# 사이드바 필터
st.sidebar.header("필터")
months = df["ym"].tolist()
start = st.sidebar.selectbox("시작 월", months, index=0)
end = st.sidebar.selectbox("종료 월", months, index=len(months)-1)
filtered = df[(df["ym"] >= start) & (df["ym"] <= end)].copy().reset_index(drop=True)

st.sidebar.divider()
st.sidebar.subheader("채널 선택")
selected_cats = st.sidebar.multiselect("대분류 채널", CATS, default=CATS)

# 요약 지표
cat_totals = {cat: sum_cols(filtered, get_cols(filtered, cat)).sum() for cat in CATS}
grand_total = sum(cat_totals.values())

cols_m = st.columns(4)
cols_m[0].metric("총 세션", f"{grand_total/1e8:.2f}억")
cols_m[1].metric("직접 비중", f"{cat_totals['직접']/grand_total*100:.1f}%")
cols_m[2].metric("PUSH 비중", f"{cat_totals['PUSH']/grand_total*100:.1f}%")
cols_m[3].metric("광고 비중", f"{cat_totals['광고']/grand_total*100:.1f}%")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["📈 채널별 추이", "📊 비중 추이", "🔍 서브채널 드릴다운", "📋 원본 데이터"])

# --- 탭1: 채널별 추이 ---
with tab1:
    fig1 = go.Figure()
    for cat in selected_cats:
        y = sum_cols(filtered, get_cols(filtered, cat))
        fig1.add_trace(go.Scatter(
            x=filtered["ym"], y=y,
            name=cat, mode="lines",
            line=dict(color=COLORS.get(cat, "#888"), width=2)
        ))
    fig1.update_layout(
        title="채널별 월별 세션 추이",
        height=420, hovermode="x unified",
        xaxis_title="월", yaxis_title="세션 수",
        legend=dict(orientation="h", yanchor="bottom", y=1.02)
    )
    st.plotly_chart(fig1, use_container_width=True)

# --- 탭2: 비중 추이 ---
with tab2:
    fig2 = go.Figure()
    for cat in selected_cats:
        y_abs = sum_cols(filtered, get_cols(filtered, cat))
        total_row = pd.Series([0.0]*len(filtered), index=filtered.index)
        for c in CATS:
            total_row += sum_cols(filtered, get_cols(filtered, c))
        y_pct = (y_abs / total_row.replace(0, 1) * 100).round(1)
        fig2.add_trace(go.Bar(
            x=filtered["ym"], y=y_pct,
            name=cat, marker_color=COLORS.get(cat, "#888")
        ))
    fig2.update_layout(
        title="채널별 세션 비중 (%)",
        barmode="stack", height=420,
        hovermode="x unified",
        xaxis_title="월", yaxis_title="비중(%)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02)
    )
    st.plotly_chart(fig2, use_container_width=True)

# --- 탭3: 서브채널 드릴다운 ---
with tab3:
    st.subheader("서브채널 드릴다운")
    drill_cat = st.selectbox("채널 선택", CATS)
    subs = get_subs(drill_cat)

    # 서브채널별 집계
    sub_totals = {}
    for sub in subs:
        val = sum_cols(filtered, get_cols(filtered, drill_cat, sub=sub)).sum()
        sub_totals[sub] = val
    sub_total_all = sum(sub_totals.values())

    # 서브채널 비중 파이차트
    col_pie, col_bar = st.columns([1, 2])
    with col_pie:
        st.markdown(f"**{drill_cat} 서브채널 비중**")
        labels = list(sub_totals.keys())
        values = list(sub_totals.values())
        fig_pie = go.Figure(go.Pie(
            labels=labels, values=values,
            hole=0.4,
            marker_colors=SUB_COLORS[:len(labels)],
            textinfo="label+percent"
        ))
        fig_pie.update_layout(height=320, showlegend=False, margin=dict(t=10,b=10,l=10,r=10))
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        st.markdown(f"**{drill_cat} 서브채널별 월별 추이**")
        selected_subs = st.multiselect(
            "서브채널 선택", subs,
            default=sorted(sub_totals, key=lambda x: -sub_totals[x])[:5]
        )
        fig_sub = go.Figure()
        for i, sub in enumerate(selected_subs):
            y = sum_cols(filtered, get_cols(filtered, drill_cat, sub=sub))
            fig_sub.add_trace(go.Scatter(
                x=filtered["ym"], y=y,
                name=sub, mode="lines",
                line=dict(color=SUB_COLORS[i % len(SUB_COLORS)], width=2)
            ))
        fig_sub.update_layout(
            height=320, hovermode="x unified",
            xaxis_title="월", yaxis_title="세션 수",
            margin=dict(t=10), legend=dict(orientation="h", yanchor="bottom", y=1.02)
        )
        st.plotly_chart(fig_sub, use_container_width=True)

    # 미디어별 드릴다운
    st.markdown("---")
    st.markdown(f"**미디어별 세부 분석** (서브채널 선택 후 보기)")
    drill_sub = st.selectbox("서브채널 선택", subs, key="drill_sub")
    medias = get_medias(drill_cat, drill_sub)
    if medias:
        media_totals = {m: sum_cols(filtered, get_cols(filtered, drill_cat, sub=drill_sub, media=m)).sum() for m in medias}
        media_sorted = sorted(media_totals.items(), key=lambda x: -x[1])

        fig_media = go.Figure()
        for i, (media, _) in enumerate(media_sorted[:10]):
            y = sum_cols(filtered, get_cols(filtered, drill_cat, sub=drill_sub, media=media))
            fig_media.add_trace(go.Bar(
                x=filtered["ym"], y=y,
                name=media,
                marker_color=SUB_COLORS[i % len(SUB_COLORS)]
            ))
        fig_media.update_layout(
            title=f"{drill_cat} > {drill_sub} 미디어별 추이",
            barmode="stack", height=360,
            hovermode="x unified",
            xaxis_title="월", yaxis_title="세션 수",
            legend=dict(orientation="h", yanchor="bottom", y=1.02)
        )
        st.plotly_chart(fig_media, use_container_width=True)

# --- 탭4: 원본 데이터 ---
with tab4:
    st.subheader("월별 채널별 집계")
    summary = filtered[["ym"]].copy()
    for cat in CATS:
        summary[cat] = sum_cols(filtered, get_cols(filtered, cat)).astype(int)
    summary["합계"] = summary[CATS].sum(axis=1)
    st.dataframe(summary, use_container_width=True, height=400)

    st.divider()
    st.subheader("서브채널별 집계 다운로드")
    dl_cat = st.selectbox("다운로드할 채널", CATS, key="dl_cat")
    dl_data = filtered[["ym"]].copy()
    for sub in get_subs(dl_cat):
        dl_data[sub] = sum_cols(filtered, get_cols(filtered, dl_cat, sub=sub)).astype(int)
    st.dataframe(dl_data, use_container_width=True, height=300)
    csv = dl_data.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"{dl_cat} 서브채널 데이터 다운로드 (CSV)",
        data=csv,
        file_name=f"{dl_cat}_subchannel.csv",
        mime="text/csv"
    )
