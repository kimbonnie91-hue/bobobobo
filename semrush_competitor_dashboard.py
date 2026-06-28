"""LF몰 경쟁사 SEO 키워드 분석 대시보드 — Semrush 데이터 시각화 (Plotly)

데이터 출처 : Semrush MCP (organic_research / overview_research)
데이터베이스 : kr (한국 / google.co.kr)
스냅샷 일자  : 2026-06-22

────────────────────────────────────────────────────────────────────────
데이터 갱신 방법 (Semrush API / MCP)
────────────────────────────────────────────────────────────────────────
아래 SNAPSHOT 딕셔너리들은 Semrush에서 추출한 정적 스냅샷이다.
최신화하려면 Semrush MCP 또는 API로 동일 리포트를 다시 뽑아 값만 교체하면 된다.

  · 도메인 개요  : overview_research → domain_rank   (params: domain, database='kr')
  · 상위 키워드  : organic_research  → domain_organic (params: domain, database='kr',
                                                       display_sort='tr_desc')
  · 키워드 갭    : organic_research  → domain_domains (params: domains, database='kr')

API 직접 호출 예시 (HTTP):
  https://api.semrush.com/?type=domain_rank&key=<API_KEY>&domain=lfmall.co.kr&database=kr
────────────────────────────────────────────────────────────────────────
실행: streamlit run semrush_competitor_dashboard.py
"""

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="경쟁사 SEO 키워드 분석", page_icon="🔍",
                   layout="wide", initial_sidebar_state="expanded")

SNAPSHOT_DATE = "2026-06-22"
DATABASE = "kr"

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.card{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:8px 0;line-height:1.65;background:#ffffff}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.tag{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:6px;margin-right:6px}
h1,h2,h3{color:#1e293b}
.cap{color:#64748b;font-size:13px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상 팔레트 (기존 대시보드 컨벤션 유지)
# ══════════════════════════════════════════════════════
PALETTE = {
    "blue":   "rgba(79,143,255,1)",
    "red":    "rgba(245,101,101,1)",
    "amber":  "rgba(237,137,54,1)",
    "purple": "rgba(159,122,234,1)",
    "slate":  "rgba(100,116,139,1)",
    "green":  "rgba(72,187,120,1)",
}

# 사이트별 색상 (LF몰 = 우리, 파랑 / SSF = 위협, 빨강)
SITE_COLOR = {
    "LF몰":    PALETTE["blue"],
    "SSF샵":   PALETTE["red"],
    "W컨셉":   PALETTE["amber"],
    "한섬":    PALETTE["purple"],
    "SI빌리지": PALETTE["slate"],
}

def base_layout(h=300, ysuffix="", xsuffix="", title="", showlegend=False):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#475569", size=12), margin=dict(l=10, r=10, t=40, b=10),
        height=h, showlegend=showlegend,
        title=dict(text=title, font=dict(color="#94a3b8", size=13)),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                    font=dict(size=11)),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=xsuffix),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=ysuffix),
    )

# ══════════════════════════════════════════════════════
# DATA SNAPSHOT — Semrush kr DB (2026-06-22)
# ══════════════════════════════════════════════════════

# ── 도메인 개요 (overview_research / domain_rank) ──
OVERVIEW = [
    # 사,      도메인,            랭크,  오가닉KW, 오가닉트래픽, 오가닉가치, 광고KW, 광고트래픽, 광고비
    ("SSF샵",   "ssfshop.com",      651,  17740, 173025, 134566, 263, 39345, 96994),
    ("W컨셉",   "wconcept.co.kr",   766,   6095, 147646,  58407,  23, 10273,  6395),
    ("LF몰",    "lfmall.co.kr",     786,  21881, 144711,  45301,  89,  8842,  2198),
    ("한섬",    "thehandsome.com", 1622,   2653,  62091,  13852,  48,  3221,   863),
    ("SI빌리지", "sivillage.com",   5892,    991,  11812,   6132,   0,     0,     0),
]
OV_COLS = ["사", "도메인", "도메인랭크", "오가닉 키워드", "월 오가닉 트래픽",
           "트래픽 가치($)", "광고 키워드", "광고 트래픽", "광고비($)"]
df_ov = pd.DataFrame(OVERVIEW, columns=OV_COLS)
df_ov["키워드당 트래픽"] = (df_ov["월 오가닉 트래픽"] / df_ov["오가닉 키워드"]).round(1)

# ── 사별 상위 키워드 (organic_research / domain_organic, 트래픽순) ──
# (키워드, 순위, 검색량, 트래픽기여%)
TOP_KEYWORDS = {
    "LF몰": [
        ("lf 몰", 1, 18100, 10.00), ("lfmall", 1, 9900, 5.47),
        ("엘에프몰", 1, 2900, 1.60), ("팅커벨", 9, 60500, 1.25),
        ("허리띠", 1, 2400, 0.77), ("모피", 5, 22200, 0.76),
        ("크롬하츠", 5, 18100, 0.62), ("여자 팬티", 2, 9900, 0.61),
        ("금팔찌", 1, 1600, 0.51), ("토니 웩", 5, 12100, 0.41),
        ("g 컵", 5, 14800, 0.40), ("경량 패딩", 5, 12100, 0.25),
    ],
    "SSF샵": [
        ("ssf", 1, 27100, 12.52), ("에잇 세컨즈", 1, 27100, 12.52),
        ("ssf 몰", 1, 12100, 5.59), ("ssfshop", 1, 8100, 3.74),
        ("ssf 샵", 1, 3600, 1.66), ("8 seconds", 1, 2900, 1.34),
        ("띠어리", 1, 4400, 1.19), ("beaker", 1, 3600, 0.97),
        ("르 베이지", 1, 1900, 0.87), ("구호", 1, 2400, 0.65),
        ("꼼데 가르송", 4, 22200, 0.64), ("핏 플랍", 1, 2400, 0.65),
    ],
    "W컨셉": [
        ("w 컨셉", 1, 33100, 17.93), ("w concept", 1, 14800, 8.01),
        ("늑대 닷컴*", 14, 1830000, 6.19), ("더블유 컨셉", 1, 8100, 4.38),
        ("w", 1, 49500, 4.35), ("비키니", 5, 60500, 2.04),
        ("하객 룩", 1, 5400, 1.71), ("니트", 2, 18100, 1.10),
        ("크로스 백", 2, 12100, 1.06), ("블라우스", 2, 22200, 0.75),
        ("부츠", 3, 12100, 0.73), ("스투시", 3, 60500, 0.73),
    ],
    "한섬": [
        ("한섬", 1, 14800, 19.06), ("한섬 몰", 1, 8100, 10.43),
        ("한섬몰", 1, 8100, 10.43), ("더한섬닷컴", 1, 4400, 5.66),
        ("무스너클", 1, 5400, 4.08), ("타임옴므", 1, 2900, 2.19),
        ("더 한섬", 1, 1600, 2.06), ("시스템옴므", 1, 2400, 1.81),
        ("dkny", 1, 2400, 1.81), ("클럽모나코", 1, 1900, 1.43),
        ("타임", 1, 5400, 1.13), ("시스템", 2, 6600, 0.95),
    ],
    "SI빌리지": [
        ("sivillage", 1, 5400, 36.57), ("자주", 2, 12100, 13.31),
        ("si 빌리지", 1, 1300, 8.80), ("jaju", 2, 4400, 4.84),
        ("시빌 리지", 1, 1000, 3.97), ("에스아이 빌리지", 1, 880, 3.49),
        ("어그", 6, 9900, 3.35), ("시마을", 2, 2400, 2.64),
        ("vov", 1, 480, 1.90), ("스튜디오 톰보이", 5, 2900, 1.22),
        ("신세계 빌리지", 1, 320, 1.26), ("신세계 인터내셔날", 5, 2400, 1.01),
    ],
}

# ── 키워드 갭: LF몰 미보유(순위 0) 고검색량 키워드 (organic_research / domain_domains) ──
# (키워드, 검색량, 보유사, 보유사순위)
GAP_MISSING = [
    ("아디다스", 90500, "SSF샵", 14), ("청바지", 74000, "SSF샵", 27),
    ("비키니", 60500, "W컨셉", 5), ("adidas", 49500, "SSF샵", 12),
    ("바지", 40500, "SSF샵", 14), ("파타고니아", 40500, "SSF샵", 14),
    ("cos", 40500, "SSF샵", 6), ("패딩", 33100, "W컨셉", 11),
    ("가방", 33100, "W컨셉", 5), ("젠틀몬스터", 33100, "SSF샵", 10),
    ("코트", 27100, "SSF샵", 18), ("화장품", 27100, "SSF샵", 32),
    ("몽벨", 27100, "SSF샵", 13), ("스케쳐스", 27100, "SSF샵", 18),
    ("신세계 몰", 27100, "SI빌리지", 17),
]

# ── LF 퀵윈: 이미 순위는 있으나 5위 밖 → 상위 진입 시 트래픽 급증 (domain_organic) ──
# (키워드, 현재순위, 검색량)
QUICK_WINS = [
    ("팅커벨", 9, 60500), ("하프 클럽", 14, 40500), ("크롬 하츠", 9, 33100),
    ("마시 모두 띠", 11, 27100), ("모피", 5, 22200), ("24k 순금 1돈 가격", 13, 22200),
    ("on", 15, 22200), ("꼼데 가르송", 26, 22200), ("크롬하츠", 5, 18100),
    ("g 컵", 5, 14800), ("토니 웩", 5, 12100), ("경량 패딩", 5, 12100),
]

# ══════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════
st.markdown("## 🔍 LF몰 경쟁사 SEO 키워드 분석")
st.markdown(
    f"<div class='cap'>Semrush 한국(<b>{DATABASE}</b>) DB · 스냅샷 {SNAPSHOT_DATE} · "
    "비교 대상: <b>LF몰 · W컨셉 · 한섬 · SSF샵 · SI빌리지</b></div>",
    unsafe_allow_html=True,
)
st.write("")

# ── KPI: LF몰 현황 ──
lf = df_ov[df_ov["사"] == "LF몰"].iloc[0]
lf_traffic_rank = int(df_ov["월 오가닉 트래픽"].rank(ascending=False)[df_ov["사"] == "LF몰"].iloc[0])
lf_kw_rank = int(df_ov["오가닉 키워드"].rank(ascending=False)[df_ov["사"] == "LF몰"].iloc[0])
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("LF몰 오가닉 키워드", f"{lf['오가닉 키워드']:,}", f"5사 중 {lf_kw_rank}위 (최다)")
k2.metric("LF몰 월 오가닉 트래픽", f"{lf['월 오가닉 트래픽']:,}", f"5사 중 {lf_traffic_rank}위")
k3.metric("키워드당 트래픽", f"{lf['키워드당 트래픽']}", "효율 최하위", delta_color="inverse")
k4.metric("미보유 대형 키워드", f"{len(GAP_MISSING)}개", "갭 분석 기회")
k5.metric("퀵윈 후보", f"{len(QUICK_WINS)}개", "5위권 밖 보유 KW")

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🔍 경쟁사 SEO 분석")
    st.caption(f"Semrush {DATABASE} DB · {SNAPSHOT_DATE}")
    st.markdown("---")
    st.markdown("**비교 대상 5사**")
    for _, r in df_ov.iterrows():
        dot = SITE_COLOR.get(r["사"], PALETTE["slate"])
        st.markdown(
            f"<span style='color:{dot};font-size:18px'>●</span> "
            f"<b>{r['사']}</b> <span class='cap'>{r['도메인']}</span>",
            unsafe_allow_html=True,
        )
    st.markdown("---")
    st.caption(
        "데이터는 정적 스냅샷입니다. 최신화하려면 Semrush MCP로 "
        "`domain_rank` / `domain_organic` / `domain_domains` 리포트를 "
        "다시 뽑아 코드 상단 SNAPSHOT 값을 교체하세요."
    )

# ══════════════════════════════════════════════════════
# 탭 구성
# ══════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["📊 체급 비교", "🏷️ 사별 상위 키워드", "🎯 키워드 갭", "🚀 LF 퀵윈", "📋 진단 요약"]
)

# ──────────────────────────────────────────────────────
# TAB 1 — 체급 비교
# ──────────────────────────────────────────────────────
with tab1:
    metric_label = st.radio(
        "비교 지표",
        ["월 오가닉 트래픽", "오가닉 키워드", "트래픽 가치($)", "광고 트래픽"],
        horizontal=True,
    )
    d = df_ov.sort_values(metric_label, ascending=True)
    colors = [SITE_COLOR.get(s, PALETTE["slate"]) for s in d["사"]]

    c1, c2 = st.columns([1.1, 1])
    with c1:
        fig = go.Figure(go.Bar(
            x=d[metric_label], y=d["사"], orientation="h",
            marker=dict(color=colors),
            text=[f"{v:,.0f}" for v in d[metric_label]],
            textposition="outside",
            hovertemplate="%{y}<br>" + metric_label + ": %{x:,.0f}<extra></extra>",
        ))
        fig.update_layout(**base_layout(h=320, title=f"{metric_label} 순위"))
        fig.update_xaxes(range=[0, d[metric_label].max() * 1.18])
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        # 키워드당 트래픽 효율
        d2 = df_ov.sort_values("키워드당 트래픽", ascending=True)
        colors2 = [SITE_COLOR.get(s, PALETTE["slate"]) for s in d2["사"]]
        fig2 = go.Figure(go.Bar(
            x=d2["키워드당 트래픽"], y=d2["사"], orientation="h",
            marker=dict(color=colors2),
            text=[f"{v}" for v in d2["키워드당 트래픽"]], textposition="outside",
            hovertemplate="%{y}<br>키워드당 트래픽: %{x}<extra></extra>",
        ))
        fig2.update_layout(**base_layout(h=320, title="키워드당 트래픽 효율 (높을수록 알짜)"))
        fig2.update_xaxes(range=[0, d2["키워드당 트래픽"].max() * 1.18])
        st.plotly_chart(fig2, use_container_width=True)

    # 효율 버블: 키워드 수 vs 트래픽
    st.markdown("##### 키워드 보유량 × 트래픽 (버블 = 광고 트래픽)")
    fig3 = go.Figure()
    for _, r in df_ov.iterrows():
        fig3.add_trace(go.Scatter(
            x=[r["오가닉 키워드"]], y=[r["월 오가닉 트래픽"]],
            mode="markers+text",
            marker=dict(
                size=max(18, (r["광고 트래픽"] ** 0.5) / 2 + 18),
                color=SITE_COLOR.get(r["사"], PALETTE["slate"]),
                opacity=0.85, line=dict(color="white", width=1.5),
            ),
            text=[r["사"]], textposition="top center",
            textfont=dict(size=12, color="#334155"),
            hovertemplate=(f"<b>{r['사']}</b><br>오가닉 키워드: {r['오가닉 키워드']:,}"
                           f"<br>월 트래픽: {r['월 오가닉 트래픽']:,}"
                           f"<br>광고 트래픽: {r['광고 트래픽']:,}<extra></extra>"),
        ))
    fig3.update_layout(**base_layout(h=380))
    fig3.update_xaxes(title="오가닉 키워드 수")
    fig3.update_yaxes(title="월 오가닉 트래픽")
    st.plotly_chart(fig3, use_container_width=True)
    st.caption(
        "💡 우상단일수록 강함. **LF몰은 키워드 수(x)는 최다지만 트래픽(y)은 중위** → "
        "걸린 키워드 대비 순위가 낮다는 신호. W컨셉은 적은 키워드로 높은 트래픽(효율왕)."
    )

    st.markdown("##### 전체 지표 테이블")
    show = df_ov.drop(columns=["도메인랭크"]).copy()
    st.dataframe(
        show, use_container_width=True, hide_index=True,
        column_config={
            "월 오가닉 트래픽": st.column_config.ProgressColumn(
                "월 오가닉 트래픽", format="%d",
                min_value=0, max_value=int(df_ov["월 오가닉 트래픽"].max())),
            "오가닉 키워드": st.column_config.ProgressColumn(
                "오가닉 키워드", format="%d",
                min_value=0, max_value=int(df_ov["오가닉 키워드"].max())),
        },
    )

# ──────────────────────────────────────────────────────
# TAB 2 — 사별 상위 키워드
# ──────────────────────────────────────────────────────
with tab2:
    site = st.selectbox("사이트 선택", list(TOP_KEYWORDS.keys()))
    kdf = pd.DataFrame(TOP_KEYWORDS[site], columns=["키워드", "순위", "검색량", "트래픽 기여%"])
    color = SITE_COLOR.get(site, PALETTE["slate"])

    c1, c2 = st.columns([1, 1])
    with c1:
        dd = kdf.sort_values("트래픽 기여%", ascending=True)
        fig = go.Figure(go.Bar(
            x=dd["트래픽 기여%"], y=dd["키워드"], orientation="h",
            marker=dict(color=color),
            text=[f"{v:.1f}%" for v in dd["트래픽 기여%"]], textposition="outside",
            hovertemplate="%{y}<br>트래픽 기여: %{x:.2f}%<extra></extra>",
        ))
        fig.update_layout(**base_layout(h=420, title=f"{site} · 트래픽 기여 상위 키워드"))
        fig.update_xaxes(range=[0, dd["트래픽 기여%"].max() * 1.2], ticksuffix="%")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown(f"<b style='color:{color}'>{site}</b> 상위 키워드 상세",
                    unsafe_allow_html=True)
        st.dataframe(
            kdf, use_container_width=True, hide_index=True, height=420,
            column_config={
                "검색량": st.column_config.NumberColumn("검색량", format="%d"),
                "트래픽 기여%": st.column_config.ProgressColumn(
                    "트래픽 기여%", format="%.1f%%", min_value=0,
                    max_value=float(kdf["트래픽 기여%"].max())),
            },
        )
    if site == "W컨셉":
        st.caption("⚠️ '늑대 닷컴*' 등은 불법 웹툰 스팸성 키워드로 실제 패션 트래픽이 아님 "
                   "→ W컨셉 트래픽 수치에는 거품이 일부 포함.")

# ──────────────────────────────────────────────────────
# TAB 3 — 키워드 갭 (LF 미보유)
# ──────────────────────────────────────────────────────
with tab3:
    st.markdown("##### LF몰이 놓치고 있는 고검색량 키워드")
    st.caption("경쟁사는 상위 100위 내 노출되지만 LF몰은 순위 0인 키워드. 카테고리/콘텐츠 페이지 기회.")
    gdf = pd.DataFrame(GAP_MISSING, columns=["키워드", "검색량", "보유사", "보유사 순위"])

    c1, c2 = st.columns([1.2, 1])
    with c1:
        dd = gdf.sort_values("검색량", ascending=True)
        bar_colors = [SITE_COLOR.get(s, PALETTE["slate"]) for s in dd["보유사"]]
        fig = go.Figure(go.Bar(
            x=dd["검색량"], y=dd["키워드"], orientation="h",
            marker=dict(color=bar_colors),
            text=[f"{v:,}" for v in dd["검색량"]], textposition="outside",
            customdata=dd[["보유사", "보유사 순위"]],
            hovertemplate="%{y}<br>검색량: %{x:,}<br>보유: %{customdata[0]} (%{customdata[1]}위)<extra></extra>",
        ))
        fig.update_layout(**base_layout(h=460, title="미보유 키워드 (막대색 = 보유 경쟁사)"))
        fig.update_xaxes(range=[0, dd["검색량"].max() * 1.18])
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(
            gdf, use_container_width=True, hide_index=True, height=460,
            column_config={
                "검색량": st.column_config.ProgressColumn(
                    "검색량", format="%d", min_value=0,
                    max_value=int(gdf["검색량"].max())),
            },
        )
        owners = gdf["보유사"].value_counts()
        top_owner = owners.index[0]
        st.caption(f"미보유 키워드를 가장 많이 선점한 곳: **{top_owner}** "
                   f"({owners.iloc[0]}개). 대부분 입점 브랜드/일반 카테고리 키워드.")

# ──────────────────────────────────────────────────────
# TAB 4 — LF 퀵윈
# ──────────────────────────────────────────────────────
with tab4:
    st.markdown("##### LF몰 퀵윈 키워드 (5위권 밖 보유 → 상위 진입 우선순위)")
    st.caption("이미 순위는 잡혀 있으나 5위 밖이라 트래픽을 거의 못 가져오는 고검색량 키워드. "
               "신규 발굴보다 ROI가 빠른 최우선 개선 대상.")
    qdf = pd.DataFrame(QUICK_WINS, columns=["키워드", "현재 순위", "검색량"])
    # 잠재 점수: 검색량이 크고 순위가 낮을수록(=개선여지 큼) 우선
    qdf["기회 점수"] = (qdf["검색량"] * (qdf["현재 순위"].clip(lower=1) / 10)).round(0).astype(int)
    qdf = qdf.sort_values("기회 점수", ascending=False)

    c1, c2 = st.columns([1.2, 1])
    with c1:
        dd = qdf.sort_values("기회 점수", ascending=True)
        # 순위가 낮을수록(페이지2) 진하게
        fig = go.Figure(go.Bar(
            x=dd["기회 점수"], y=dd["키워드"], orientation="h",
            marker=dict(color=dd["현재 순위"], colorscale="Blues", showscale=True,
                        colorbar=dict(title="현재<br>순위", thickness=12, len=0.6)),
            customdata=dd[["현재 순위", "검색량"]],
            hovertemplate="%{y}<br>현재 %{customdata[0]}위 · 검색량 %{customdata[1]:,}<extra></extra>",
        ))
        fig.update_layout(**base_layout(h=460, title="기회 점수 = 검색량 × 순위 갭"))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(
            qdf, use_container_width=True, hide_index=True, height=460,
            column_config={
                "검색량": st.column_config.NumberColumn("검색량", format="%d"),
                "기회 점수": st.column_config.ProgressColumn(
                    "기회 점수", format="%d", min_value=0,
                    max_value=int(qdf["기회 점수"].max())),
            },
        )
    st.caption("💡 특히 '팅커벨'(검색량 60,500·현재 9위), '하프클럽'(40,500·14위)은 "
               "1페이지(상위 10위) 진입만으로 트래픽이 크게 늘어날 여지가 큼.")

# ──────────────────────────────────────────────────────
# TAB 5 — 진단 요약
# ──────────────────────────────────────────────────────
with tab5:
    st.markdown("##### 한 줄 결론")
    st.markdown(
        "<div class='card'>🏆 <b>SSF샵</b>이 SEO·PPC 리더, <b>W컨셉</b>이 효율왕, "
        "<b>LF몰</b>은 <b>키워드는 최다인데 순위가 낮은 '잠자는 거인'</b>.<br>"
        "→ LF 최우선 과제는 신규 키워드 발굴이 아니라 <b>이미 5~20위에 걸린 키워드를 "
        "상위로 끌어올리는 것</b> (퀵윈 탭).</div>",
        unsafe_allow_html=True,
    )

    cards = [
        ("SSF샵", "red", "SEO·PPC 최강",
         "트래픽 1위(173K). 입점 브랜드(에잇세컨즈·빈폴·구호·비이커·띠어리)마다 1위 다수 "
         "확보 + 일반 키워드(아디다스·청바지·바지·코트·화장품)까지 침투. 광고비도 압도적."),
        ("W컨셉", "amber", "적은 키워드로 고효율",
         "6천 키워드로 14.7만 트래픽(효율 최고). 브랜드명 + 여성 카테고리(비키니·니트·가방) 장악. "
         "단 '늑대닷컴' 등 불법웹툰 스팸 키워드 트래픽이 섞여 실제 체급엔 거품."),
        ("LF몰", "blue", "키워드 1위 · 트래픽 3위",
         "21,881개로 키워드 최다이나 트래픽은 3위 = 순위가 하단. 대표 키워드 '팅커벨'(60,500)도 "
         "9위에 머물러 트래픽 미흡. 상위 진입 최적화 시 업사이드가 가장 큼."),
        ("한섬", "purple", "브랜드 의존형",
         "트래픽 대부분이 자사·취급 브랜드명(한섬·무스너클·타임·DKNY·클럽모나코). "
         "일반 카테고리 키워드 거의 없음. 명품/디자이너 SEO 집중."),
        ("SI빌리지", "slate", "최약체 · 추월 1순위",
         "키워드 991개, PPC 0. 브랜드명(sivillage·자주·어그)에만 의존. SEO 투자 거의 없어 "
         "LF가 가장 쉽게 앞설 수 있는 상대."),
    ]
    for name, col, headline, body in cards:
        c = PALETTE[col]
        st.markdown(
            f"<div class='card'>"
            f"<span class='tag' style='background:{c};color:white'>{name}</span>"
            f"<b>{headline}</b><br><span style='color:#475569'>{body}</span></div>",
            unsafe_allow_html=True,
        )

    st.markdown("##### 추천 액션")
    st.markdown(
        "<div class='card'>"
        "1️⃣ <b>퀵윈 우선</b> — 팅커벨·하프클럽·크롬하츠 등 5~20위 키워드 상위 진입<br>"
        "2️⃣ <b>카테고리 갭 메우기</b> — 아디다스·청바지·가방·패딩 등 미보유 대형 키워드 페이지 강화<br>"
        "3️⃣ <b>SI빌리지부터 추월</b> — 가장 약한 경쟁사 대상으로 가시성 역전<br>"
        "4️⃣ <b>주간 순위 추적</b> — Semrush Position Tracking으로 위 키워드 모니터링"
        "</div>",
        unsafe_allow_html=True,
    )

st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
st.caption(f"데이터: Semrush {DATABASE} DB · 스냅샷 {SNAPSHOT_DATE} · "
           "지표는 Semrush 추정치이며 실제 검색 트래픽과 차이가 있을 수 있음.")
