"""LF몰 SEO 통합 대시보드

사이드바에서 3개 뷰 전환:
  1) 경쟁사 분석   — LF몰 vs W컨셉·한섬·SSF샵·SI빌리지 (패션 62개, Semrush)
  2) 키워드 리서치 — 전체 카테고리 키워드 (Semrush 검색량) + 엑셀 다운로드
  3) 네이버 쇼핑   — 네이버 검색광고/데이터랩 지표 (구글 vs 네이버 비교)
  4) CEP 키워드    — 카테고리 진입점(상황·맥락) 수요조사 (pSEO 설계용)

데이터:
  data/lfmall_keyword_research.csv     (키워드 리서치)
  data/naver_keyword_metrics.csv       (네이버, 있으면 자동 표시)
  경쟁사 62개는 본 파일에 내장

실행: streamlit run lf_seo_dashboard.py
"""

import io
import os
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="LF몰 SEO 대시보드", page_icon="🧭",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.card{border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin:8px 0;line-height:1.6;background:#ffffff}
.sdiv{border-top:1px solid #e2e8f0;margin:20px 0}
.tag{display:inline-block;font-size:11px;font-weight:600;padding:2px 8px;border-radius:6px;margin-right:6px;color:white}
.cap{color:#64748b;font-size:13px}
h1,h2,h3{color:#1e293b}
</style>
""", unsafe_allow_html=True)

# ── 공통 ──
STATUS_COLOR = {"Strong": "#48bb78", "Weak": "#ed8936", "Missing": "#f56565",
                "공백": "#4f8fff", "미수집": "#9ca3af"}
PALETTE = {"blue": "#4f8fff", "red": "#f56565", "amber": "#ed8936",
           "green": "#48bb78", "purple": "#9f7aea", "slate": "#64748b"}
SITE_COLOR = {"LF몰": PALETTE["blue"], "W컨셉": PALETTE["amber"], "한섬": PALETTE["purple"],
              "SSF샵": PALETTE["red"], "SI빌리지": PALETTE["slate"]}
SITES = ["LF몰", "W컨셉", "한섬", "SSF샵", "SI빌리지"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# 섹션이 20개+라 Plotly 기본 10색으론 색이 돌고 돌아 여러 섹션이 같은 색으로 겹친다.
# 24색 팔레트로 섹션마다 고유색을 명시 배정(정렬 기준이라 같은 섹션은 항상 같은 색). 기타는 중립 회색.
_SECTION_PALETTE = px.colors.qualitative.Light24


def section_colors(sections):
    secs = sorted({str(s) for s in sections if str(s) and str(s) != "기타"})
    cmap = {s: _SECTION_PALETTE[i % len(_SECTION_PALETTE)] for i, s in enumerate(secs)}
    cmap["기타"] = "#cbd5e1"
    return cmap


def base_layout(h=320, title="", showlegend=False):
    return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#475569", size=12), margin=dict(l=10, r=10, t=42, b=10),
                height=h, showlegend=showlegend,
                title=dict(text=title, font=dict(color="#64748b", size=13)),
                legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1,
                            font=dict(size=11)),
                xaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b", size=11)),
                yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b", size=11)))


def to_excel(d, sheet="data"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        d.to_excel(w, index=False, sheet_name=sheet)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════
# VIEW 1 — 경쟁사 분석 (패션 62개)
# ══════════════════════════════════════════════════════════════════
SUBJECT, COMPETITORS = "LF몰", ["W컨셉", "한섬", "SSF샵", "SI빌리지"]
COMP_DATA = [
    ("패딩", 33100, 21, 0, 11, 0, 0, 0, "아우터"), ("코트", 27100, 23, 0, 0, 0, 18, 0, "아우터"),
    ("재킷", 22200, 32, 0, 0, 49, 8, 0, "아우터"), ("모피", 22200, 29, 5, 0, 0, 0, 0, "아우터"),
    ("가디건", 12100, 20, 0, 0, 0, 14, 0, "아우터"), ("무스탕", 12100, 20, 13, 2, 0, 0, 0, "아우터"),
    ("트렌치 코트", 12100, 27, 0, 0, 0, 7, 0, "아우터"), ("경량 패딩", 12100, 16, 5, 0, 0, 20, 0, "아우터"),
    ("바람막이", 14800, 19, 0, 0, 0, 20, 0, "아우터"), ("블레이저", 12100, 28, 0, 0, 0, 29, 0, "아우터"),
    ("블루종", 6600, 28, 15, 5, 65, 9, 0, "아우터"),
    ("블라우스", 22200, 24, 0, 2, 0, 16, 0, "상의"), ("셔츠", 22200, 29, 0, 23, 0, 24, 0, "상의"),
    ("니트", 18100, 19, 43, 2, 0, 8, 0, "상의"), ("시스루", 18100, 27, 0, 0, 0, 9, 0, "상의"),
    ("티셔츠", 12100, 16, 0, 5, 0, 32, 0, "상의"), ("후드 집업", 12100, 14, 0, 5, 0, 23, 0, "상의"),
    ("크롭 티", 9900, 17, 32, 2, 0, 3, 0, "상의"), ("탱크 탑", 8100, 23, 0, 0, 0, 21, 0, "상의"),
    ("오프 숄더", 8100, 17, 0, 3, 0, 0, 0, "상의"),
    ("청바지", 74000, 23, 0, 0, 0, 27, 0, "하의"), ("바지", 40500, 19, 0, 0, 0, 14, 0, "하의"),
    ("돌핀 팬츠", 22200, 27, 0, 0, 0, 5, 0, "하의"), ("데님", 22200, 22, 0, 0, 0, 46, 0, "하의"),
    ("미니 스커트", 14800, 31, 18, 0, 0, 19, 0, "하의"), ("버뮤다 팬츠", 14800, 18, 20, 0, 0, 24, 0, "하의"),
    ("슬랙스", 14800, 16, 0, 14, 0, 0, 0, "하의"), ("치마", 14800, 24, 0, 0, 0, 43, 0, "하의"),
    ("반바지", 8100, 16, 0, 0, 0, 24, 0, "하의"),
    ("드레스", 22200, 31, 35, 0, 0, 0, 0, "드레스/수영"), ("비키니", 60500, 35, 0, 5, 0, 0, 0, "드레스/수영"),
    ("모노 키니", 12100, 12, 23, 2, 0, 0, 0, "드레스/수영"),
    ("부츠", 12100, 26, 0, 3, 0, 48, 0, "신발"), ("구두", 9900, 39, 0, 0, 0, 30, 0, "신발"),
    ("로퍼", 12100, 20, 0, 19, 0, 0, 0, "신발"), ("샌들", 6600, 21, 30, 11, 0, 0, 0, "신발"),
    ("더비 슈즈", 8100, 17, 10, 6, 0, 0, 0, "신발"),
    ("가방", 33100, 35, 0, 5, 0, 0, 0, "가방"), ("크로스 백", 12100, 16, 0, 2, 0, 0, 0, "가방"),
    ("백팩", 9900, 20, 0, 0, 0, 24, 0, "가방"), ("토트 백", 9900, 19, 0, 2, 0, 46, 0, "가방"),
    ("메신저 백", 6600, 23, 0, 20, 0, 29, 0, "가방"), ("에코 백", 8100, 10, 26, 3, 0, 0, 0, "가방"),
    ("모자", 18100, 18, 0, 0, 0, 12, 0, "액세서리"), ("지갑", 14800, 20, 0, 0, 0, 30, 0, "액세서리"),
    ("키링", 14800, 18, 0, 0, 0, 49, 0, "액세서리"), ("넥타이", 9900, 32, 0, 0, 0, 21, 0, "액세서리"),
    ("벨트", 6600, 26, 0, 0, 67, 0, 0, "액세서리"), ("뿔테 안경", 9900, 20, 17, 0, 0, 19, 0, "액세서리"),
    ("반다나", 9900, 23, 12, 6, 0, 24, 0, "액세서리"), ("목도리", 8100, 20, 8, 3, 0, 0, 0, "액세서리"),
    ("비니", 8100, 27, 0, 0, 0, 23, 0, "액세서리"), ("카드 지갑", 8100, 19, 23, 11, 0, 42, 0, "액세서리"),
    ("니삭스", 6600, 31, 0, 2, 0, 12, 0, "액세서리"),
    ("속옷", 18100, 23, 0, 0, 0, 19, 0, "언더웨어"), ("여자 속옷", 8100, 14, 9, 26, 0, 0, 0, "언더웨어"),
    ("여자 팬티", 9900, 20, 2, 0, 0, 0, 0, "언더웨어"), ("티 팬티", 12100, 29, 23, 0, 0, 19, 0, "언더웨어"),
    ("잠옷", 8100, 19, 19, 0, 0, 0, 0, "언더웨어"),
    ("정장", 18100, 20, 0, 0, 26, 15, 0, "정장/포멀"), ("세미 정장", 6600, 21, 41, 0, 0, 0, 0, "정장/포멀"),
    ("버건디", 9900, 29, 0, 7, 0, 0, 0, "컬러"),
]


@st.cache_data
def comp_df():
    df = pd.DataFrame(COMP_DATA, columns=["키워드", "MSV", "KD"] + SITES + ["카테고리"])

    def status(r):
        if r[SUBJECT] == 0:
            return "Missing"
        rk = [r[c] for c in COMPETITORS if r[c] > 0]
        return "Strong" if (not rk or r[SUBJECT] < min(rk)) else "Weak"

    def tier(r):
        s, kd, msv, gw = r["Status"], r["KD"], r["MSV"], r[SUBJECT]
        if s == "Missing":
            if kd <= 25 and msv >= 10000:
                return "🥇 즉시 선점"
            return "🥈 선점 후보" if (kd <= 25 or msv >= 10000) else "🥉 롱테일 선점"
        if s == "Weak":
            return "🔧 최적화(탈환)"
        return "🔧 최적화(순위↑)" if (s == "Strong" and gw > 10) else "✅ 방어"

    def prio(r):
        sf = {"Missing": 1.0, "Weak": 0.65}.get(
            r["Status"], 0.45 if r[SUBJECT] > 10 else 0.15)
        return int(round(r["MSV"] * (100 - r["KD"]) / 100 * sf))

    df["Status"] = df.apply(status, axis=1)
    df["타겟 등급"] = df.apply(tier, axis=1)
    df["우선순위"] = df.apply(prio, axis=1)
    return df


def glossary():
    """모든 뷰 상단에 펼쳐보는 용어 사전. 처음 보는 사람도 이해하도록 상세 설명."""
    with st.expander("📖 용어 설명 — 처음이세요? Status·검색량·우선순위·CEP 뜻풀이 (클릭해서 펼치기)"):
        st.markdown("""
**📊 기본 지표**

| 용어 | 한 줄 뜻 | 자세히 |
|---|---|---|
| **대표검색량** | 한국 실수요 대표값 | **네이버 우선**(없으면 구글). 우선순위·정렬의 기준. 구글에 안 잡히는 키워드도 네이버로 반영 |
| **검색량 (구글·MSV)** | 한 달 평균 구글 검색 수 | Semrush 한국(kr) 구글 DB · 최근 12개월 평균 월간 검색량. MSV = Monthly Search Volume(월간 검색량). ⚠️ 구글의 한국 점유율이 낮아 한글 키워드 상당수가 0(데이터 없음) — 이때 네이버검색량으로 보완 |
| **네이버검색량** | 한 달 네이버 검색 수 | 네이버 검색광고 API · PC+모바일 합산 월간 검색량. 한국은 네이버 비중이 커서 구글보다 클 때가 많음 |
| **난이도 (KD)** | 상위노출 경쟁 강도 (0~100) | KD = Keyword Difficulty. 높을수록 검색 1페이지 진입이 어려움 (80↑ 매우 치열, 30↓ 비교적 쉬움) |
| **섹션** | 키워드를 묶은 카테고리 | 의류·뷰티/향수·신발·가전·리빙/홈 등. 분류 규칙으로 자동 지정 |
| **패션 (Y/N)** | 우선순위 부여 대상인지 | 패션·뷰티·골프 등 핵심 섹션만 Y → 순위를 매기는 대상 |
| **우선순위 (N순위)** | 섹션 안에서의 공략 순서 | 각 카테고리 '내부'에서 1순위 = 가장 먼저 공략할 키워드 |
| **추이** | 최근 검색 흐름 | 네이버 데이터랩 12개월 기준: 급상승 > 상승 > 유지 > 하락 > 급하락 |
| **갭 (네이버−구글)** | 두 검색량의 차이 | 양수가 클수록 한국에서 네이버 실수요가 구글보다 큼 → 네이버 SEO·쇼핑 우선 투자 |
| **CEP** | 카테고리 진입점 | Category Entry Points. "결혼식 하객룩"처럼 상황·맥락으로 카테고리를 떠올리는 검색 |
| **pSEO** | 대량 SEO 페이지 | Programmatic SEO. [CEP]×[카테고리] 조합으로 템플릿 페이지를 대량 자동 생성 |

**🎯 난이도(KD) 점수대 — 낮을수록 상위노출이 쉬움 (Semrush Keyword Difficulty 기준)**

| KD 점수 | 난이도 | 의미 — 검색 1페이지에 들어가려면 |
|---|---|---|
| **0~14** | 매우 쉬움 | 신규/소규모 사이트도 단기간에 진입 가능 |
| **15~29** | 쉬움 | 기본 온페이지 최적화만으로 노려볼 만 |
| **30~49** | 보통 | 양질의 콘텐츠 + 약간의 도메인 권위 필요 |
| **50~69** | 어려움 | 강한 콘텐츠 + 백링크가 있어야 경쟁 가능 |
| **70~84** | 매우 어려움 | 높은 도메인 권위 필수 (대형 사이트 영역) |
| **85~100** | 극히 어려움 | 톱 브랜드·대형 포털이 장악, 사실상 정공법 불가 |

> *pSEO 타겟 선정 기준: **검색량 높고(↑) KD 낮은(↓)** 키워드 = 황금. KD 50 이상은 롱테일([CEP]×[카테고리])로 우회해 경쟁을 낮춘다.*

**🚦 Status (상태) — LF몰이 4대 경쟁사(W컨셉·한섬·SSF샵·SI빌리지) 대비 검색 순위 어디에 있나**

| 값 | 의미 | 무엇을 해야 하나 |
|---|---|---|
| 🟢 **Strong (강점)** | LF몰 순위가 경쟁사보다 높음 | 현 순위 유지·방어 |
| 🟠 **Weak (약점)** | LF몰도 있지만 경쟁사보다 낮음 | 콘텐츠 보강해 순위 끌어올리기 |
| 🔴 **Missing (공략기회)** | 경쟁사는 있는데 LF몰만 없음 | **최우선 선점 대상** — 경쟁사가 이미 먹고 있는 검색어 |
| ⚪ **공백 (화이트스페이스)** | 5사 모두 순위 없음 | 무주공산 — 먼저 만들면 독점 가능 |
| ◽ **미수집** | 아직 5사 SERP 순위를 조사하지 않음 | 추가 수집 필요 (현재 비패션 카테고리 대부분이 여기) |

> *SERP = Search Engine Result Page(검색 결과 페이지). 순위 1~100, 0 또는 빈값은 해당 사이트가 그 키워드로 노출되지 않음을 뜻합니다.*
""")


def render_competitor():
    df = comp_df()
    st.markdown("## 🥊 경쟁사 분석 — LF몰 vs W컨셉·한섬·SSF샵·SI빌리지")
    st.markdown("<div class='cap'>패션 카테고리 키워드 62개 · Semrush 한국(kr) DB · "
                "순위 1~100(0=없음) · Status는 LF몰 기준</div>", unsafe_allow_html=True)
    st.write("")
    glossary()
    n_missing = int((df["Status"] == "Missing").sum())
    prime = df[(df["Status"] == "Missing") & (df["KD"] <= 25) & (df["MSV"] >= 10000)]
    k = st.columns(5)
    k[0].metric("분석 키워드", f"{len(df)}개")
    k[1].metric("🟢 Strong", int((df["Status"] == "Strong").sum()))
    k[2].metric("🟠 Weak", int((df["Status"] == "Weak").sum()))
    k[3].metric("🔴 Missing", n_missing, "경쟁사만 보유")
    k[4].metric("🥇 즉시 선점", f"{len(prime)}개", "고MSV·저난이도")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    t1, t2, t3, t4, t5, t6 = st.tabs(
        ["🎯 기회 매트릭스", "🚀 선점 타겟", "🗂️ 카테고리 클러스터",
         "📊 Status 분석", "🔍 도메인 커버리지", "📋 진단"])

    with t1:
        st.markdown("##### 기회 매트릭스 — 좌상단(고검색량·저난이도)이 황금존")
        fig = go.Figure()
        fig.add_shape(type="rect", x0=0, x1=25, y0=10000, y1=df["MSV"].max() * 1.05,
                      fillcolor="rgba(72,187,120,0.08)", line=dict(width=0), layer="below")
        for s in ["Missing", "Weak", "Strong"]:
            d = df[df["Status"] == s]
            fig.add_trace(go.Scatter(
                x=d["KD"], y=d["MSV"], mode="markers", name=s,
                marker=dict(size=(d["MSV"] / 1800 + 8), color=STATUS_COLOR[s],
                            opacity=0.78, line=dict(color="white", width=1)),
                customdata=d[["키워드", "타겟 등급", SUBJECT]],
                hovertemplate="<b>%{customdata[0]}</b><br>MSV %{y:,} · KD %{x}"
                              "<br>LF몰 %{customdata[2]}위<br>%{customdata[1]}<extra></extra>"))
        fig.update_layout(**base_layout(h=440, showlegend=True))
        fig.update_xaxes(title="Keyword Difficulty (낮을수록 쉬움) →", autorange="reversed")
        fig.update_yaxes(title="검색량(MSV)")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("##### 🥇 즉시 선점 타겟 (Missing·KD≤25·MSV≥10,000)")
        st.dataframe(prime.sort_values("MSV", ascending=False)[
            ["키워드", "카테고리", "MSV", "KD"] + COMPETITORS],
            use_container_width=True, hide_index=True,
            column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")})

    with t2:
        st.markdown("##### 선점 우선순위 (MSV×난이도할인×상태가중)")
        topn = st.slider("상위 N개", 10, 50, 25, step=5, key="comp_topn")
        rk = df.sort_values("우선순위", ascending=False).head(topn)
        c1, c2 = st.columns([1.1, 1])
        with c1:
            dd = rk.sort_values("우선순위", ascending=True)
            fig = go.Figure(go.Bar(x=dd["우선순위"], y=dd["키워드"], orientation="h",
                                   marker=dict(color=[STATUS_COLOR[s] for s in dd["Status"]]),
                                   customdata=dd[["MSV", "KD", "Status", "타겟 등급"]],
                                   hovertemplate="<b>%{y}</b><br>우선순위 %{x:,}<br>"
                                   "MSV %{customdata[0]:,}·KD %{customdata[1]}<br>"
                                   "%{customdata[2]}·%{customdata[3]}<extra></extra>"))
            fig.update_layout(**base_layout(h=max(420, topn * 17), title="우선순위(색=Status)"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(rk[["키워드", "타겟 등급", "MSV", "KD", SUBJECT, "Status", "우선순위"]],
                         use_container_width=True, hide_index=True, height=max(420, topn * 17),
                         column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")})
        tier_order = ["🥇 즉시 선점", "🥈 선점 후보", "🥉 롱테일 선점",
                      "🔧 최적화(탈환)", "🔧 최적화(순위↑)", "✅ 방어"]
        summ = (df.groupby("타겟 등급").agg(키워드수=("키워드", "count"), 총MSV=("MSV", "sum"))
                  .reindex([t for t in tier_order if t in df["타겟 등급"].values]).reset_index())
        st.markdown("##### 타겟 등급별 요약")
        st.dataframe(summ, use_container_width=True, hide_index=True,
                     column_config={"총MSV": st.column_config.NumberColumn("총MSV", format="%d")})

    with t3:
        st.markdown("##### 카테고리 클러스터 — 허브·스포크 페이지 아키텍처")
        st.caption("블록 면적=MSV(로그스케일), 색=Status. 빨간(Missing) 블록이 많은 카테고리 = 선점 여지 큰 클러스터.")
        dt = df.copy()
        dt["_area"] = np.log10(dt["MSV"] + 1)
        fig = px.treemap(dt, path=[px.Constant("전체"), "카테고리", "키워드"], values="_area",
                         color="Status", color_discrete_map=STATUS_COLOR, custom_data=["KD", SUBJECT, "MSV"])
        fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                          hovertemplate="<b>%{label}</b><br>MSV %{customdata[2]:,}<br>KD %{customdata[0]}<extra></extra>")
        fig.update_layout(margin=dict(l=8, r=8, t=10, b=8), height=460)
        st.plotly_chart(fig, use_container_width=True)
        cat = (df.groupby("카테고리").agg(
                   키워드수=("키워드", "count"), 총MSV=("MSV", "sum"), 평균KD=("KD", "mean"),
                   Missing=("Status", lambda s: (s == "Missing").sum()),
                   Weak=("Status", lambda s: (s == "Weak").sum()),
                   Strong=("Status", lambda s: (s == "Strong").sum()))
               .reset_index().sort_values("총MSV", ascending=False))
        cat["평균KD"] = cat["평균KD"].round(0).astype(int)
        cat["선점여지%"] = ((cat["Missing"] + cat["Weak"]) / cat["키워드수"] * 100).round(0).astype(int)
        st.markdown("##### 카테고리별 선점 여지")
        st.dataframe(cat, use_container_width=True, hide_index=True,
                     column_config={
                         "총MSV": st.column_config.ProgressColumn("총MSV", format="%d",
                                                                  min_value=0, max_value=int(cat["총MSV"].max())),
                         "선점여지%": st.column_config.ProgressColumn("선점여지%", format="%d%%",
                                                                    min_value=0, max_value=100)})

    with t4:
        st.markdown("##### Status 분석 (필터 + 다운로드)")
        c0, c1, c2, c3 = st.columns([1.2, 1.2, 1, 1])
        sel_status = c0.multiselect("Status", ["Strong", "Weak", "Missing"],
                                    default=["Missing", "Weak", "Strong"], key="comp_st")
        sel_cat = c1.multiselect("카테고리", sorted(df["카테고리"].unique()), key="comp_cat")
        min_msv = c2.slider("최소 MSV", 0, int(df["MSV"].max()), 0, step=1000, key="comp_msv")
        max_kd = c3.slider("최대 KD", 0, 40, 40, step=1, key="comp_kd")
        fdf = df[df["Status"].isin(sel_status) & (df["MSV"] >= min_msv) & (df["KD"] <= max_kd)]
        if sel_cat:
            fdf = fdf[fdf["카테고리"].isin(sel_cat)]
        cc1, cc2 = st.columns(2)
        sc = df["Status"].value_counts().reindex(["Strong", "Weak", "Missing"]).fillna(0)
        sm = df.groupby("Status")["MSV"].sum().reindex(["Strong", "Weak", "Missing"]).fillna(0)
        with cc1:
            fig = go.Figure(go.Bar(x=sc.index, y=sc.values,
                                   marker_color=[STATUS_COLOR[s] for s in sc.index],
                                   text=sc.values.astype(int), textposition="outside"))
            fig.update_layout(**base_layout(h=260, title="Status별 키워드 수"))
            st.plotly_chart(fig, use_container_width=True)
        with cc2:
            fig = go.Figure(go.Bar(x=sm.index, y=sm.values,
                                   marker_color=[STATUS_COLOR[s] for s in sm.index],
                                   text=[f"{int(v):,}" for v in sm.values], textposition="outside"))
            fig.update_layout(**base_layout(h=260, title="Status별 총 MSV"))
            st.plotly_chart(fig, use_container_width=True)
        st.markdown(f"##### 키워드 상세 ({len(fdf)}개)")
        st.dataframe(fdf[["키워드", "카테고리", "MSV", "KD"] + SITES + ["Status", "타겟 등급"]]
                     .sort_values("MSV", ascending=False),
                     use_container_width=True, hide_index=True, height=420,
                     column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d")})
        st.download_button("⬇️ CSV 다운로드", fdf.to_csv(index=False).encode("utf-8-sig"),
                           "competitor_keywords.csv", "text/csv")

    with t5:
        st.markdown("##### 도메인별 커버리지 (이 62개 패션 키워드 기준)")
        rows = []
        for c in SITES:
            m = df[c] > 0
            rows.append((c, int(m.sum()), int(df[c].between(1, 10).sum()),
                         round(df.loc[m, c].mean(), 1) if m.any() else 0))
        cov = pd.DataFrame(rows, columns=["도메인", "노출 키워드", "1페이지(1~10위)", "평균 순위"])
        c1, c2 = st.columns([1.2, 1])
        with c1:
            fig = go.Figure()
            fig.add_trace(go.Bar(name="노출", x=cov["도메인"], y=cov["노출 키워드"],
                                 marker_color=PALETTE["slate"], text=cov["노출 키워드"],
                                 textposition="outside"))
            fig.add_trace(go.Bar(name="1페이지", x=cov["도메인"], y=cov["1페이지(1~10위)"],
                                 marker_color=PALETTE["green"], text=cov["1페이지(1~10위)"],
                                 textposition="outside"))
            fig.update_layout(**base_layout(h=340, showlegend=True), barmode="group")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(cov, use_container_width=True, hide_index=True, height=230)
            st.caption("SSF샵·W컨셉이 패션 일반 키워드를 가장 넓게 선점. LF몰은 노출이 적어 "
                       "**Missing 선점 여지가 큼**.")
        st.markdown("##### LF몰 vs 최강 경쟁사 — 키워드별 맞대결")
        duel = df.copy()

        def _bestcomp(r):
            rk = [(r[c], c) for c in COMPETITORS if r[c] > 0]
            return min(rk) if rk else (0, "-")

        bc = duel.apply(_bestcomp, axis=1, result_type="expand")
        duel["경쟁사 최고순위"] = bc[0]
        duel["경쟁사"] = bc[1]

        def _win(r):
            lf, b = r[SUBJECT], r["경쟁사 최고순위"]
            if lf > 0 and (b == 0 or lf < b):
                return "LF몰 우위"
            return "경쟁사 우위" if b > 0 else "-"

        duel["우위"] = duel.apply(_win, axis=1)
        w = int((duel["우위"] == "LF몰 우위").sum())
        lose = int((duel["우위"] == "경쟁사 우위").sum())
        st.caption(f"전체 {len(duel)}개 중 LF몰 우위 **{w}개** / 경쟁사 우위 **{lose}개** "
                   "(Missing 다수 → 선점 기회).")
        st.dataframe(duel[["키워드", "MSV", "KD", SUBJECT, "경쟁사", "경쟁사 최고순위", "Status", "우위"]]
                     .sort_values("MSV", ascending=False),
                     use_container_width=True, hide_index=True, height=320,
                     column_config={"MSV": st.column_config.NumberColumn("MSV", format="%d"),
                                    SUBJECT: st.column_config.NumberColumn("LF몰 순위")})

    with t6:
        st.markdown("##### 한 줄 결론")
        st.markdown(
            "<div class='card'>🏆 <b>SSF샵</b>이 SEO 리더, <b>W컨셉</b>이 효율왕, "
            "<b>LF몰</b>은 키워드는 많은데 순위가 낮은 '잠자는 거인'. → 이미 5~20위에 걸린 키워드를 "
            "상위로 끌어올리는 것이 가장 빠른 ROI.</div>", unsafe_allow_html=True)
        cards = [
            ("SSF샵", "red", "SEO·PPC 최강",
             "트래픽 1위. 입점 브랜드(에잇세컨즈·빈폴·구호·비이커·띠어리)마다 1위 다수 + "
             "일반 키워드(아디다스·청바지·바지·코트)까지 침투. 광고비도 압도적."),
            ("W컨셉", "amber", "적은 키워드로 고효율",
             "적은 키워드로 높은 트래픽. 브랜드명 + 여성 카테고리(비키니·니트·가방) 장악. "
             "단 '늑대닷컴' 등 스팸 키워드 트래픽이 섞여 실제 체급엔 거품."),
            ("LF몰", "blue", "키워드 多 · 순위 下",
             "키워드는 많지만 순위가 하단. 이미 5~20위에 걸린 키워드를 상위로 끌어올리면 업사이드 최대."),
            ("한섬", "purple", "브랜드 의존형",
             "트래픽 대부분이 자사·취급 브랜드명(한섬·무스너클·타임·DKNY·클럽모나코). "
             "일반 카테고리 키워드 거의 없음. 명품/디자이너 SEO 집중."),
            ("SI빌리지", "slate", "최약체 · 추월 1순위",
             "키워드·PPC 최소. 브랜드명(sivillage·자주·어그) 위주. LF가 가장 쉽게 앞설 수 있는 상대."),
        ]
        for name, col, headline, body in cards:
            c = PALETTE[col]
            st.markdown(
                f"<div class='card'><span class='tag' style='background:{c}'>{name}</span>"
                f"<b>{headline}</b><br><span style='color:#475569'>{body}</span></div>",
                unsafe_allow_html=True)
        st.markdown("##### 추천 액션")
        st.markdown(
            "<div class='card'>"
            "1️⃣ <b>황금존 Missing 선점</b> — 청바지·바지·코트·셔츠·블라우스·티셔츠·후드집업 등 "
            "고MSV·저KD 미보유 키워드 카테고리 페이지 생성<br>"
            "2️⃣ <b>Weak 탈환</b> — 니트·크롭티·카드지갑 등 순위는 있으나 경쟁사에 밀리는 키워드 강화<br>"
            "3️⃣ <b>Strong 방어·순위↑</b> — 드레스·모피·세미정장 등 10위 밖 Strong 온페이지 최적화<br>"
            "4️⃣ <b>SI빌리지부터 추월</b> + 주간 순위 추적(Position Tracking)</div>",
            unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# VIEW 2 — 키워드 리서치 (전체 카테고리)
# ══════════════════════════════════════════════════════════════════
def load_kw():
    df = pd.read_csv("data/lfmall_keyword_research.csv", encoding="utf-8-sig")
    for col, dv in [("순위", 0), ("우선순위", ""), ("패션", "N"),
                    ("섹션", "기타"), ("Status", "미수집")]:
        if col not in df.columns:
            df[col] = dv
    # 대표검색량 = 네이버 우선(없으면 구글). 구버전 CSV 호환: 없으면 구글검색량으로 대체
    if "대표검색량" not in df.columns:
        df["대표검색량"] = df["검색량"]
    df["대표검색량"] = pd.to_numeric(df["대표검색량"], errors="coerce").fillna(0).astype(int)
    return df


def render_keyword():
    df = load_kw()
    st.markdown(f"## 🔎 키워드 리서치 — 전체 카테고리 {len(df):,}개")
    st.markdown("<div class='cap'>대상 lfmall.co.kr · 경쟁사 W컨셉·한섬·SSF샵·SI빌리지 · "
                "Semrush 한국(kr) DB 검색량 실측</div>", unsafe_allow_html=True)
    st.write("")
    k = st.columns(5)
    k[0].metric("카테고리 키워드", f"{len(df):,}개",
                f"검색량 보유 {int((df['대표검색량'] > 0).sum())}개")
    k[1].metric("총 검색량(대표)", f"{int(df['대표검색량'].sum()):,}", "네이버 우선")
    k[2].metric("섹션", f"{df['섹션'].nunique()}개")
    k[3].metric("🔴 Missing", int((df["Status"] == "Missing").sum()), "경쟁사만 보유")
    k[4].metric("🎯 패션 우선순위", int((df["순위"] > 0).sum()), "1순위~ 부여")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
    glossary()

    t1, t2, t3, t4, t5 = st.tabs(
        ["🚀 우선순위", "🗂️ 섹션 분석", "🎯 경쟁사 Status", "📋 전체 데이터(엑셀)", "📖 가이드"])

    with t1:
        st.markdown("##### 카테고리별 선점 우선순위 (각 카테고리 안에서 1순위 = 최우선)")
        st.caption("패션·뷰티·골프 각 카테고리 '내부'에서 √검색량·난이도·경쟁상태로 1순위부터 부여. "
                   "막대 = 대표검색량(네이버 우선, 없으면 구글). 카테고리를 선택하면 그 안의 순위가 나옵니다.")
        rank_df = df[df["순위"] > 0]
        cats = rank_df.groupby("섹션")["대표검색량"].sum().sort_values(ascending=False).index.tolist()
        c0, c1 = st.columns([1.5, 1])
        sec_sel = c0.selectbox("카테고리 선택", cats, key="kw_secsel")
        topn = c1.slider("상위 N개", 5, 40, 20, step=5, key="kw_topn")
        d = rank_df[rank_df["섹션"] == sec_sel].sort_values("순위").head(topn)
        c1a, c2a = st.columns([1.1, 1])
        with c1a:
            dd = d.sort_values("순위", ascending=False)
            fig = go.Figure(go.Bar(
                x=dd["대표검색량"], y=dd["키워드"], orientation="h",
                marker=dict(color=[STATUS_COLOR.get(s, "#9ca3af") for s in dd["Status"]]),
                text=dd["우선순위"], textposition="outside",
                customdata=dd[["우선순위", "Status"]],
                hovertemplate="<b>%{y}</b><br>%{customdata[0]} · 대표검색량 %{x:,} · "
                              "%{customdata[1]}<extra></extra>"))
            fig.update_layout(**base_layout(h=max(380, len(d) * 24),
                                            title=f"{sec_sel} 카테고리 내 우선순위 (막대=대표검색량)"))
            fig.update_xaxes(range=[0, max(1, dd["대표검색량"].max()) * 1.18])
            st.plotly_chart(fig, use_container_width=True)
        with c2a:
            st.dataframe(d[["우선순위", "키워드", "대표검색량", "검색량", "네이버검색량", "Status"]],
                         use_container_width=True, hide_index=True, height=max(380, len(d) * 24),
                         column_config={c: st.column_config.NumberColumn(c, format="%d")
                                        for c in ["대표검색량", "검색량", "네이버검색량"]})
        st.caption("대표검색량 = 한국 실수요 대표값(네이버 우선) · 검색량 = 구글(Semrush) · 네이버검색량 = 네이버")
        st.markdown("##### 각 카테고리 1순위 (TOP1 모음)")
        top1 = rank_df[rank_df["순위"] == 1].sort_values("대표검색량", ascending=False)
        st.dataframe(top1[["섹션", "키워드", "대표검색량", "검색량", "네이버검색량", "Status"]],
                     use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%d")
                                    for c in ["대표검색량", "검색량", "네이버검색량"]})

    with t2:
        st.markdown("##### 섹션별 검색 수요 & 분포")
        st.caption("총검색량 = 대표검색량(네이버 우선) 합계 — 구글에 안 잡히는 한국 실수요까지 반영")
        sec = (df.groupby("섹션").agg(키워드수=("키워드", "count"), 총검색량=("대표검색량", "sum"))
                 .reset_index().sort_values("총검색량", ascending=False))
        c1, c2 = st.columns(2)
        with c1:
            dd = sec.sort_values("총검색량", ascending=True)
            fig = go.Figure(go.Bar(x=dd["총검색량"], y=dd["섹션"], orientation="h",
                                   marker_color="#4f8fff",
                                   text=[f"{int(v):,}" for v in dd["총검색량"]], textposition="outside"))
            fig.update_layout(**base_layout(h=560, title="섹션별 총 검색량(대표)"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.caption("블록 면적 = 로그스케일(멱법칙 분포라 면적 그대로면 1개가 화면 독차지) · hover = 실제 검색량")
            dt = df[df["대표검색량"] > 0].copy()
            dt["_area"] = np.log10(dt["대표검색량"] + 1)
            fig = px.treemap(dt, path=[px.Constant("전체"), "섹션", "키워드"],
                             values="_area", color="섹션",
                             color_discrete_map=section_colors(dt["섹션"]),
                             custom_data=["대표검색량"])
            fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                              hovertemplate="<b>%{label}</b><br>대표검색량 %{customdata[0]:,.0f}<extra></extra>")
            fig.update_layout(margin=dict(l=4, r=4, t=10, b=4), height=560)
            st.plotly_chart(fig, use_container_width=True)

    with t3:
        st.markdown("##### 경쟁사 대비 Status (순위 확보된 키워드)")
        known = df[df["Status"] != "미수집"]
        c1, c2 = st.columns([1, 1.4])
        with c1:
            sc = known["Status"].value_counts()
            fig = go.Figure(go.Bar(x=sc.index, y=sc.values,
                                   marker_color=[STATUS_COLOR.get(s, "#9ca3af") for s in sc.index],
                                   text=sc.values, textposition="outside"))
            fig.update_layout(**base_layout(h=320, title="Status 분포(순위 확보분)"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.dataframe(known.sort_values("검색량", ascending=False)[
                ["키워드", "섹션", "검색량", "난이도"] + SITES + ["Status"]],
                use_container_width=True, hide_index=True, height=320,
                column_config={"검색량": st.column_config.NumberColumn("검색량", format="%d")})
        st.info(f"비패션 카테고리의 경쟁사 순위는 미수집 상태입니다. 전체 {len(df):,}개 5사 SERP는 추가 수집 필요.")

    with t4:
        st.markdown("##### 전체 키워드 리서치 데이터 (필터 + 다운로드)")
        f0, f1, f2, f3 = st.columns([1.3, 1.3, 1, 1])
        fsec = f0.multiselect("섹션", sorted(df["섹션"].unique()), key="kw_fsec")
        fstat = f1.multiselect("Status", list(STATUS_COLOR.keys()), key="kw_fstat")
        fmsv = f2.slider("최소 대표검색량", 0, int(df["대표검색량"].max()), 0,
                         step=1000, key="kw_fmsv")
        q = f3.text_input("키워드 검색", "", key="kw_q")
        fdf = df[df["대표검색량"] >= fmsv]
        if fsec:
            fdf = fdf[fdf["섹션"].isin(fsec)]
        if fstat:
            fdf = fdf[fdf["Status"].isin(fstat)]
        if q:
            fdf = fdf[fdf["키워드"].str.contains(q, case=False, na=False)]
        fdf = fdf.sort_values("대표검색량", ascending=False)
        st.markdown(f"**{len(fdf):,}개** 키워드")
        d1, d2 = st.columns(2)
        d1.download_button("⬇️ 엑셀(.xlsx)", to_excel(fdf, "키워드리서치"),
                           "lfmall_keyword_research.xlsx", XLSX_MIME, use_container_width=True)
        d2.download_button("⬇️ CSV", fdf.to_csv(index=False).encode("utf-8-sig"),
                           "lfmall_keyword_research.csv", "text/csv", use_container_width=True)
        st.dataframe(fdf[["키워드", "섹션", "대표검색량", "검색량", "네이버검색량", "난이도"]
                         + SITES + ["Status", "우선순위"]],
                     use_container_width=True, hide_index=True, height=520,
                     column_config={c: st.column_config.NumberColumn(c, format="%d")
                                    for c in ["대표검색량", "검색량", "네이버검색량"]})
        st.caption("대표검색량 = 네이버 우선 실수요 · 검색량 = 구글(Semrush) · 네이버검색량 = 네이버. "
                   "구글이 0이어도 네이버 수요가 크면 대표검색량에 반영됩니다.")

    with t5:
        st.markdown("##### ❓ 왜 '검색량(구글)'이 0인 키워드가 많나요?")
        st.markdown(
            "<div class='card'>Semrush는 <b>구글 기준</b>인데 한국은 구글 점유율이 낮고 "
            "한글 키워드 DB 커버리지가 약합니다. 그래서 <b>크로스백·골프공·헤어드라이기</b>처럼 "
            "실제 수요가 큰 키워드도 구글 검색량이 0(데이터 없음)으로 나옵니다.<br>"
            "→ 그래서 <b>네이버 검색광고 API</b>를 붙였고, <b>대표검색량(네이버 우선)</b>으로 "
            "우선순위를 매깁니다. 전체 1,268개 중 <b>약 1,262개</b>가 대표검색량을 보유합니다.</div>",
            unsafe_allow_html=True)
        st.markdown("##### 📅 기간 기준")
        st.markdown("<div class='card'>검색량 = Semrush 한국(kr) DB <b>최근 12개월 평균 월간</b>, "
                    "네이버검색량 = 네이버 검색광고 API <b>최근 1개월</b>. 계절 키워드(패딩·수영복)는 "
                    "성수기 단월이 연평균보다 큼.</div>", unsafe_allow_html=True)
        st.markdown("##### 🧮 연산 기준")
        st.markdown(
            "<div class='card'><code>우선순위 점수 = √검색량 × 난이도할인 × 상태가중</code><br>"
            "• √검색량 — 초고볼륨(뮬·시계·크림) 지배력 완화<br>"
            "• 난이도할인 = (100−KD)/100 (미상은 40 가정)<br>"
            "• 상태가중: Missing 1.0·공백 0.9·미수집 0.65·Weak 0.6·Strong 0.4<br>"
            "• 점수 내림차순 1순위~. 대상 = 패션+뷰티+골프 섹션 & 검색량&gt;0</div>"
            "<div class='card'><b>Status 판정</b> — LF몰=0&amp;경쟁사도 0→공백 / LF몰=0&amp;경쟁사 있음→Missing / "
            "LF몰&lt;경쟁사 최고순위→Strong / 그 외→Weak / 순위데이터 없음→미수집</div>",
            unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# VIEW 3 — 네이버 쇼핑
# ══════════════════════════════════════════════════════════════════
NAVER_CSV = "data/naver_keyword_metrics.csv"
TREND_COLOR = {"급상승": "#e11d48", "상승": "#f97316", "유지": "#94a3b8",
               "하락": "#3b82f6", "급하락": "#1d4ed8", "": "#9ca3af"}


def render_naver():
    st.markdown("## 🛒 네이버 쇼핑·검색 — 구글 vs 네이버 통합")
    st.markdown("<div class='cap'>네이버 검색광고 API(월간 검색량)+데이터랩(12개월 추이)을 "
                "Semrush(구글) 검색량·카테고리(섹션)와 한 화면에 합쳐 비교. "
                "구글로는 안 잡히는 한국 실수요를 보정합니다.</div>", unsafe_allow_html=True)
    st.write("")
    glossary()

    if not os.path.exists(NAVER_CSV):
        st.warning("아직 네이버 데이터가 없습니다. 아래 절차로 키를 넣고 수집하면 이 탭이 활성화됩니다.")
        st.markdown(
            "<div class='card'><b>① 키 발급 (개인계정 가능)</b><br>"
            "• 검색광고 API — searchad.naver.com → 도구 → API 사용 관리 "
            "(<code>API_KEY</code>·<code>SECRET</code>·<code>CUSTOMER_ID</code>)<br>"
            "• 데이터랩/쇼핑인사이트 — developers.naver.com → 애플리케이션 등록 "
            "(<code>CLIENT_ID</code>·<code>CLIENT_SECRET</code>)</div>"
            "<div class='card'><b>② GitHub Secrets 등록 후 Actions 탭에서 수집 실행</b><br>"
            "Actions → '네이버 키워드 데이터 수집' → Run workflow</div>",
            unsafe_allow_html=True)
        return

    nv = pd.read_csv(NAVER_CSV, encoding="utf-8-sig")
    # 키워드 리서치 전체에서 섹션 + 구글검색량 + Status 병합 → 카테고리화 & 구글/네이버 비교
    kw = load_kw()[["키워드", "섹션", "검색량", "패션", "Status"]].rename(
        columns={"검색량": "구글검색량"})
    nv = nv.merge(kw, on="키워드", how="left")
    nv["섹션"] = nv["섹션"].fillna("기타")
    nv["구글검색량"] = pd.to_numeric(nv["구글검색량"], errors="coerce").fillna(0).astype(int)
    nv["네이버검색량"] = pd.to_numeric(nv["네이버검색량"], errors="coerce").fillna(0).astype(int)
    nv = nv.sort_values("네이버검색량", ascending=False)
    has_trend = nv["추이"].astype(str).str.len().gt(0).any() if "추이" in nv else False

    # 섹션(카테고리) 필터 — 경쟁사 분석·키워드 리서치와 동일한 분류 체계
    secs = sorted(nv["섹션"].unique())
    sel_secs = st.multiselect("🗂️ 카테고리(섹션) 필터 — 비우면 전체", secs, key="nv_secs")
    fv = nv[nv["섹션"].isin(sel_secs)] if sel_secs else nv

    k = st.columns(4)
    k[0].metric("수집 키워드", f"{len(fv):,}개")
    k[1].metric("네이버 검색량 합", f"{int(fv['네이버검색량'].sum()):,}")
    k[2].metric("구글 검색량 합", f"{int(fv['구글검색량'].sum()):,}")
    rising = int(fv["추이"].isin(["상승", "급상승"]).sum()) if has_trend else 0
    k[3].metric("📈 상승 키워드", f"{rising}개")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    has_monthly = "월별추이" in fv.columns and fv["월별추이"].astype(str).str.len().gt(0).any()

    def parse_series(s):
        try:
            return [float(x) for x in str(s).split("|") if x != ""]
        except ValueError:
            return []

    t1, t2, t3, t4, t5 = st.tabs(
        ["📈 검색량 TOP", "🆚 구글 vs 네이버", "🗂️ 섹션 분석", "📅 기간 추이", "📋 전체(엑셀)"])

    with t1:
        topn = st.slider("상위 N개", 10, 60, 30, step=5, key="nv_topn")
        d = fv.head(topn).sort_values("네이버검색량")
        fig = go.Figure(go.Bar(x=d["네이버검색량"], y=d["키워드"], orientation="h",
                               marker_color=PALETTE["green"],
                               customdata=d[["PC", "모바일", "섹션"]],
                               hovertemplate="<b>%{y}</b> · %{customdata[2]}<br>네이버 %{x:,}"
                               "<br>PC %{customdata[0]:,}·모바일 %{customdata[1]:,}<extra></extra>"))
        fig.update_layout(**base_layout(h=max(420, topn * 16), title="네이버 월간 검색량 TOP"))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("PC+모바일 합산 월간 검색수. 막대에 마우스를 올리면 소속 카테고리가 보입니다.")

    with t2:
        st.markdown("##### 구글(Semrush) vs 네이버 검색량 — 한국 실수요 갭")
        st.caption("대각선 위쪽(네이버 ≫ 구글) = 한국에서 네이버 수요가 훨씬 큰 키워드. "
                   "LF몰이 네이버 SEO·쇼핑에 우선 투자해야 할 영역입니다.")
        cmp = fv[(fv["네이버검색량"] > 0) | (fv["구글검색량"] > 0)].copy()
        if len(cmp):
            cmp["갭(네이버-구글)"] = cmp["네이버검색량"] - cmp["구글검색량"]
            cmp["_size"] = cmp["네이버검색량"].clip(lower=1)
            # 검색량 분포가 10~수만으로 극단적 → 로그-로그축이라야 점들이 안 뭉친다.
            # 0(데이터 없음)은 로그축에 못 찍으므로 1로 클립(최좌·최하단에 모임).
            cmp["구글_p"] = cmp["구글검색량"].clip(lower=1)
            cmp["네이버_p"] = cmp["네이버검색량"].clip(lower=1)
            fig = px.scatter(cmp, x="구글_p", y="네이버_p", color="섹션",
                             hover_name="키워드", size="_size", size_max=26,
                             log_x=True, log_y=True,
                             custom_data=["구글검색량", "네이버검색량", "섹션"])
            fig.update_traces(hovertemplate="<b>%{hovertext}</b> · %{customdata[2]}<br>"
                              "구글 %{customdata[0]:,} · 네이버 %{customdata[1]:,}<extra></extra>")
            mx = int(max(cmp["구글검색량"].max(), cmp["네이버검색량"].max(), 1))
            fig.add_trace(go.Scatter(x=[1, mx], y=[1, mx], mode="lines",
                                     line=dict(dash="dash", color="#94a3b8"),
                                     showlegend=False, hoverinfo="skip"))
            fig.update_layout(**base_layout(h=460, showlegend=True,
                                            title="구글 vs 네이버 검색량 (로그축, 점=키워드)"))
            fig.update_xaxes(title="구글 검색량 (로그 · 1=데이터 없음)")
            fig.update_yaxes(title="네이버 검색량 (로그)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("##### 네이버 실수요가 구글보다 큰 키워드 TOP")
            gap = cmp.sort_values("갭(네이버-구글)", ascending=False)
            num_cols = ["네이버검색량", "구글검색량", "갭(네이버-구글)"]
            st.dataframe(
                gap[["키워드", "섹션"] + num_cols + ["Status"]].head(40),
                use_container_width=True, hide_index=True, height=360,
                column_config={c: st.column_config.NumberColumn(c, format="%d") for c in num_cols})
        else:
            st.info("비교할 검색량 데이터가 없습니다.")

    with t3:
        st.markdown("##### 섹션(카테고리)별 검색 수요 — 네이버 vs 구글")
        sec = (fv.groupby("섹션").agg(키워드수=("키워드", "count"),
                                     네이버검색량=("네이버검색량", "sum"),
                                     구글검색량=("구글검색량", "sum"))
                 .reset_index().sort_values("네이버검색량", ascending=False))
        c1, c2 = st.columns(2)
        with c1:
            dd = sec.sort_values("네이버검색량")
            fig = go.Figure()
            fig.add_trace(go.Bar(y=dd["섹션"], x=dd["네이버검색량"], orientation="h",
                                 name="네이버", marker_color=PALETTE["green"]))
            fig.add_trace(go.Bar(y=dd["섹션"], x=dd["구글검색량"], orientation="h",
                                 name="구글", marker_color=PALETTE["blue"]))
            fig.update_layout(**base_layout(h=520, showlegend=True,
                                            title="섹션별 검색량 (네이버 vs 구글)"))
            fig.update_layout(barmode="group", legend=dict(orientation="h"))
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            pos = fv[fv["네이버검색량"] > 0].copy()
            if len(pos):
                # 검색량은 멱법칙 분포(크림 등 1개가 수십배) → 면적 그대로면 한 칸이 화면 독차지.
                # 면적은 로그스케일로 완만하게, 실제값은 hover로.
                pos["_area"] = np.log10(pos["네이버검색량"] + 1)
                fig = px.treemap(pos, path=[px.Constant("전체"), "섹션", "키워드"],
                                 values="_area", color="섹션",
                                 color_discrete_map=section_colors(pos["섹션"]),
                                 custom_data=["네이버검색량"])
                fig.update_layout(**base_layout(h=520,
                                  title="섹션 트리맵 (면적=로그스케일 · hover=실제 네이버검색량)"))
                fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                                  hovertemplate="<b>%{label}</b><br>네이버 %{customdata[0]:,.0f}<extra></extra>")
                st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sec, use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%d")
                                    for c in ["네이버검색량", "구글검색량"]})

    with t4:
        if not has_monthly:
            st.info("데이터랩 키(CLIENT_ID/SECRET) 추가 후 재수집하면 **기간 슬라이더 + 추이 라인차트**가 "
                    "활성화됩니다. (검색광고 키만으로는 절대 검색량만 수집)")
        else:
            months = st.slider("📅 기간 (최근 N개월)", 3, 12, 12, step=1, key="nv_months")
            nv2 = fv.copy()
            nv2["_s"] = nv2["월별추이"].map(parse_series)

            def grow(s):
                w = s[-months:]
                if len(w) < 2:
                    return 0
                n = max(1, len(w) // 3)
                e = sum(w[:n]) / n
                l = sum(w[-n:]) / n
                return round((l - e) / e * 100) if e else 0

            nv2["기간성장%"] = nv2["_s"].map(grow)
            up = nv2[nv2["기간성장%"] >= 5].sort_values("기간성장%", ascending=False)
            st.markdown(f"##### 최근 {months}개월 상승 키워드 ({len(up)}개)")
            sel = st.multiselect("추이 비교 키워드", nv2["키워드"].tolist(),
                                 default=up["키워드"].head(6).tolist(), key="nv_sel")
            if sel:
                fig = go.Figure()
                for kw_ in sel:
                    s = nv2.loc[nv2["키워드"] == kw_, "_s"].iloc[0][-months:]
                    fig.add_trace(go.Scatter(y=s, x=list(range(-len(s) + 1, 1)),
                                             mode="lines+markers", name=kw_))
                fig.update_layout(**base_layout(h=360, showlegend=True,
                                                title=f"최근 {months}개월 상대 검색추이(데이터랩 지수)"))
                fig.update_xaxes(title="개월 전 (0 = 최근月)")
                st.plotly_chart(fig, use_container_width=True)
            st.dataframe(up[["키워드", "섹션", "네이버검색량", "기간성장%", "추이지수"]].head(60),
                         use_container_width=True, hide_index=True, height=320,
                         column_config={"네이버검색량": st.column_config.NumberColumn("네이버검색량", format="%d")})

    with t5:
        st.markdown(f"##### 전체 통합 지표 ({len(fv):,}개) — 섹션·구글·네이버")
        st.download_button("⬇️ 엑셀(.xlsx)", to_excel(fv, "네이버지표"),
                           "naver_keyword_metrics.xlsx", XLSX_MIME)
        num_cfg = {c: st.column_config.NumberColumn(c, format="%d")
                   for c in ["네이버검색량", "구글검색량", "PC", "모바일"] if c in fv}
        st.dataframe(fv, use_container_width=True, hide_index=True, height=480,
                     column_config=num_cfg)


# ══════════════════════════════════════════════════════════════════
# VIEW 4 — CEP 키워드 (카테고리 진입점)
# ══════════════════════════════════════════════════════════════════
CEP_CSV = "data/cep_keyword_research.csv"


def render_cep():
    st.markdown("## 🎯 CEP 키워드 — 카테고리 진입점(상황·맥락) 수요조사")
    st.markdown("<div class='cap'>CEP(Category Entry Points) = 소비자가 특정 상황·니즈에서 "
                "카테고리를 떠올리는 검색 맥락(예: 결혼식 하객룩, 캠핑 옷, 여름 휴가룩). "
                "카테고리 키워드와 조합해 롱테일 pSEO 페이지를 설계합니다.</div>",
                unsafe_allow_html=True)
    st.write("")

    st.markdown(
        "<div class='card'><b>📍 pSEO 키워드 전략 로드맵</b> "
        "<span style='color:#888'>(programmatic-seo 스킬 기준 점검 반영)</span><br>"
        "① 카테고리 키워드 수요조사 — <b>✅ 완료</b> (키워드 리서치 탭)<br>"
        "② CEP 키워드 수요조사 — <b>✅ 완료</b> (이 탭)<br>"
        "③ ①·② 고검색량 키워드 중심 조합 ← <b>현재 단계</b><br>"
        "④ [CEP]×[카테고리] 롱테일 수요조사 <b>+ 난이도(KD)·경쟁 검증</b><br>"
        "⑤ pSEO 페이지 제작 — 5단계로 세분화:<br>"
        "&nbsp;&nbsp;5a. 플레이북 확정 (Occasion·Persona + Curation + Examples)<br>"
        "&nbsp;&nbsp;5b. <b>⭐ 고유 데이터 매핑</b> (LF몰 실제 상품 큐레이션 — thin content 방지)<br>"
        "&nbsp;&nbsp;5c. 템플릿 설계 (제목·메타·H구조·스키마)<br>"
        "&nbsp;&nbsp;5d. URL 구조 + 내부링킹 (허브앤스포크)<br>"
        "&nbsp;&nbsp;5e. 인덱싱 전략 (sitemap · thin은 noindex)</div>",
        unsafe_allow_html=True)
    glossary()

    if not os.path.exists(CEP_CSV):
        st.warning("아직 CEP 키워드 데이터가 없습니다. CEP 키워드 목록을 정하면 "
                   "카테고리 키워드와 동일한 방식(Semrush+네이버)으로 수요조사를 채웁니다.")
        st.markdown(
            "<div class='card'><b>CEP 키워드 예시 (상황·맥락 트리거)</b><br>"
            "· 시즌/이벤트: 여름 휴가룩, 가을 코디, 결혼식 하객룩, 졸업식 정장<br>"
            "· 활동/TPO: 캠핑 옷, 등산 복장, 출근룩, 골프웨어, 홈트레이닝복<br>"
            "· 대상/관계: 엄마 선물, 남자친구 선물, 신생아 준비물<br>"
            "· 니즈/속성: 키 커보이는 코디, 여름 시원한 이불, 미니멀 인테리어</div>",
            unsafe_allow_html=True)
        st.info("CEP 키워드 리스트를 주시면 `.cep_keywords.json` 으로 넣고 "
                "Semrush·네이버 수집 → 이 탭에 자동 표시되게 만들겠습니다.")
        return

    df = pd.read_csv(CEP_CSV, encoding="utf-8-sig")
    for col, dv in [("순위", 0), ("우선순위", ""), ("섹션", "기타"),
                    ("Status", "미수집"), ("네이버검색량", 0), ("검색량", 0)]:
        if col not in df.columns:
            df[col] = dv
    df["검색량"] = pd.to_numeric(df["검색량"], errors="coerce").fillna(0).astype(int)
    df["네이버검색량"] = pd.to_numeric(df["네이버검색량"], errors="coerce").fillna(0).astype(int)
    if "대표검색량" not in df.columns:
        df["대표검색량"] = df[["검색량", "네이버검색량"]].max(axis=1)
    df["대표검색량"] = pd.to_numeric(df["대표검색량"], errors="coerce").fillna(0).astype(int)

    naver_ready = int(df["네이버검색량"].sum()) > 0
    if not naver_ready:
        st.info("현재는 **Semrush(구글)** 검색량만 표시됩니다. CEP는 `~룩`·`하객룩` 등 한국 신조어가 많아 "
                "네이버 수집을 붙이면 수요가 크게 늘어납니다 → Actions 탭에서 'CEP 네이버 수집' 실행.")

    k = st.columns(4)
    k[0].metric("CEP 키워드", f"{len(df):,}개")
    k[1].metric("총 대표검색량", f"{int(df['대표검색량'].sum()):,}", "네이버 우선")
    k[2].metric("구글 보유", f"{int((df['검색량'] > 0).sum())}개")
    k[3].metric("CEP 축", f"{df['섹션'].nunique()}개")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)
    glossary()

    t1, t2, t3, t4 = st.tabs(["🚀 축별 우선순위", "📈 검색량 TOP", "🗂️ 섹션 트리맵", "📋 전체(엑셀)"])

    with t1:
        st.caption("각 CEP 축 '내부'에서 대표검색량 순으로 1순위~. 이 우선순위가 ④ [CEP]×[카테고리] "
                   "롱테일 조합의 출발점이 됩니다.")
        axes = df.groupby("섹션")["대표검색량"].sum().sort_values(ascending=False).index.tolist()
        ax_sel = st.selectbox("CEP 축 선택", axes, key="cep_axsel")
        d = df[df["섹션"] == ax_sel].sort_values("순위")
        dd = d.sort_values("대표검색량")
        fig = go.Figure(go.Bar(x=dd["대표검색량"], y=dd["키워드"], orientation="h",
                               marker_color=PALETTE["purple"], text=dd["우선순위"],
                               textposition="outside"))
        fig.update_layout(**base_layout(h=max(360, len(d) * 26),
                                        title=f"{ax_sel} — 대표검색량 순"))
        fig.update_xaxes(range=[0, max(1, dd["대표검색량"].max()) * 1.18])
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(d[["우선순위", "키워드", "대표검색량", "검색량", "네이버검색량"]],
                     use_container_width=True, hide_index=True,
                     column_config={c: st.column_config.NumberColumn(c, format="%d")
                                    for c in ["대표검색량", "검색량", "네이버검색량"]})

    with t2:
        topn = st.slider("상위 N개", 10, 60, 30, step=5, key="cep_topn")
        d = df.sort_values("대표검색량", ascending=False).head(topn).sort_values("대표검색량")
        fig = go.Figure(go.Bar(x=d["대표검색량"], y=d["키워드"], orientation="h",
                               marker_color=PALETTE["purple"],
                               customdata=d[["섹션"]],
                               hovertemplate="<b>%{y}</b> · %{customdata[0]}<br>"
                               "대표검색량 %{x:,}<extra></extra>"))
        fig.update_layout(**base_layout(h=max(420, topn * 16), title="CEP 키워드 대표검색량 TOP"))
        st.plotly_chart(fig, use_container_width=True)
    with t3:
        st.caption("CEP축 > 키워드 계층. 블록 면적 = 로그스케일(멱법칙 분포 대응) · hover = 실제 대표검색량")
        dt = df[df["대표검색량"] > 0].copy()
        if len(dt):
            dt["_area"] = np.log10(dt["대표검색량"] + 1)
            fig = px.treemap(dt, path=[px.Constant("전체"), "섹션", "키워드"],
                             values="_area", color="섹션",
                             color_discrete_map=section_colors(dt["섹션"]),
                             custom_data=["대표검색량"])
            fig.update_traces(marker=dict(line=dict(color="white", width=1)),
                              hovertemplate="<b>%{label}</b><br>대표검색량 %{customdata[0]:,.0f}<extra></extra>")
            fig.update_layout(**base_layout(h=560, title="CEP 섹션 트리맵 (면적=로그스케일)"))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("대표검색량 데이터가 없습니다.")
    with t4:
        st.download_button("⬇️ 엑셀(.xlsx)", to_excel(df, "CEP키워드"),
                           "cep_keyword_research.xlsx", XLSX_MIME)
        st.dataframe(df, use_container_width=True, hide_index=True, height=480,
                     column_config={c: st.column_config.NumberColumn(c, format="%d")
                                    for c in ["대표검색량", "검색량", "네이버검색량"] if c in df})


COMBO_CSV = "data/combo_candidates.csv"


@st.cache_data
def load_combo_df(naver_sig):
    """combo_candidates.csv는 생성물이라 git에서 제외(충돌 방지). 매번 입력으로 재빌드.
    naver_sig(naver_combo_metrics.csv의 mtime+행수)가 바뀌면 캐시 무효화 → 재생성.
    입력(cep/lfmall/naver_combo/.combo_vol)은 git에 있으므로 앱에서 재생성 가능."""
    try:
        import build_combos
        build_combos.main()       # 항상 최신 입력으로 재생성 (네이버 데이터 갱신 반영)
    except Exception as e:
        st.warning(f"조합 데이터 생성 실패: {e}")
    if not os.path.exists(COMBO_CSV):
        return None
    return pd.read_csv(COMBO_CSV, encoding="utf-8-sig")


def _naver_combo_sig():
    """네이버 조합 원본의 변경 시그니처(mtime, 행수) — 데이터 갱신 시 캐시 무효화용."""
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "naver_combo_metrics.csv")
    if not os.path.exists(p):
        return (0, 0)
    return (os.path.getmtime(p), os.path.getsize(p))


def render_combo():
    st.subheader("🔗 [CEP]×[카테고리] 조합 — ③·④단계")
    st.caption("CEP(상황·맥락)×카테고리(상품) 롱테일 후보(③)를 Semrush로 실제 검색량 조회해 검증(④). "
               "**검증됨** = 실제 검색 수요 확인 → pSEO 페이지 우선순위 대상.")

    df = load_combo_df(_naver_combo_sig())
    if df is None:
        st.info("조합 데이터를 생성할 수 없습니다. 입력 CSV(cep/lfmall/naver_combo)를 확인하세요.")
        return

    for c in ["CEP검색량", "카테고리검색량", "조합점수", "우선순위"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    if "검증" not in df.columns:
        df["검증"] = "미조회"
    df["실제검색량"] = pd.to_numeric(df.get("실제검색량"), errors="coerce")
    df["네이버검색량"] = pd.to_numeric(df.get("네이버검색량"), errors="coerce")
    df["KD"] = pd.to_numeric(df.get("KD"), errors="coerce")
    df["대표실제검색량"] = pd.to_numeric(df.get("대표실제검색량"), errors="coerce")
    df["대표실제검색량"] = df["대표실제검색량"].fillna(df["실제검색량"])
    verified = df[df["검증"] == "검증됨"]
    naver_done = df["네이버검색량"].fillna(0).gt(0).any()

    k = st.columns(4)
    k[0].metric("조합 후보", f"{len(df):,}개")
    k[1].metric("④ 검증됨", f"{len(verified):,}개", "실제 수요 확인")
    k[2].metric("검증 대표검색량 합", f"{int(verified['대표실제검색량'].sum()):,}",
                "네이버 우선" if naver_done else "구글")
    k[3].metric("CEP축", f"{df['CEP축'].nunique()}개")
    if not naver_done:
        st.caption("ℹ️ 현재 검증은 **구글(Semrush)** 기준. 네이버 조합 검색량은 "
                   "'네이버 조합키워드 수집(④단계)' 워크플로 실행 후 반영됩니다.")
    st.markdown("<div class='sdiv'></div>", unsafe_allow_html=True)

    has_comp = "네이버경쟁" in df.columns and \
        (df["네이버경쟁"].fillna("").astype(str).str.len() > 0).any()
    if has_comp:
        st.caption("🎯 난이도는 **네이버 경쟁정도**(높음/중간/낮음)로 봅니다 — 네이버 검색광고 API 제공. "
                   "한국 시장 기준이라 Semrush KD(구글)보다 정확합니다.")
    elif df["KD"].notna().sum() == 0:
        st.warning("난이도 지표 없음 — 네이버 조합 수집(④ 워크플로)을 돌리면 **경쟁정도**가 함께 들어옵니다. "
                   "(Semrush KD는 구글 기준 + 유닛 필요라, 한국은 네이버 경쟁정도 권장)")

    t1, t2 = st.tabs(["✅ ④ 검증된 조합", "🌱 ③ 전체 후보"])

    with t1:
        if len(verified) == 0:
            st.info("아직 검증된 조합이 없습니다.")
        else:
            has_atk = "공략점수" in verified.columns and verified["공략점수"].notna().any()
            sort_key = "공략점수" if has_atk else "대표실제검색량"
            d = verified.copy()
            d[sort_key] = pd.to_numeric(d[sort_key], errors="coerce").fillna(0)
            d = d.sort_values(sort_key, ascending=False)
            # 공략 그룹 3분할 — 기준 설명 + 그룹별 개수 + 필터
            if "공략그룹" in d.columns and (d["공략그룹"] != "").any():
                with st.expander("📐 공략 그룹 3분할 기준 — 공략점수 = 대표검색량 × 경쟁가중(낮음1.0·중간0.6·높음0.3)"):
                    st.markdown(
                        "| 그룹 | 기준 | 의미 |\n|---|---|---|\n"
                        "| 🥇 **1군 즉시공략** | 검색량 ≥ 500 & 경쟁 낮음/중간 | 수요 충분 + 경쟁 낮음 = ROI 최고, **최우선 제작** |\n"
                        "| 🥈 **2군 성장공략** | (검색량 ≥ 500 & 경쟁 높음) 또는 100~500 | 수요 크나 경쟁 세거나 중간 수요. 콘텐츠 보강 시 가치 |\n"
                        "| 🥉 **3군 롱테일풀** | 검색량 < 100 | 개별 수요는 작으나 다수. 템플릿 대량생성(AirOps) 후보 |")
                gc = d["공략그룹"].value_counts()
                m = st.columns(3)
                m[0].metric("🥇 1군 즉시공략", f"{int(gc.get('1군 즉시공략', 0))}개", "검색량≥500·경쟁↓")
                m[1].metric("🥈 2군 성장공략", f"{int(gc.get('2군 성장공략', 0))}개", "수요크나 경쟁↑/중간")
                m[2].metric("🥉 3군 롱테일풀", f"{int(gc.get('3군 롱테일풀', 0))}개", "검색량<100·대량생성")
                gsel = st.radio("공략 그룹 필터", ["전체", "1군 즉시공략", "2군 성장공략", "3군 롱테일풀"],
                                horizontal=True, key="combo_grp")
                if gsel != "전체":
                    d = d[d["공략그룹"] == gsel]
            st.caption(f"검색 수요 확인 {len(d)}개 — 공략점수순(검색량 크고 경쟁 낮을수록 위)")
            st.download_button("⬇️ 엑셀(.xlsx)", to_excel(d, "검증조합"),
                               "combo_verified.xlsx", XLSX_MIME, key="combo_dl_v")
            comp_color = {"낮음": PALETTE["green"], "중간": PALETTE["amber"], "높음": PALETTE["red"]}
            dd = d.head(40).sort_values(sort_key)
            colors = ([comp_color.get(c, PALETTE["slate"]) for c in dd["네이버경쟁"]]
                      if "네이버경쟁" in dd.columns else PALETTE["green"])
            fig = go.Figure(go.Bar(
                x=dd[sort_key], y=dd["조합키워드"], orientation="h", marker_color=colors,
                customdata=dd[["CEP축", "대표실제검색량", "네이버경쟁"]] if "네이버경쟁" in dd.columns
                else dd[["CEP축", "대표실제검색량", "CEP축"]],
                hovertemplate="<b>%{y}</b> · %{customdata[0]}<br>"
                "검색량 %{customdata[1]:,} · 경쟁 %{customdata[2]}<extra></extra>"))
            fig.update_layout(**base_layout(h=max(420, len(dd) * 18),
                              title=f"④ 검증 조합 — {sort_key} TOP (🟢낮음 🟠중간 🔴높음)"))
            st.plotly_chart(fig, use_container_width=True)
            cols = ["공략그룹", "조합키워드", "CEP축", "공략점수", "대표실제검색량", "네이버검색량"]
            if "네이버경쟁" in d.columns:
                cols.insert(4, "네이버경쟁")
            cols += ["실제검색량", "CEP", "카테고리"]
            cols = [c for c in cols if c in d.columns]
            st.dataframe(
                d[cols], use_container_width=True, hide_index=True, height=420,
                column_config={c: st.column_config.NumberColumn(c, format="%d")
                               for c in ["공략점수", "대표실제검색량", "실제검색량", "네이버검색량"]})

    with t2:
        c1, c2 = st.columns(2)
        axes = ["전체"] + sorted(df["CEP축"].unique())
        ax = c1.selectbox("CEP축", axes, key="combo_ax")
        secs = ["전체"] + sorted(df["섹션"].unique())
        sec = c2.selectbox("카테고리 섹션", secs, key="combo_sec")
        d = df
        if ax != "전체":
            d = d[d["CEP축"] == ax]
        if sec != "전체":
            d = d[d["섹션"] == sec]
        d = d.sort_values("조합점수", ascending=False)
        st.caption(f"필터 결과 {len(d):,}개 · 조합점수 = √(CEP검색량)×√(카테고리검색량) (예측 점수)")
        st.download_button("⬇️ 엑셀(.xlsx)", to_excel(d, "조합후보"),
                           "combo_candidates.xlsx", XLSX_MIME, key="combo_dl_all")
        st.dataframe(
            d[["우선순위", "조합키워드", "검증", "실제검색량", "CEP축",
               "CEP검색량", "카테고리", "카테고리검색량", "조합점수"]].head(300),
            use_container_width=True, hide_index=True, height=480,
            column_config={c: st.column_config.NumberColumn(c, format="%d")
                           for c in ["실제검색량", "CEP검색량", "카테고리검색량",
                                     "조합점수", "우선순위"]})


# ══════════════════════════════════════════════════════════════════
# 라우팅
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🧭 LF몰 SEO 대시보드")
    view = st.radio("보기 선택",
                    ["🥊 경쟁사 분석", "🔎 키워드 리서치", "🛒 네이버 쇼핑",
                     "🎯 CEP 키워드", "🔗 조합·롱테일"])
    st.markdown("---")
    st.caption("Semrush 한국(kr) DB + 네이버 검색광고/데이터랩. "
               "데이터는 추정치로 실제와 차이가 있을 수 있습니다.")

if view.startswith("🥊"):
    render_competitor()
elif view.startswith("🔎"):
    render_keyword()
elif view.startswith("🛒"):
    render_naver()
elif view.startswith("🎯"):
    render_cep()
else:
    render_combo()
