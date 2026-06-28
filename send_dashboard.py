import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy import stats
import io, json, os, datetime

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="발송 분석 대시보드", page_icon="",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "insights_store.json"

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.vg{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.65;background:#ffffff}
.vr{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.65;background:#ffffff}
.va{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.65;background:#ffffff}
.vb{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.65;background:#ffffff}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.stat-label{font-size:11px;color:#545c6a;margin-bottom:3px;font-weight:500;letter-spacing:.04em;text-transform:uppercase}
.appendix{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;margin-top:12px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상
# ══════════════════════════════════════════════════════
PALETTE = {
    "blue":   ("rgba(79,143,255,1)",  "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)", "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",  "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",  "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)", "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",  "rgba(56,178,172,0.15)"),
    "pink":   ("rgba(236,72,153,1)",  "rgba(236,72,153,0.15)"),
}
def clr(n): return PALETTE[n][0]
def cbg(n): return PALETTE[n][1]

def base_layout(h=280, ysuffix="", title=""):
    return dict(paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
                font=dict(color="#475569",size=11), margin=dict(l=10,r=10,t=36,b=10),
                height=h, showlegend=False, title=dict(text=title,font=dict(color="#c0c8d8",size=13)),
                xaxis=dict(gridcolor="rgba(0,0,0,0)",linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b",size=11)),
                yaxis=dict(gridcolor="#f1f5f9",linecolor="#e2e8f0",
                           tickfont=dict(color="#64748b",size=11),ticksuffix=ysuffix))

# ══════════════════════════════════════════════════════
# 인사이트 저장소
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE,"r",encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(data):
    with open(INSIGHT_FILE,"w",encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "insights" not in st.session_state:
    st.session_state.insights = load_insights()

# ══════════════════════════════════════════════════════
# 통계 헬퍼
# ══════════════════════════════════════════════════════
def linreg(x, y):
    mask = ~np.isnan(x.astype(float)) & ~np.isnan(y.astype(float))
    if mask.sum() < 5: return dict(slope=np.nan, intercept=np.nan, r2=np.nan, p=np.nan)
    sl, ic, r, p, _ = stats.linregress(x[mask], y[mask])
    return dict(slope=sl, intercept=ic, r2=r**2, p=p)

def dow_residual(df, col):
    vals = df[col].values.astype(float)
    res  = vals.copy()
    for d in range(7):
        idx = df["dow"].values == d
        if idx.sum() > 0: res[idx] -= np.nanmean(res[idx])
    return res

def sig_label(p):
    """p값을 누구나 알기 쉬운 한국어로 설명 (p = 이 패턴이 우연일 확률)"""
    if np.isnan(p): return "–"
    if p < 0.001: return "p<0.001 (우연일 확률 0.1% 미만 — 매우 확실)"
    if p < 0.01:  return "p<0.01 (우연일 확률 1% 미만 — 신뢰할 수 있음)"
    if p < 0.05:  return "p<0.05 (우연일 확률 5% 미만 — 유의함)"
    return "유의하지 않음 (우연일 수 있음)"

def r2_label(r2):
    """R² 값을 누구나 알기 쉬운 한국어로 설명 (R² = 패턴이 얼마나 뚜렷한지)"""
    if np.isnan(r2): return "–"
    if r2 >= 0.3:  return f"R²={r2:.3f} → 뚜렷한 경향 (변동의 {r2*100:.0f}% 설명)"
    if r2 >= 0.1:  return f"R²={r2:.3f} → 약한 경향 (변동의 {r2*100:.0f}% 설명)"
    return f"R²={r2:.3f} → 경향 거의 없음 ({r2*100:.0f}% 미만 설명)"

def corr_label(r):
    """상관계수를 누구나 알기 쉬운 한국어로 설명 (두 지표가 얼마나 함께 움직이는지)"""
    if np.isnan(r): return "–"
    ar = abs(r)
    sign = "반대 방향" if r < 0 else "같은 방향"
    if ar >= 0.7:   strength = "매우 강하게"
    elif ar >= 0.5: strength = "강하게"
    elif ar >= 0.3: strength = "보통으로"
    elif ar >= 0.1: strength = "약하게"
    else:           strength = "거의 관계 없이"
    return f"{strength} {sign}으로 움직임"

def pct(a, b):
    return (b-a)/a*100 if a and not np.isnan(a) else 0

# ══════════════════════════════════════════════════════
# 파싱
# ══════════════════════════════════════════════════════
ALL_METRICS = {
    "perSend":     ["인당발송건수"],
    "revenue":     ["거래액"],
    "rps":         ["발송건당거래액"],
    "totalSend":   ["총발송건수"],
    "customers":   ["유니크발송고객수"],
    "ctr":         ["CTR"],
    "uniqueInflow":["유니크유입"],
    "totalInflow": ["총유입"],
    "visitPerPerson":["인당방문횟수"],
    "purchaseCust":["구매고객수"],
    "purchaseCnt": ["구매건수"],
    "purchasePerPerson":["인당구매건수"],
    "avgOrderVal": ["객단가"],
    "unitPrice":   ["건단가"],
    "mRevenue":    ["M당거래액"],
    "pointM":      ["적립M"],
}

METRIC_LABELS = {
    "perSend":"인당 발송 건수","revenue":"거래액","rps":"발송건당 거래액",
    "totalSend":"총 발송 건수","customers":"유니크 발송 고객수","ctr":"CTR",
    "uniqueInflow":"유니크 유입","totalInflow":"총 유입",
    "visitPerPerson":"인당 방문 횟수","purchaseCust":"구매 고객수",
    "purchaseCnt":"구매 건수","purchasePerPerson":"인당 구매 건수",
    "avgOrderVal":"객단가","unitPrice":"건단가",
    "mRevenue":"M당 거래액","pointM":"적립M",
    "purchaseRate":"구매전환율(CR)","rpc":"고객당 매출",
}

METRIC_FORMAT = {
    "perSend":("{:.2f}건","건"), "revenue":("{:.0f}원","억원"),
    "rps":("{:.0f}원","원"), "totalSend":("{:.0f}건","M건"),
    "customers":("{:.0f}명","만명"), "ctr":("{:.2%}","%"),
    "uniqueInflow":("{:.0f}명","만명"), "totalInflow":("{:.0f}명","만명"),
    "visitPerPerson":("{:.2f}회","회"), "purchaseCust":("{:.0f}명","명"),
    "purchaseCnt":("{:.0f}건","건"), "purchasePerPerson":("{:.3f}건","건"),
    "avgOrderVal":("{:.0f}원","원"), "unitPrice":("{:.0f}원","원"),
    "mRevenue":("{:.0f}원","원"), "pointM":("{:.0f}",""),
    "purchaseRate":("{:.3%}","%"), "rpc":("{:.0f}원","원"),
}

@st.cache_data(show_spinner=False)
def parse_xlsx(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    ws  = pd.read_excel(xls, header=None, sheet_name=xls.sheet_names[0])
    date_row   = ws.iloc[1, 2:]
    metric_col = ws.iloc[3:, 0].astype(str).str.strip()

    dates = []
    for v in date_row:
        if pd.isnull(v): continue
        try:
            d = pd.Timestamp("1899-12-30")+pd.Timedelta(days=int(v)) if isinstance(v,(int,float)) else pd.Timestamp(v)
            dates.append(d)
        except: dates.append(pd.NaT)

    def get_m(keywords):
        for kw in keywords:
            match = metric_col[metric_col.str.replace(" ","") == kw.replace(" ","")]
            if not match.empty:
                vals = ws.iloc[match.index[0], 2:2+len(dates)].values
                return pd.to_numeric(vals, errors="coerce")
        return np.full(len(dates), np.nan)

    data = {k: get_m(v) for k, v in ALL_METRICS.items()}
    data["date"] = dates
    df = pd.DataFrame(data).dropna(subset=["perSend","revenue"]).copy()
    df["date"]    = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
    df["t"]       = np.arange(len(df))
    df["dow"]     = df["date"].dt.dayofweek
    df["month"]   = df["date"].dt.to_period("M").astype(str)
    df["quarter"] = df["date"].dt.to_period("Q").astype(str)
    df["year"]    = df["date"].dt.year
    df["week"]    = df["date"].dt.isocalendar().week.astype(int)
    # 파생 지표 (0 나눗셈 방지: 발송고객수 0 → NaN 처리)
    _cust = df["customers"].replace(0, np.nan)
    df["purchaseRate"] = (df["purchaseCust"] / _cust).clip(0, 1)  # 구매고객/발송고객 → [0,1] 범위
    df["rpc"]          = df["revenue"] / _cust
    return df

@st.cache_data(show_spinner=False)
def compute(files_bytes):
    """단일 bytes 또는 bytes 튜플(다중 파일) → 날짜 기준 이어붙이기 후 집계.
    같은 날짜가 여러 파일에 있으면 나중에 업로드한 파일 값을 사용."""
    if isinstance(files_bytes, bytes):
        files_bytes = (files_bytes,)
    dfs = [parse_xlsx(fb) for fb in files_bytes]
    df  = (pd.concat(dfs, ignore_index=True)
             .drop_duplicates(subset="date", keep="last")
             .sort_values("date").reset_index(drop=True))
    df["t"] = np.arange(len(df))  # 병합 후 시계열 인덱스 재생성

    # Aggregations
    monthly   = df.groupby("month",   sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    quarterly = df.groupby("quarter", sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    yearly    = df.groupby("year",    sort=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()

    # Buckets
    BINS = [0,2.0,2.5,3.0,3.5,4.0,4.5,99]
    LBLS = ["~2.0건","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5건+"]
    df["bucket"] = pd.cut(df["perSend"], bins=BINS, labels=LBLS)
    buckets = df.groupby("bucket", observed=True).agg(n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}).reset_index()
    buckets = buckets[buckets["n"] >= 30].reset_index(drop=True)

    # Quintile — 인덱스를 균등 5분할 (나머지 일수 손실 방지, DataFrame 컬럼 접근 유지)
    df_s   = df.sort_values("totalSend").reset_index(drop=True)
    qidx   = np.array_split(np.arange(len(df_s)), 5)
    quintile = pd.DataFrame([
        dict(label=["Q1 최소","Q2","Q3","Q4","Q5 최대"][i],
             n=len(qidx[i]),
             **{k: df_s.iloc[qidx[i]][k].mean() for k in list(ALL_METRICS)+["purchaseRate","rpc"]})
        for i in range(5)
    ])

    # DoW stats
    DOW = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    dow_comp = []
    for d in [0,1,2,3,4]:
        sub = df[df["dow"]==d]
        if len(sub)<20: continue
        med = sub["perSend"].median()
        lo, hi = sub[sub["perSend"]<=med], sub[sub["perSend"]>med]
        dow_comp.append(dict(dow=DOW[d], median=med,
            lowRev=lo["revenue"].mean(), highRev=hi["revenue"].mean(),
            diff=lo["revenue"].mean()-hi["revenue"].mean(),
            lowRps=lo["rps"].mean(), highRps=hi["rps"].mean(),
            lowCtr=lo["ctr"].mean(), highCtr=hi["ctr"].mean()))
    dow_corr = []
    for d in range(7):
        sub = df[df["dow"]==d]
        if len(sub)<20: continue
        dow_corr.append(dict(dow=DOW[d],
            corrRevenue=float(np.corrcoef(sub["perSend"],sub["revenue"])[0,1]),
            corrRps=float(np.corrcoef(sub["perSend"],sub["rps"])[0,1])))

    # Regression
    t = df["t"].values.astype(float)
    reg = {k: linreg(t, dow_residual(df,k) if k not in ["perSend"] else df[k].values)
           for k in ["perSend","ctr","purchaseRate","rps","revenue"]}

    meta = dict(start=str(df["date"].min().date()), end=str(df["date"].max().date()), days=len(df))
    return dict(df=df, monthly=monthly, quarterly=quarterly, yearly=yearly,
                buckets=buckets, quintile=quintile,
                dow_comp=dow_comp, dow_corr=dow_corr, reg=reg, meta=meta)

# ══════════════════════════════════════════════════════
# 동의 현황 파싱
# ══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_consent(file_bytes):
    """SMS/PUSH 동의 현황 엑셀 파싱
    구조: 행1=날짜(1/1~), 행2~13=기존/신규/Total × 수신동의·증감·신규추가·이탈
    """
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
    date_strs = df_raw.iloc[1, 2:].tolist()

    # 시작 연도 자동 감지 (행 0에 연도 마커가 있으면 사용, 없으면 2025 기본값)
    year = 2025
    for v in df_raw.iloc[0, :]:
        try:
            iv = int(float(v))
            if 2020 <= iv <= 2030:
                year = iv; break
        except Exception:
            pass

    # 날짜 파싱 — M/D 형식, 월이 줄어들면(12→1) 연도 롤오버
    dates = []
    prev_month = None
    for ds in date_strs:
        if pd.isnull(ds):
            dates.append(pd.NaT)
            continue
        try:
            mm = int(str(ds).split("/")[0])
            if prev_month is not None and mm < prev_month:
                year += 1  # 연도 전환 (예: 12/31 → 1/1)
            prev_month = mm
            dates.append(pd.Timestamp(f"{year}/{ds}"))
        except:
            dates.append(pd.NaT)

    n = len(dates)

    def get_row(i):
        return pd.to_numeric(df_raw.iloc[i, 2:2+n].values, errors='coerce')

    df = pd.DataFrame({
        'date':           dates,
        'existing_total': get_row(2),   # 기존 수신동의
        'existing_add':   get_row(4),   # 기존 신규추가
        'existing_leave': get_row(5),   # 기존 이탈
        'new_total':      get_row(6),   # 신규 수신동의
        'new_add':        get_row(8),   # 신규 추가
        'new_leave':      get_row(9),   # 신규 이탈
        'total_consent':  get_row(10),  # Total 수신동의
        'total_change':   get_row(11),  # Total 증감
        'total_add':      get_row(12),  # Total 신규추가
        'total_leave':    get_row(13),  # Total 이탈
    })
    df['date'] = pd.to_datetime(df['date'])
    df = df.dropna(subset=['date']).sort_values('date').reset_index(drop=True)
    df['month'] = df['date'].dt.to_period('M').astype(str)
    df['dow']   = df['date'].dt.dayofweek
    return df


def render_consent_page(df_c, channel):
    """PUSH/SMS 동의 현황 분석 페이지 공통 렌더러"""
    if df_c is None:
        st.info(f"왼쪽 사이드바에서 **{channel} 동의 현황 파일**을 업로드하면 분석이 시작됩니다.")
        st.stop()

    start_d = df_c['date'].min().strftime("%Y-%m-%d")
    end_d   = df_c['date'].max().strftime("%Y-%m-%d")
    latest  = df_c.iloc[-1]
    total_net = int(df_c['total_change'].sum())

    st.caption(f"데이터 기간: {start_d} ~ {end_d} · {len(df_c)}일")
    memo_block(f"top_{channel.lower()}")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # ── KPI 카드 ──
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 수신동의 (최신)",
              f"{int(latest['total_consent']):,}명",
              help="가장 최근 날짜 기준 전체 수신동의 고객 수")
    k2.metric("기간 순증",
              f"{total_net:+,}명",
              f"신규추가 {int(df_c['total_add'].sum()):,} / 이탈 {int(df_c['total_leave'].sum()):,}",
              delta_color="normal" if total_net >= 0 else "inverse",
              help="분석 기간 동안의 총 순증 (신규추가 − 이탈)")
    k3.metric("일평균 신규추가",
              f"{df_c['total_add'].mean():.0f}명/일",
              help="하루 평균 새로 수신동의한 고객 수")
    k4.metric("일평균 이탈",
              f"{df_c['total_leave'].mean():.0f}명/일",
              help="하루 평균 수신동의를 취소한 고객 수 (낮을수록 좋음)")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # ── 총 수신동의 추이 ──
    st.subheader("수신동의 총계 추이")
    fig_total = go.Figure()
    fig_total.add_trace(go.Scatter(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=df_c['total_consent'].tolist(),
        mode="lines", line=dict(color=clr("blue"), width=2),
        fill="tozeroy", fillcolor=cbg("blue"), name="Total 수신동의"
    ))
    fig_total.add_trace(go.Scatter(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=df_c['existing_total'].tolist(),
        mode="lines", line=dict(color=clr("teal"), width=1.5, dash="dot"), name="기존 수신동의"
    ))
    fig_total.add_trace(go.Scatter(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=df_c['new_total'].tolist(),
        mode="lines", line=dict(color=clr("amber"), width=1.5, dash="dot"), name="신규 수신동의"
    ))
    layout_t = base_layout(300, title="수신동의 총계 추이")
    layout_t["showlegend"] = True
    layout_t["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    layout_t["xaxis"]["tickangle"] = -45
    layout_t["xaxis"]["nticks"] = 20
    fig_total.update_layout(**layout_t)
    st.plotly_chart(fig_total, use_container_width=True)

    # ── 일별 신규추가 vs 이탈 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("일별 신규추가 vs 이탈")
    st.caption("초록=신규추가(수신동의 증가), 빨강=이탈(수신동의 취소)")
    fig_add = go.Figure()
    fig_add.add_trace(go.Bar(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=df_c['total_add'].tolist(),
        name="신규추가", marker_color=cbg("green"), marker_line_color=clr("green"), marker_line_width=0.5
    ))
    fig_add.add_trace(go.Bar(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=(-df_c['total_leave']).tolist(),
        name="이탈(−)", marker_color=cbg("red"), marker_line_color=clr("red"), marker_line_width=0.5
    ))
    layout_add = base_layout(260, title="일별 신규추가 / 이탈")
    layout_add["barmode"] = "overlay"
    layout_add["showlegend"] = True
    layout_add["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    layout_add["xaxis"]["nticks"] = 20
    layout_add["xaxis"]["tickangle"] = -45
    fig_add.update_layout(**layout_add)
    st.plotly_chart(fig_add, use_container_width=True)

    # ── 순증감 추이 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("일별 순증감")
    st.caption("양수(초록)=그날 수신동의자 수가 증가, 음수(빨강)=감소")
    bar_colors_c = [cbg("green") if v >= 0 else cbg("red") for v in df_c['total_change']]
    bar_line_c   = [clr("green") if v >= 0 else clr("red") for v in df_c['total_change']]
    fig_net = go.Figure(go.Bar(
        x=df_c['date'].dt.strftime("%m/%d").tolist(),
        y=df_c['total_change'].tolist(),
        marker_color=bar_colors_c, marker_line_color=bar_line_c, marker_line_width=0.5,
        name="순증감"
    ))
    layout_net = base_layout(220, title="일별 순증감 (신규추가 − 이탈)")
    layout_net["shapes"] = [dict(type="line", x0=0, x1=1, xref="paper", y0=0, y1=0,
                                  line=dict(color="#94a3b8", width=1, dash="dot"))]
    layout_net["xaxis"]["nticks"] = 20
    layout_net["xaxis"]["tickangle"] = -45
    fig_net.update_layout(**layout_net)
    st.plotly_chart(fig_net, use_container_width=True)

    # ── 월별 집계 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("월별 집계")
    monthly_c = df_c.groupby("month").agg(
        신규추가=("total_add",   "sum"),
        이탈=("total_leave", "sum"),
        순증=("total_change", "sum"),
        일수=("date", "count"),
    ).reset_index()
    monthly_c["이탈"] = monthly_c["이탈"].apply(lambda x: f"{x:,}")
    monthly_c["신규추가"] = monthly_c["신규추가"].apply(lambda x: f"{x:,}")
    monthly_c["순증"] = monthly_c["순증"].apply(lambda x: f"{x:+,}")
    st.dataframe(monthly_c.rename(columns={"month":"월"}), use_container_width=True, hide_index=True)

    # ── 월별 신규추가/이탈 막대 ──
    monthly_raw = df_c.groupby("month").agg(
        add=("total_add","sum"), leave=("total_leave","sum")
    ).reset_index()
    fig_m = go.Figure()
    fig_m.add_trace(go.Bar(x=monthly_raw["month"].tolist(), y=monthly_raw["add"].tolist(),
        name="신규추가", marker_color=cbg("green"), marker_line_color=clr("green"), marker_line_width=1.5))
    fig_m.add_trace(go.Bar(x=monthly_raw["month"].tolist(), y=monthly_raw["leave"].tolist(),
        name="이탈", marker_color=cbg("red"), marker_line_color=clr("red"), marker_line_width=1.5))
    layout_m = base_layout(240, title="월별 신규추가 vs 이탈")
    layout_m["barmode"] = "group"
    layout_m["showlegend"] = True
    layout_m["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    fig_m.update_layout(**layout_m)
    st.plotly_chart(fig_m, use_container_width=True)

    # ── 근거 데이터 ──
    show_appendix(df_c[["date","total_consent","total_add","total_leave","total_change",
                         "existing_total","new_total"]].rename(columns={
        "date":"날짜","total_consent":"총수신동의","total_add":"신규추가",
        "total_leave":"이탈","total_change":"순증감",
        "existing_total":"기존수신동의","new_total":"신규수신동의"
    }), f"{channel}_동의현황")
    memo_block(f"btm_{channel.lower()}")

# ══════════════════════════════════════════════════════
# 텍스트 포맷 헬퍼
# ══════════════════════════════════════════════════════
FMT_DEFAULTS = {"bold": False, "color": "#334155", "size": 13}
SIZE_OPTS    = [11, 12, 13, 14, 15, 16, 18, 20, 24]

def _norm(val, base=None):
    """저장된 텍스트 값을 {text, bold, color, size} 딕셔너리로 정규화 (하위 호환)"""
    d = (base or FMT_DEFAULTS).copy()
    if isinstance(val, str):
        d["text"] = val
        return d
    if isinstance(val, dict):
        merged = d.copy()
        merged.update({k: v for k, v in val.items() if k != "editing"})
        merged.setdefault("text", "")
        return merged
    d["text"] = ""
    return d

def _style(fmt):
    """포맷 딕셔너리 → CSS 인라인 스타일 문자열"""
    w = "700" if fmt.get("bold") else "400"
    c = fmt.get("color", "#334155")
    s = fmt.get("size", 13)
    return f"font-weight:{w};color:{c};font-size:{s}px;line-height:1.7"

def _fmt_ui(key, fmt):
    """편집 모드 포맷 컨트롤 (굵게·색상·크기) 렌더링, (bold, color, size) 반환"""
    fc1, fc2, fc3 = st.columns([1, 2, 2])
    bold = fc1.checkbox("굵게", value=bool(fmt.get("bold", False)), key=f"_ck_b_{key}")
    raw_c = fmt.get("color", "#334155")
    safe_c = raw_c if (isinstance(raw_c, str) and raw_c.startswith("#") and len(raw_c) == 7) else "#334155"
    color = fc2.color_picker("색상", value=safe_c, key=f"_ck_c_{key}")
    sz = fmt.get("size", 13)
    idx = SIZE_OPTS.index(sz) if sz in SIZE_OPTS else SIZE_OPTS.index(13)
    size = fc3.selectbox("크기(px)", SIZE_OPTS, index=idx, key=f"_ck_s_{key}")
    return bold, color, size

# ══════════════════════════════════════════════════════
# 인사이트 에디터
# ══════════════════════════════════════════════════════
def insight_editor(page_key, default_lines):
    """동적 메모 에디터 — 추가/삭제/숨기기/포맷 가능"""
    store = st.session_state.insights
    if page_key not in store:
        store[page_key] = []

    items    = store[page_key]
    hide_key = f"__hide_memo_{page_key}__"
    if hide_key not in st.session_state:
        st.session_state[hide_key] = False

    # 헤더 행: "메모" 텍스트 + 숨기기/보이기 토글
    h1, h2, h3 = st.columns([8, 1, 1])
    h1.markdown("**메모**", unsafe_allow_html=True)
    if h2.button("숨기기" if not st.session_state[hide_key] else "펼치기",
                 key=f"hide_{page_key}", use_container_width=True):
        st.session_state[hide_key] = not st.session_state[hide_key]
        st.rerun()
    if h3.button("+ 추가", key=f"add_{page_key}", use_container_width=True):
        items.append({"text": "", "editing": True, "bold": False, "color": "#334155", "size": 13})
        store[page_key] = items
        st.session_state.insights = store
        save_insights(store)
        st.rerun()

    if not st.session_state[hide_key]:
        if not items:
            st.markdown(
                "<p style='color:#94a3b8;font-size:12px;margin:4px 0 0 2px'>"
                "+ 추가 버튼으로 메모를 입력하세요.</p>",
                unsafe_allow_html=True
            )
        to_delete = []
        for i, raw_item in enumerate(items):
            item = _norm(raw_item) if not isinstance(raw_item, dict) else raw_item
            item.setdefault("bold", False)
            item.setdefault("color", "#334155")
            item.setdefault("size", 13)
            edit_active = item.get("editing", False)
            c1, c2, c3 = st.columns([10, 1, 1])
            if edit_active:
                new_text = c1.text_area(
                    "", item.get("text",""),
                    key=f"memo_ta_{page_key}_{i}",
                    height=72, label_visibility="collapsed",
                    placeholder="내용을 입력하세요."
                )
                bold, color, size = _fmt_ui(f"memo_{page_key}_{i}", item)
                if c2.button("저장", key=f"memo_save_{page_key}_{i}", use_container_width=True):
                    items[i] = {"text": new_text, "editing": False, "bold": bold, "color": color, "size": size}
                    store[page_key] = items
                    st.session_state.insights = store
                    save_insights(store)
                    st.rerun()
                if c3.button("삭제", key=f"memo_del_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)
            else:
                txt = item.get("text","")
                c1.markdown(
                    f"<p style='{_style(item)};margin:2px 0'>{txt}</p>",
                    unsafe_allow_html=True
                )
                if c2.button("편집", key=f"memo_edit_{page_key}_{i}", use_container_width=True):
                    items[i]["editing"] = True
                    store[page_key] = items
                    st.session_state.insights = store
                    save_insights(store)
                    st.rerun()
                if c3.button("삭제", key=f"memo_del2_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)

        for idx in sorted(to_delete, reverse=True):
            items.pop(idx)
        if to_delete:
            store[page_key] = items
            st.session_state.insights = store
            save_insights(store)
            st.rerun()

    store[page_key] = items
    st.session_state.insights = store
    save_insights(store)



# ══════════════════════════════════════════════════════
# 편집 가능 텍스트 시스템
# ══════════════════════════════════════════════════════
if "editable_texts" not in st.session_state:
    st.session_state.editable_texts = load_insights().get("__texts__", {})

TAG_FMT = {
    "h2": {"bold": True,  "color": "#1e293b", "size": 24},
    "h3": {"bold": True,  "color": "#1e293b", "size": 20},
    "h4": {"bold": True,  "color": "#1e293b", "size": 16},
    "p":  {"bold": False, "color": "#334155", "size": 13},
}

def editable_text(key, default, tag="p", style=""):
    """클릭하면 인라인 편집 가능한 텍스트 (글자색·크기·볼드 포맷 지원)"""
    base = TAG_FMT.get(tag, FMT_DEFAULTS).copy()
    texts = st.session_state.editable_texts

    if key not in texts:
        texts[key] = {"text": default, **base}
    item = _norm(texts[key], base)

    edit_key = f"__editing_{key}__"
    if edit_key not in st.session_state:
        st.session_state[edit_key] = False

    if st.session_state[edit_key]:
        col1, col2 = st.columns([10, 1])
        new_val = col1.text_area("", item["text"], key=f"ta_{key}", height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(key, item)
        if col2.button("확인", key=f"save_{key}"):
            new_item = {"text": new_val, "bold": bold, "color": color, "size": size}
            texts[key] = new_item
            st.session_state[edit_key] = False
            all_data = load_insights()
            all_data["__texts__"] = texts
            save_insights(all_data)
            st.session_state.editable_texts = texts
            st.rerun()
    else:
        item = _norm(texts[key], base)
        rendered = item["text"]
        st.markdown(
            f'<{tag} style="{_style(item)};cursor:pointer;{style}" title="클릭하여 편집">{rendered}</{tag}>',
            unsafe_allow_html=True
        )
        if st.button("편집", key=f"edit_{key}", help="편집"):
            st.session_state[edit_key] = True
            st.rerun()
    return item["text"]

def verdict_box(key, default_text, default_color="vg", emoji=""):
    """편집 가능한 텍스트 박스 (글자색·크기·볼드 포맷 지원)"""
    store = st.session_state.insights
    vkey  = f"__verdict_{key}__"
    ekey  = f"__editing_v_{key}__"

    if vkey not in store:
        store[vkey] = {"text": default_text, "bold": False, "color": "#334155", "size": 13}
    item = _norm(store[vkey])

    if ekey not in st.session_state:
        st.session_state[ekey] = False

    if st.session_state[ekey]:
        c1, c2 = st.columns([11, 1])
        new_text = c1.text_area("", item["text"], key=f"vta_{key}", height=100, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"v_{key}", item)
        if c2.button("확인", key=f"vsave_{key}"):
            new_item = {"text": new_text, "bold": bold, "color": color, "size": size}
            store[vkey] = new_item
            st.session_state[ekey] = False
            all_data = load_insights()
            all_data.update(store)
            save_insights(all_data)
            st.session_state.insights = store
            st.rerun()
    else:
        item = _norm(store[vkey])
        txt = item["text"].replace("\n", "<br>")
        col_a, col_b = st.columns([20, 1])
        col_a.markdown(
            f'<div class="vg"><span style="{_style(item)}">{txt}</span></div>',
            unsafe_allow_html=True
        )
        if col_b.button("편집", key=f"vedit_{key}"):
            st.session_state[ekey] = True
            st.rerun()


# ══════════════════════════════════════════════════════
# 기간 비교 헬퍼
# ══════════════════════════════════════════════════════
COMPARE_OPTS = ["전년비 (YoY)", "전분기비 (QoQ)", "전월비 (MoM)", "전주비 (WoW)", "전체 기간 처음↔끝"]

def get_compare_periods(df_all, mode, yoy_ref=None):
    """비교 기준 기간과 현재 기간 반환 → (df_current, df_prev, label_cur, label_prev)"""
    df_all = df_all.sort_values("date")
    last_date = df_all["date"].max()

    if mode == "전년비 (YoY)":
        # 동일 기간 비교: 기준일까지의 올해 vs 전년도 같은 월/일까지
        ref = yoy_ref if yoy_ref is not None else last_date
        cur_year, prev_year = ref.year, ref.year - 1
        cur = df_all[(df_all["date"].dt.year == cur_year) & (df_all["date"] <= ref)]
        try:
            prev_ref = ref.replace(year=prev_year)
        except ValueError:  # 2/29 윤년 보정
            prev_ref = ref.replace(year=prev_year, day=28)
        prev = df_all[(df_all["date"].dt.year == prev_year) & (df_all["date"] <= prev_ref)]
        return (cur, prev,
                f"{cur_year}년 (~{ref.strftime('%m/%d')})",
                f"{prev_year}년 (~{prev_ref.strftime('%m/%d')})")
    elif mode == "전분기비 (QoQ)":
        cur_q  = pd.Period(last_date, "Q")
        prev_q = cur_q - 1
        cur  = df_all[df_all["date"].dt.to_period("Q") == cur_q]
        prev = df_all[df_all["date"].dt.to_period("Q") == prev_q]
        return cur, prev, str(cur_q), str(prev_q)
    elif mode == "전월비 (MoM)":
        cur_m  = pd.Period(last_date, "M")
        prev_m = cur_m - 1
        cur  = df_all[df_all["date"].dt.to_period("M") == cur_m]
        prev = df_all[df_all["date"].dt.to_period("M") == prev_m]
        return cur, prev, str(cur_m), str(prev_m)
    elif mode == "전주비 (WoW)":
        cur_w  = last_date.isocalendar()[1]
        prev_w = cur_w - 1
        cur_y  = last_date.year
        cur  = df_all[(df_all["date"].dt.isocalendar().week == cur_w) & (df_all["date"].dt.year == cur_y)]
        prev = df_all[(df_all["date"].dt.isocalendar().week == prev_w) & (df_all["date"].dt.year == cur_y)]
        return cur, prev, f"{cur_y}W{cur_w}", f"{cur_y}W{prev_w}"
    else:  # 전체
        mid = df_all["date"].median()
        cur  = df_all[df_all["date"] >= mid]
        prev = df_all[df_all["date"] <  mid]
        return cur, prev, "후반기", "전반기"

def cmp_selectbox(key, df_all):
    """비교 기준 selectbox + YoY일 때 기준일 date_input → (mode, yoy_ref) 반환"""
    mode = st.selectbox("비교 기준", COMPARE_OPTS, key=f"cmp_{key}")
    yoy_ref = None
    if mode == "전년비 (YoY)":
        cur_year = int(df_all["date"].dt.year.max())
        year_df  = df_all[df_all["date"].dt.year == cur_year]
        ref_date = st.date_input(
            f"기준일 ({cur_year}년 기준)",
            value=year_df["date"].max().date(),
            min_value=year_df["date"].min().date(),
            max_value=year_df["date"].max().date(),
            key=f"yoy_ref_{key}",
            help=f"선택한 날짜까지의 {cur_year}년 데이터와 전년도 같은 기간(1/1~같은 월일)을 비교합니다",
        )
        yoy_ref = pd.Timestamp(ref_date)
    return mode, yoy_ref

def compare_table(df_cur, df_prev, label_cur, label_prev, keys=None):
    """기간 비교 테이블 생성"""
    if keys is None:
        keys = ["perSend","revenue","rps","totalSend","customers","ctr",
                "purchaseRate","purchaseCust","purchaseCnt","avgOrderVal","unitPrice",
                "uniqueInflow","totalInflow","visitPerPerson","rpc"]
    rows = []
    for k in keys:
        try:
            v_cur  = df_cur[k].mean()  if len(df_cur)  else np.nan
            v_prev = df_prev[k].mean() if len(df_prev) else np.nan
            chg = pct(v_prev, v_cur) if not np.isnan(v_prev) and v_prev!=0 else np.nan
            rows.append({
                "지표": METRIC_LABELS.get(k, k),
                label_prev: fmt_val(k, v_prev),
                label_cur:  fmt_val(k, v_cur),
                "증감율": f"{chg:+.1f}%" if not np.isnan(chg) else "–",
                "_chg": chg, "_k": k,
            })
        except: pass
    df_out = pd.DataFrame(rows)
    return df_out

def styled_compare(df_tbl):
    """증감율에 색상 적용"""
    positive_good = {"revenue","purchaseCust","purchaseCnt","ctr","purchaseRate",
                     "rpc","avgOrderVal","uniqueInflow","totalInflow"}
    def color_row(row):
        chg = row.get("_chg", np.nan)
        k   = row.get("_k","")
        if np.isnan(chg): return [""] * len(row)
        good = (chg > 0) == (k in positive_good)
        color = "color: #16a34a; font-weight:600" if good else "color: #dc2626; font-weight:600"
        result = [""] * len(row)
        result[list(row.index).index("증감율")] = color
        return result

    display = df_tbl.drop(columns=["_chg","_k"], errors="ignore")
    try:
        return display.style.apply(color_row, axis=1)
    except:
        return display


# ══════════════════════════════════════════════════════
# Appendix
# ══════════════════════════════════════════════════════
def show_appendix(df_raw, label="근거 데이터"):
    with st.expander(f"근거 데이터 — {label}", expanded=False):
        st.markdown("<div class='appendix'>", unsafe_allow_html=True)
        st.dataframe(df_raw.reset_index(drop=True), use_container_width=True)
        csv = df_raw.to_csv(index=False).encode("utf-8-sig")
        st.download_button("CSV 다운로드", csv, file_name=f"{label}.csv", mime="text/csv")
        st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 통계 해설
# ══════════════════════════════════════════════════════
def stat_explainer():
    with st.expander("통계 용어 해설 (클릭하여 펼치기)", expanded=False):
        verdict_box("stat_r2",
            "R² (결정계수) — \"이 지표가 얼마나 일관된 패턴을 보이는가\"\n0~1 사이 숫자. 0.3 이상이면 뚜렷한 경향, 0.1~0.3은 약한 경향, 0.1 미만은 거의 패턴 없음.\n예: R²=0.286 → 시간 흐름에 따른 하락 패턴이 28.6% 설명됨 — 꽤 뚜렷한 경향",
            "vb")
        verdict_box("stat_p",
            "p값 (유의확률) — \"이 패턴이 우연일 확률\"\np<0.001  → 우연일 확률 0.1% 미만 = 거의 확실한 경향\np<0.01  → 우연일 확률 1% 미만 = 신뢰할 수 있는 경향\np<0.05  → 우연일 확률 5% 미만 = 통계적으로 유의함\np≥0.05 ns → 우연일 수도 있음 = 근거 부족",
            "vg")
        verdict_box("stat_corr",
            "상관계수 (Correlation) — \"두 지표가 얼마나 함께 움직이는가\"\n-1~+1 사이. 음수 = 한쪽이 오르면 다른 쪽 내려감, 양수 = 같이 움직임.\n-0.4 이하면 꽤 강한 음의 관계. 예: 발송↑ → 건당거래액↓",
            "va")
        verdict_box("stat_slope",
            "선형회귀 기울기 (slope) — \"하루에 얼마씩 변하는가\"\n예: slope=-0.0028%p/일 → CTR이 매일 0.0028%p씩 하락 = 1년이면 약 1%p 하락",
            "vb")


# ══════════════════════════════════════════════════════
# 메모 블록 — 상단/하단 어디든 삽입 가능
# ══════════════════════════════════════════════════════
DEFAULTS = {
    # 페이지 상단
    "top_01": "인당 발송 건수와 주요 효율 지표(CTR, 구매율, 건당 거래액)는 분석 기간 내내 반대 방향으로 움직였습니다. 발송이 늘어난 기간에 효율이 낮아지는 패턴이 일관되게 나타납니다.",
    "top_02": "발송량을 늘려도 총 거래액이 비례해서 증가하지 않습니다. 총 발송량 기준 하위 20% 구간(Q1)의 거래액이 상위 20% 구간(Q5)보다 오히려 높습니다.",
    "top_03": "인당 발송 건수가 늘어난 기간과 CTR·구매율이 하락한 기간이 일치합니다. 요일 효과를 통제한 후에도 이 패턴은 p<0.001 수준으로 유의합니다.",
    "top_04": "발송량과 거래액의 관계는 요일을 통제하면 일관된 방향성이 사라집니다. '발송을 늘리면 매출이 오른다'는 인과 주장은 이 데이터로 직접 입증되지 않습니다.",
    "top_05": "두 지표 간 상관계수가 강하다고 해서 인과관계가 있는 건 아닙니다. 패턴 확인 후 가설을 세우고, A/B 테스트로 검증하는 순서가 필요합니다.",
    "top_06": "요일별 발송 패턴과 효율 차이를 확인합니다. 특정 요일에 발송이 집중되어 있거나 효율이 낮다면, 발송 스케줄 조정이 효과적인 개선 방법이 될 수 있습니다.",
    "top_07": "거래액, 건당 거래액, CTR, 구매율에 가중치를 부여해 구간별 종합 점수를 산출합니다. 가중치를 조정하면 경영 우선순위에 맞는 최적 구간이 달라집니다.",
    "top_08": "발송 구간이 높아질수록 건당 거래액 증가폭이 줄어들다가 결국 마이너스로 전환됩니다. 이 시점이 추가 발송의 한계 수익이 0이 되는 구간입니다.",
    # 차트 하단
    "btm_trend":      "월별·분기별 추이에서 계절성이 보인다면 비교 기준 설정 시 같은 시즌끼리 비교하는 것이 더 정확합니다.",
    "btm_kpi":        "전체 평균 대비 증감은 필터 적용 기간 기준입니다. 기간을 바꾸면 수치가 달라집니다.",
    "btm_change":     "분기 처음과 끝을 비교하므로 중간의 등락은 반영되지 않습니다. 변화율이 크다면 월별 추이 차트로 구체적인 시점을 확인하세요.",
    "btm_cmp":        "기간 비교는 선택한 기준 기간의 일평균값 기준입니다. 기간 길이가 다른 경우 절대값보다 비율로 판단하는 것이 더 정확합니다.",
    "btm_bucket":     "표본이 적은 구간(10일 미만)은 집계에서 제외됩니다. 구간 우측 'n=N일'이 표본 수입니다. 30일 이상인 구간을 기준으로 해석하세요.",
    "btm_quintile":   "Q1~Q5는 총 발송 건수 기준 5등분입니다. 요일 구성이 구간마다 다를 수 있어 교란 효과가 완전히 제거되지는 않습니다.",
    "btm_dual":       "두 지표가 서로 반대 방향으로 움직이는 구간을 집중적으로 보세요. 방향이 바뀌는 시점이 정책 변화나 외부 요인과 겹치는지 확인할 필요가 있습니다.",
    "btm_reg":        "R²가 0.2 이상이면 시간에 따른 추세가 어느 정도 설명됩니다. p<0.001이면 이 패턴이 우연일 확률이 0.1% 미만입니다.",
    "btm_dow_comp":   "같은 요일 내 비교이므로 요일 효과는 통제됩니다. 방향이 요일마다 다르면 발송량 자체보다 다른 요인이 매출을 결정하고 있을 가능성이 높습니다.",
    "btm_corr_mat":   "대각선(자기 자신)은 항상 1입니다. 절대값이 0.5 이상인 셀을 중심으로 보세요.",
    "btm_scatter":    "산점도의 점이 추세선 주변에 얼마나 모여 있는지가 R²로 표현됩니다. 퍼져 있을수록 관계가 약합니다.",
    "btm_dow_heat":   "표준화 값이므로 절대 수치가 아닌 요일 간 상대적 차이를 나타냅니다.",
    "btm_score":      "가중 점수는 각 지표를 0~1로 정규화한 후 가중치를 곱한 합산입니다. 가중치 합이 1이 아니어도 됩니다.",
    "btm_marginal":   "한계 변화량이 음수로 전환되는 구간부터는 발송을 늘릴수록 해당 지표가 낮아지는 구조입니다.",
    "btm_dow_bar":    "요일별 편차가 크다면 발송 스케줄을 효율이 높은 요일에 집중하는 방안을 검토해볼 수 있습니다.",
}

def memo_block(key, location="bottom"):
    """편집 가능한 메모 블록 (글자색·크기·볼드 포맷 지원)"""
    store = st.session_state.insights
    mkey  = f"__memo_{key}__"
    ekey  = f"__medit_{key}__"
    hkey  = f"__mhide_{key}__"

    default_text = DEFAULTS.get(key, "")
    memo_base = {"bold": False, "color": "#475569", "size": 13}

    if mkey not in store:
        store[mkey] = {"text": default_text, **memo_base}
    item = _norm(store[mkey], memo_base)

    if ekey not in st.session_state:
        st.session_state[ekey] = False
    if hkey not in st.session_state:
        st.session_state[hkey] = False

    # 숨김 상태면 아무것도 안 보임
    if st.session_state[hkey]:
        if st.button("펼치기", key=f"mshow_{key}", help="메모 표시"):
            st.session_state[hkey] = False
            st.rerun()
        return

    if st.session_state[ekey]:
        c1, c2, c3 = st.columns([10, 1, 1])
        new_val = c1.text_area("", item["text"], key=f"mta_{key}",
                                height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"m_{key}", item)
        if c2.button("저장", key=f"msave_{key}", use_container_width=True):
            new_item = {"text": new_val, "bold": bold, "color": color, "size": size}
            store[mkey] = new_item
            st.session_state[ekey] = False
            all_data = load_insights(); all_data.update(store)
            save_insights(all_data)
            st.session_state.insights = store
            st.rerun()
        if c3.button("숨기기", key=f"mhide_{key}", use_container_width=True):
            st.session_state[hkey] = True
            st.session_state[ekey] = False
            st.rerun()
    else:
        item = _norm(store[mkey], memo_base)
        txt = item["text"]
        if txt:
            c1, c2 = st.columns([20, 1])
            c1.markdown(
                f'<p style="{_style(item)};margin:6px 0">{txt}</p>',
                unsafe_allow_html=True
            )
            if c2.button("편집", key=f"medit_{key}", use_container_width=True):
                st.session_state[ekey] = True
                st.rerun()
        else:
            c1, c2 = st.columns([20, 1])
            c1.markdown(
                '<p style="font-size:12px;color:#cbd5e1;margin:4px 0">메모를 입력하세요.</p>',
                unsafe_allow_html=True
            )
            if c2.button("편집", key=f"medit_{key}", use_container_width=True):
                st.session_state[ekey] = True
                st.rerun()

    store[mkey] = store.get(mkey, {"text": "", **memo_base})
    st.session_state.insights = store


# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("##  발송 분석")
    uploaded_list = st.file_uploader(
        "엑셀 업로드", type=["xlsx","xls"], label_visibility="collapsed",
        accept_multiple_files=True,
        help="기간이 다른 파일 여러 개를 올리면 날짜 기준으로 자동으로 이어붙입니다. 같은 날짜가 겹치면 나중 파일 값을 사용합니다.")

    if uploaded_list:
        files_bytes = tuple(f.read() for f in uploaded_list)
        G    = compute(files_bytes)
        meta = G["meta"]
        df   = G["df"]
        merged_note = f" · 파일 {len(uploaded_list)}개 병합" if len(uploaded_list) > 1 else ""
        st.success(f" {meta['days']}일 데이터{merged_note}")
        st.caption(f"{meta['start']} ~ {meta['end']}")

        # 날짜 필터
        st.markdown("---")
        st.markdown("** 기간 필터**")
        date_min = pd.to_datetime(meta["start"])
        date_max = pd.to_datetime(meta["end"])
        d1 = st.date_input("시작일", date_min, min_value=date_min, max_value=date_max)
        d2 = st.date_input("종료일", date_max, min_value=date_min, max_value=date_max)
        mask_date = (df["date"] >= pd.to_datetime(d1)) & (df["date"] <= pd.to_datetime(d2))
        df_f = df[mask_date].copy()

        # 요일 필터
        st.markdown("** 요일 필터**")
        DOW_KR = ["월","화","수","목","금","토","일"]
        sel_dow = st.multiselect("요일 선택", DOW_KR, default=DOW_KR)
        dow_map = {v:i for i,v in enumerate(["월","화","수","목","금","토","일"])}
        sel_dow_idx = [dow_map[d] for d in sel_dow]
        df_f = df_f[df_f["dow"].isin(sel_dow_idx)]

        st.markdown("---")
        st.caption(f"필터 적용 후: {len(df_f)}일")
    else:
        G = None; df_f = None

    # ── 동의 현황 파일 업로드 ──
    st.markdown("---")
    st.markdown("## 동의 현황")
    sms_up   = st.file_uploader("SMS 동의 파일", type=["xlsx","xls"], key="sms_up",  label_visibility="collapsed",
                                accept_multiple_files=True)
    st.caption("SMS 동의 현황 xlsx (여러 파일 가능)")
    push_up  = st.file_uploader("PUSH 동의 파일", type=["xlsx","xls"], key="push_up", label_visibility="collapsed",
                                accept_multiple_files=True)
    st.caption("PUSH 동의 현황 xlsx (여러 파일 가능)")

    def _load_consent_multi(files):
        """여러 동의 현황 파일을 날짜 기준으로 이어붙임 (중복 날짜는 나중 파일 우선)"""
        dfs = [parse_consent(f.read()) for f in files]
        return (pd.concat(dfs, ignore_index=True)
                  .drop_duplicates(subset="date", keep="last")
                  .sort_values("date").reset_index(drop=True))

    if sms_up:
        df_sms = _load_consent_multi(sms_up)
        merged = f" · {len(sms_up)}개 병합" if len(sms_up) > 1 else ""
        st.success(f"SMS {len(df_sms)}일{merged}")
        st.caption(f"{str(df_sms['date'].min().date())} ~ {str(df_sms['date'].max().date())}")
    else:
        df_sms = None

    if push_up:
        df_push = _load_consent_multi(push_up)
        merged = f" · {len(push_up)}개 병합" if len(push_up) > 1 else ""
        st.success(f"PUSH {len(df_push)}일{merged}")
        st.caption(f"{str(df_push['date'].min().date())} ~ {str(df_push['date'].max().date())}")
    else:
        df_push = None

    st.markdown("---")
    PAGE_LIST = [
        "01. 전체 요약",
        "02. 발송 빈도 효율",
        "03. 피로도 시계열",
        "04. 인과 검증",
        "05. 지표 상관 분석",
        "06. 요일별 패턴",
        "07. 발송 최적 구간",
        "08. 한계수익 분석",
        "09. PUSH 동의 현황",
        "10. SMS 동의 현황",
    ]
    page = st.radio("분석 주제", PAGE_LIST, label_visibility="collapsed")

    if G:
        st.markdown("---")
        if st.button("💾 전체 저장", key="sd_save_all", use_container_width=True,
                     help="편집한 모든 메모·텍스트를 insights_store.json에 저장합니다"):
            all_data = load_insights()
            all_data.update(st.session_state.insights)
            all_data["__texts__"] = st.session_state.get("editable_texts", {})
            save_insights(all_data)
            st.success("저장됐습니다 ✓")
        if st.button("초기화", use_container_width=True):
            st.session_state.insights = {}
            save_insights({})
            st.rerun()
        lines = ["# 발송 분석 인사이트"]
        for pk, items in st.session_state.insights.items():
            if pk.startswith("__"): continue
            if not isinstance(items, list): continue
            lines.append(f"\n## {pk}")
            for it in items:
                txt = it.get("text", str(it)) if isinstance(it, dict) else str(it)
                lines.append(f"- {txt}")
        st.download_button("메모 내보내기", "\n".join(lines),
                           file_name="insights.txt", mime="text/plain", use_container_width=True)

# ══════════════════════════════════════════════════════
# 동의 현황 페이지 — 발송 파일 없이도 독립 접근 가능
# ══════════════════════════════════════════════════════
if page == "09. PUSH 동의 현황":
    _t = editable_text("title_push", "PUSH 동의 현황 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    render_consent_page(df_push, "PUSH")
    st.stop()

if page == "10. SMS 동의 현황":
    _t = editable_text("title_sms", "SMS 동의 현황 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    render_consent_page(df_sms, "SMS")
    st.stop()

# ══════════════════════════════════════════════════════
# 업로드 전 안내 (발송 분석 페이지)
# ══════════════════════════════════════════════════════
if G is None:
    st.title("발송 분석 대시보드")
    st.markdown("왼쪽에서 **MTD 발송 상세 엑셀 파일**을 업로드하면 분석이 시작됩니다.")
    c1,c2,c3,c4 = st.columns(4)
    for cw,title,items in [
        (c1,"발송 지표",["인당 발송 건수","총 발송 건수","유니크 고객수"]),
        (c2,"효율 지표",["CTR","발송건당 거래액","구매전환율(CR)"]),
        (c3,"유입 지표",["유니크 유입","총 유입","인당 방문 횟수"]),
        (c4,"구매 지표",["구매 고객수","구매 건수","객단가"]),
    ]:
        with cw: st.info(f"**{title}**\n\n"+"\n\n".join(f"• {i}" for i in items))
    st.stop()

# ── 공통 데이터 ──
quarterly = G["quarterly"]
monthly   = G["monthly"]
buckets   = G["buckets"]
quintile  = G["quintile"]
dow_comp  = G["dow_comp"]
dow_corr  = G["dow_corr"]
reg       = G["reg"]

# 필터 적용 재집계
def reagg(df_filt, grp):
    if len(df_filt) == 0: return pd.DataFrame()
    return df_filt.groupby(grp, sort=True).agg(
        n=("revenue","count"),
        **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}
    ).reset_index()

monthly_f   = reagg(df_f, "month")
quarterly_f = reagg(df_f, "quarter")
df_f["bucket_f"] = pd.cut(df_f["perSend"], bins=[0,2,2.5,3,3.5,4,4.5,99],
                           labels=["~2.0","2.0~2.5","2.5~3.0","3.0~3.5","3.5~4.0","4.0~4.5","4.5+"])
buckets_f = df_f.groupby("bucket_f", observed=True).agg(
    n=("revenue","count"), **{k:pd.NamedAgg(k,"mean") for k in list(ALL_METRICS)+["purchaseRate","rpc"]}
).reset_index()
buckets_f = buckets_f[buckets_f["n"]>=10].reset_index(drop=True)

fq = quarterly_f.iloc[0] if len(quarterly_f) else quarterly.iloc[0]
lq = quarterly_f.iloc[-1] if len(quarterly_f) else quarterly.iloc[-1]

ALL_METRIC_KEYS = list(ALL_METRICS.keys()) + ["purchaseRate","rpc"]
metric_select_list = [METRIC_LABELS.get(k,k) for k in ALL_METRIC_KEYS]
def key_from_label(lbl): return next((k for k,v in METRIC_LABELS.items() if v==lbl), None)

def fmt_val(k, v):
    if k == "revenue": return f"{v/1e8:.3f}억"
    if k in ("totalSend","customers","uniqueInflow","totalInflow"):
        return f"{v/1e4:.1f}만" if v>10000 else f"{v:.0f}"
    if k in ("ctr","purchaseRate"): return f"{v*100:.2f}%"
    fmt = METRIC_FORMAT.get(k, ("{:.2f}",""))[0]
    try: return fmt.format(v)
    except: return str(round(v,2))

def metric_color(k):
    cmap = {"perSend":"amber","revenue":"blue","rps":"green","totalSend":"teal",
            "customers":"blue","ctr":"red","purchaseRate":"purple",
            "rpc":"green","purchaseCust":"purple","purchaseCnt":"teal",
            "avgOrderVal":"amber","unitPrice":"amber"}
    return cmap.get(k,"blue")

# ══════════════════════════════════════════════════════
# PAGE: 전체 요약
# ══════════════════════════════════════════════════════
if page == "01. 전체 요약":
    _t = editable_text("title_overview", "전체 요약", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption(f"데이터 기간: {meta['start']} ~ {meta['end']} · 전체 {meta['days']}일 · 필터 적용 {len(df_f)}일")

    memo_block("top_01")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    kpi_keys = ["perSend","revenue","ctr","purchaseRate","rps","purchaseCust","avgOrderVal","totalSend"]
    cols = st.columns(4)
    for i, k in enumerate(kpi_keys):
        v_now = df_f[k].mean() if len(df_f) else np.nan
        v_all = G["df"][k].mean()
        delta = pct(v_all, v_now)
        inv = k not in ["revenue","purchaseCust","avgOrderVal","ctr","purchaseRate","rps","rpc"]
        cols[i%4].metric(METRIC_LABELS[k], fmt_val(k,v_now),
                         f"{delta:+.1f}% vs 전체평균",
                         delta_color="inverse" if inv else "normal")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    memo_block("btm_kpi")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 주요 지표 트렌드 차트 (지표 선택)
    m1, m2 = st.columns([2,1])
    sel_metric_label = m1.selectbox("추이 지표 선택", metric_select_list, index=0)
    agg_unit = m2.selectbox("집계 단위", ["월별","분기별","연도별"])
    sel_k = key_from_label(sel_metric_label)

    agg_df = {"월별":monthly_f,"분기별":quarterly_f,"연도별":reagg(df_f,"year")}[agg_unit]
    x_col  = {"월별":"month","분기별":"quarter","연도별":"year"}[agg_unit]

    if sel_k and len(agg_df):
        yv = agg_df[sel_k].tolist()
        cn = metric_color(sel_k)
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=agg_df[x_col].astype(str).tolist(), y=yv,
            mode="lines+markers", line=dict(color=clr(cn),width=2),
            marker=dict(size=5,color=clr(cn)),
            fill="tozeroy", fillcolor=cbg(cn),
        ))
        layout = base_layout(260, title=f"{sel_metric_label} 추이")
        layout["xaxis"]["tickangle"] = -30
        fig.update_layout(**layout)
        st.plotly_chart(fig, use_container_width=True)
        show_appendix(agg_df[[x_col,"n",sel_k]].rename(columns={sel_k:sel_metric_label}), f"전체요약_{sel_metric_label}")
        memo_block("btm_trend")

    # 핵심 변화 요약
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간별 변화")
    cc1,cc2 = st.columns(2)
    for col_w, keys in [(cc1,["perSend","ctr","purchaseRate","rps"]),
                        (cc2,["revenue","purchaseCust","avgOrderVal","totalSend"])]:
        rows = []
        for k in keys:
            if len(quarterly_f) >= 2:
                chg = pct(fq[k], lq[k])
                rows.append({"지표":METRIC_LABELS[k], f"{fq['quarter']}":fmt_val(k,fq[k]),
                              f"{lq['quarter']}":fmt_val(k,lq[k]), "변화율":f"{chg:+.1f}%"})
        if rows: col_w.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    memo_block("btm_change")

    # ── 기간 비교 전체 테이블 ──
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 전체 지표")
    cmp_mode, _yoy_ref = cmp_selectbox("ov", G["df"])
    df_cur, df_prev, lbl_cur, lbl_prev = get_compare_periods(G["df"], cmp_mode, _yoy_ref)
    if len(df_cur) and len(df_prev):
        tbl_all = compare_table(df_cur, df_prev, lbl_cur, lbl_prev)
        st.dataframe(styled_compare(tbl_all), use_container_width=True, hide_index=True)
        show_appendix(tbl_all.drop(columns=["_chg","_k"],errors="ignore"), f"기간비교_{cmp_mode}")
    else:
        st.info("비교 가능한 이전 기간 데이터가 없습니다.")

# ══════════════════════════════════════════════════════
# PAGE: 발송 빈도 효율
# ══════════════════════════════════════════════════════
elif page == "02. 발송 빈도 효율":
    memo_block("top_02")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_freq", "발송 빈도 효율 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")

    q1v, q5v = quintile.iloc[0], quintile.iloc[4]
    verdict_box("freq_v1",
        f' "많이 보낼수록 매출이 오른다" — 기각\nQ1(평균 {q1v["totalSend"]/1e6:.2f}M건) 거래액 {q1v["revenue"]/1e8:.3f}억 vs Q5({q5v["totalSend"]/1e6:.2f}M건) {q5v["revenue"]/1e8:.3f}억. 2배 이상 더 보내도 매출은 오히려 낮습니다.',
        "vr")
    verdict_box("freq_v3", "", "vg")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("인당 발송 구간별 성과 (30일+ 구간)")

    sel_m2 = st.selectbox("지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"))
    sel_k2 = key_from_label(sel_m2)
    cn2    = metric_color(sel_k2) if sel_k2 else "blue"

    if sel_k2 and len(buckets_f):
        labels = buckets_f["bucket_f"].astype(str).tolist()
        yv2    = buckets_f[sel_k2].tolist()
        text2  = [fmt_val(sel_k2, v) for v in yv2]
        fig3 = go.Figure(go.Bar(
            x=labels, y=yv2,
            marker_color=cbg(cn2), marker_line_color=clr(cn2), marker_line_width=1.5,
            text=text2, textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout3 = base_layout(300, title=f"구간별 {sel_m2}")
        for i,row in buckets_f.iterrows():
            fig3.add_annotation(x=str(row["bucket_f"]),y=0,text=f"n={row['n']}일",
                showarrow=False,yanchor="top",yshift=-16,font=dict(size=10,color="#94a3b8"))
        fig3.update_layout(**layout3)
        st.plotly_chart(fig3, use_container_width=True)
        show_appendix(buckets_f[["bucket_f","n",sel_k2]].rename(columns={"bucket_f":"구간",sel_k2:sel_m2}), f"구간별_{sel_m2}")
        memo_block("btm_bucket")

    # 5분위 이중축
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("총 발송량 5분위 비교")
    ma, mb = st.columns(2)
    ym1_label = ma.selectbox("좌축 지표", metric_select_list, index=metric_select_list.index("거래액"), key="q5_l")
    ym2_label = mb.selectbox("우축 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="q5_r")
    ym1, ym2 = key_from_label(ym1_label), key_from_label(ym2_label)
    if ym1 and ym2:
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=list(quintile["label"]),y=list(quintile[ym1]),
            name=ym1_label,yaxis="y",marker_color=cbg("blue"),
            marker_line_color=clr("blue"),marker_line_width=1.5))
        fig4.add_trace(go.Scatter(x=list(quintile["label"]),y=list(quintile[ym2]),
            name=ym2_label,yaxis="y2",mode="lines+markers",
            line=dict(color=clr("amber"),width=2),marker=dict(size=6,color=clr("amber"))))
        layout4 = base_layout(280, title="5분위 비교")
        layout4["showlegend"] = True
        layout4["legend"]  = dict(orientation="h",y=1.05,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
        layout4["yaxis"]["title"] = ym1_label
        layout4["yaxis2"] = dict(overlaying="y",side="right",title=ym2_label,
                                  tickfont=dict(color="#64748b",size=11),gridcolor="rgba(0,0,0,0)")
        fig4.update_layout(**layout4)
        st.plotly_chart(fig4, use_container_width=True)
        show_appendix(quintile[["label",ym1,ym2]].rename(columns={"label":"5분위",ym1:ym1_label,ym2:ym2_label}), "5분위분석")
        memo_block("btm_quintile")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 발송 효율")
    cmp_mode_fr, _yoy_ref_fr = cmp_selectbox("fr", G["df"])
    df_cur_fr, df_prev_fr, lbl_cur_fr, lbl_prev_fr = get_compare_periods(G["df"], cmp_mode_fr, _yoy_ref_fr)
    if len(df_cur_fr) and len(df_prev_fr):
        tbl_fr = compare_table(df_cur_fr, df_prev_fr, lbl_cur_fr, lbl_prev_fr,
                               keys=["perSend","totalSend","customers","rps","revenue",
                                     "ctr","purchaseRate","purchaseCnt"])
        st.dataframe(styled_compare(tbl_fr), use_container_width=True, hide_index=True)
        show_appendix(tbl_fr.drop(columns=["_chg","_k"],errors="ignore"), f"발송효율비교_{cmp_mode_fr}")

# ══════════════════════════════════════════════════════
# PAGE: 피로도 시계열
# ══════════════════════════════════════════════════════
elif page == "03. 피로도 시계열":
    memo_block("top_03")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_ts", "피로도 시계열 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("요일 통제 후 잔차 회귀 기준")

    sc = pct(fq["perSend"], lq["perSend"])
    cc = pct(fq["ctr"], lq["ctr"])

    # 지표 선택 트렌드
    sel_ts = st.selectbox("분석 지표", metric_select_list, index=0, key="ts_sel")
    ts_k   = key_from_label(sel_ts)
    ts_agg = st.selectbox("집계 단위", ["월별","분기별"], key="ts_agg")
    ts_df  = monthly_f if ts_agg=="월별" else quarterly_f
    x_col  = "month" if ts_agg=="월별" else "quarter"

    if ts_k and len(ts_df):
        cn_ts = metric_color(ts_k)
        fig5 = go.Figure(go.Scatter(
            x=ts_df[x_col].astype(str).tolist(), y=ts_df[ts_k].tolist(),
            mode="lines+markers", line=dict(color=clr(cn_ts),width=2),
            marker=dict(size=5,color=clr(cn_ts)),
            fill="tozeroy", fillcolor=cbg(cn_ts),
        ))
        layout5 = base_layout(260, title=f"{sel_ts} 추이 ({ts_agg})")
        layout5["xaxis"]["tickangle"] = -30
        fig5.update_layout(**layout5)
        st.plotly_chart(fig5, use_container_width=True)

    # 이중축: 발송 vs 선택 지표
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("인당 발송 vs 효율 지표")
    dual_sel = st.selectbox("비교 지표 (우축)", metric_select_list,
                             index=metric_select_list.index("CTR"), key="dual_sel")
    dual_k = key_from_label(dual_sel)
    if dual_k and len(quarterly_f):
        fig6 = go.Figure()
        fig6.add_trace(go.Scatter(x=quarterly_f["quarter"].astype(str).tolist(),
            y=quarterly_f["perSend"].tolist(), name="인당발송(건)", yaxis="y",
            mode="lines+markers",line=dict(color=clr("blue"),width=2),
            marker=dict(size=5,color=clr("blue"))))
        fig6.add_trace(go.Scatter(x=quarterly_f["quarter"].astype(str).tolist(),
            y=quarterly_f[dual_k].tolist(), name=dual_sel, yaxis="y2",
            mode="lines+markers",line=dict(color=clr("red"),width=2,dash="dot"),
            marker=dict(size=5,color=clr("red"))))
        layout6 = base_layout(280, title="발송 vs 효율 (이중축)")
        layout6["showlegend"] = True
        layout6["legend"] = dict(orientation="h",y=1.05,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
        layout6["yaxis"]["title"]    = "인당발송(건)"
        layout6["yaxis"]["ticksuffix"] = "건"
        layout6["yaxis2"] = dict(overlaying="y",side="right",title=dual_sel,
                                  tickfont=dict(color="#64748b",size=11),gridcolor="rgba(0,0,0,0)")
        fig6.update_layout(**layout6)
        st.plotly_chart(fig6, use_container_width=True)
        show_appendix(quarterly_f[["quarter","perSend",dual_k]].rename(
            columns={"quarter":"분기","perSend":"인당발송",dual_k:dual_sel}), "피로도이중축")
        memo_block("btm_dual")

    # 회귀 통계
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("통계 유의성")
    st.caption(
        "slope = 하루 변화량 (양수=증가, 음수=감소) · "
        "R² = 패턴 강도 (0~1, 0.3↑이면 뚜렷) · "
        "p = 우연일 확률 (0.05 미만이면 신뢰 가능)"
    )
    stat_explainer()

    SLOPE_HELP = {
        "perSend":     "인당 발송 건수의 시간 흐름에 따른 변화. 양수면 발송이 늘어나는 추세, 음수면 줄어드는 추세.",
        "ctr":         "클릭률(CTR)의 시간 흐름에 따른 변화. 음수면 매일 조금씩 하락하는 중. 요일 효과를 제거한 잔차 기준.",
        "purchaseRate":"구매전환율의 시간 흐름에 따른 변화. 요일 효과를 제거한 잔차 기준.",
        "rps":         "발송건당 거래액의 시간 흐름에 따른 변화. 요일 효과를 제거한 잔차 기준.",
        "revenue":     "거래액의 시간 흐름에 따른 변화. 요일 효과를 제거한 잔차 기준.",
    }
    reg_cols = st.columns(len(reg))
    for i,(k,v) in enumerate(reg.items()):
        lbl = METRIC_LABELS.get(k,k)
        inv = k not in ["perSend","revenue"]
        slope_str = f"{v['slope']:+.5f}/일" if not np.isnan(v['slope']) else "–"
        r2_str    = r2_label(v['r2'])
        p_str     = sig_label(v['p'])
        reg_cols[i].metric(
            f"{lbl} 추세",
            slope_str,
            f"{r2_str}  |  {p_str}",
            delta_color="inverse" if inv else "normal",
            help=SLOPE_HELP.get(k, "")
        )
    memo_block("btm_reg")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 효율 지표")
    cmp_mode_ts, _yoy_ref_ts = cmp_selectbox("ts", G["df"])
    df_cur_ts, df_prev_ts, lbl_cur_ts, lbl_prev_ts = get_compare_periods(G["df"], cmp_mode_ts, _yoy_ref_ts)
    if len(df_cur_ts) and len(df_prev_ts):
        tbl_ts = compare_table(df_cur_ts, df_prev_ts, lbl_cur_ts, lbl_prev_ts,
                               keys=["perSend","ctr","purchaseRate","rps","revenue","rpc"])
        st.dataframe(styled_compare(tbl_ts), use_container_width=True, hide_index=True)
        show_appendix(tbl_ts.drop(columns=["_chg","_k"],errors="ignore"), f"피로도비교_{cmp_mode_ts}")

# ══════════════════════════════════════════════════════
# PAGE: 인과 검증
# ══════════════════════════════════════════════════════
elif page == "04. 인과 검증":
    memo_block("top_04")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_causal", "인과 검증", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    verdict_box("causal_v2", "", "vg")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일별 통제 비교")
    dc_rows = []
    for r in dow_comp:
        dc_rows.append({"요일":r["dow"]+"요일","기준":f"≤{r['median']:.1f}건",
            "적은날 거래액":f"{r['lowRev']/1e8:.3f}억","많은날 거래액":f"{r['highRev']/1e8:.3f}억",
            "차이":f"{'▲' if r['diff']>0 else '▼'}{abs(r['diff'])/1e6:.1f}M ({'적은날↑' if r['diff']>0 else '많은날↑'})",
            "건당거래(적↔많)":f"{r['lowRps']:.0f}원↔{r['highRps']:.0f}원"})
    dc_df = pd.DataFrame(dc_rows)
    st.dataframe(dc_df, use_container_width=True, hide_index=True)
    show_appendix(dc_df, "요일통제비교")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일별 상관계수")
    dc_corr = pd.DataFrame(dow_corr)
    # 요일 순서 고정
    DOW_ORDER = ["월","화","수","목","금","토","일"]
    dc_corr["dow"] = pd.Categorical(dc_corr["dow"], categories=DOW_ORDER, ordered=True)
    dc_corr = dc_corr.sort_values("dow").reset_index(drop=True)
    fig7 = go.Figure()
    # 단일 Bar — 양수 파란색, 음수 빨간색을 marker_color 리스트로
    bar_colors = [cbg("blue") if v >= 0 else cbg("red") for v in dc_corr["corrRevenue"]]
    bar_line_colors = [clr("blue") if v >= 0 else clr("red") for v in dc_corr["corrRevenue"]]
    fig7.add_trace(go.Bar(
        x=dc_corr["dow"].tolist(),
        y=dc_corr["corrRevenue"].round(3).tolist(),
        name="vs 거래액",
        marker_color=bar_colors,
        marker_line_color=bar_line_colors,
        marker_line_width=1.5,
    ))
    fig7.add_trace(go.Scatter(
        x=dc_corr["dow"].tolist(),
        y=dc_corr["corrRps"].round(3).tolist(),
        name="vs 건당거래액", mode="lines+markers",
        line=dict(color=clr("red"),width=2),
        marker=dict(size=6,color=clr("red")),
    ))
    layout7 = base_layout(260, title="상관계수 — 인당 발송 vs 지표")
    layout7["showlegend"] = True
    layout7["legend"] = dict(orientation="h",y=1.1,bgcolor="rgba(0,0,0,0)",font=dict(color="#64748b"))
    layout7["yaxis"]["range"] = [-1, 0.5]
    layout7["yaxis"]["title"] = "상관계수"
    layout7["shapes"] = [dict(type="line",x0=-0.5,x1=len(dc_corr)-0.5,y0=0,y1=0,
                               line=dict(color="#cbd5e1",width=1,dash="dot"))]
    fig7.update_layout(**layout7)
    st.plotly_chart(fig7, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    verdict_box("causal_ab2", "", "vg")
    st.subheader("기간 비교 — 인과 관련 지표")
    cmp_mode_ca, _yoy_ref_ca = cmp_selectbox("ca", G["df"])
    df_cur_ca, df_prev_ca, lbl_cur_ca, lbl_prev_ca = get_compare_periods(G["df"], cmp_mode_ca, _yoy_ref_ca)
    if len(df_cur_ca) and len(df_prev_ca):
        tbl_ca = compare_table(df_cur_ca, df_prev_ca, lbl_cur_ca, lbl_prev_ca,
                               keys=["perSend","revenue","rps","ctr","purchaseRate","rpc"])
        st.dataframe(styled_compare(tbl_ca), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════
# PAGE: 지표 상관 분석
# ══════════════════════════════════════════════════════
elif page == "05. 지표 상관 분석":
    memo_block("top_05")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_corr", "지표 상관 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("두 지표 간의 관계를 산점도와 상관계수로 확인합니다.")

    ca, cb = st.columns(2)
    xm_label = ca.selectbox("X축 지표", metric_select_list, index=metric_select_list.index("인당 발송 건수"), key="corr_x")
    ym_label = cb.selectbox("Y축 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="corr_y")
    xk = key_from_label(xm_label); yk = key_from_label(ym_label)

    if xk and yk and len(df_f) > 5:
        xv = df_f[xk].dropna(); yv_s = df_f[yk].dropna()
        common = df_f[[xk,yk]].dropna()
        corr_val = np.corrcoef(common[xk], common[yk])[0,1]
        reg_r = linreg(common[xk].values, common[yk].values)

        c1,c2,c3 = st.columns(3)
        c1.metric(
            "상관계수",
            f"{corr_val:.3f}",
            corr_label(corr_val),
            help="두 지표가 얼마나 함께 움직이는지 나타냅니다. -1~+1 사이. 음수=한쪽 오를 때 다른 쪽 내려감, 양수=같이 움직임, 0=무관계. 절대값이 0.5 이상이면 강한 관계입니다."
        )
        c2.metric(
            "R²",
            f"{reg_r['r2']:.3f}",
            sig_label(reg_r["p"]),
            help="추세선이 데이터를 얼마나 잘 설명하는지 나타냅니다. 0~1 사이. 0.3 이상=뚜렷한 경향, 0.1~0.3=약한 경향, 0.1 미만=경향 거의 없음."
        )
        c3.metric(
            "표본 수",
            f"{len(common)}일",
            help="분석에 사용된 데이터 일수입니다. 30일 이상이어야 통계 결과를 신뢰할 수 있습니다."
        )
        st.caption(
            f"해석: 두 지표는 {corr_label(corr_val)} · {r2_label(reg_r['r2'])} · {sig_label(reg_r['p'])}"
        )
        stat_explainer()

        fig8 = go.Figure(go.Scatter(
            x=common[xk].tolist(), y=common[yk].tolist(),
            mode="markers", marker=dict(color=clr("blue"),size=5,opacity=0.6),
            text=df_f.loc[common.index,"date"].dt.strftime("%Y-%m-%d").tolist(),
        ))
        # 추세선 (기울기·절편 모두 반영 — 데이터를 실제로 통과하는 회귀선)
        xfit = np.linspace(common[xk].min(), common[xk].max(), 100)
        yfit = reg_r["slope"]*xfit + reg_r["intercept"]
        fig8.add_trace(go.Scatter(x=xfit.tolist(), y=yfit.tolist(),
            mode="lines", line=dict(color=clr("red"),width=1.5,dash="dot"), name="추세선"))
        layout8 = base_layout(300, title=f"{xm_label} vs {ym_label} (r={corr_val:.3f})")
        layout8["xaxis"]["title"] = xm_label
        layout8["yaxis"]["title"] = ym_label
        fig8.update_layout(**layout8)
        st.plotly_chart(fig8, use_container_width=True)
        show_appendix(common.rename(columns={xk:xm_label,yk:ym_label}), f"산점도_{xm_label}_{ym_label}")
        memo_block("btm_scatter")

    # 상관 매트릭스
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("지표 간 상관관계")
    mat_keys = ["perSend","revenue","rps","ctr","purchaseRate","purchaseCnt","avgOrderVal"]
    mat_labels = [METRIC_LABELS[k] for k in mat_keys]
    if len(df_f) > 10:
        corr_mat = df_f[mat_keys].corr().round(3)
        corr_mat.columns = mat_labels; corr_mat.index = mat_labels
        fig9 = go.Figure(go.Heatmap(
            z=corr_mat.values.tolist(), x=mat_labels, y=mat_labels,
            colorscale=[[0,"rgba(245,101,101,0.8)"],[0.5,"rgba(248,249,252,1)"],[1,"rgba(79,143,255,0.8)"]],
            zmid=0, text=corr_mat.round(2).values.tolist(), texttemplate="%{text}",
            showscale=True,
        ))
        layout9 = base_layout(380, title="지표 간 상관 매트릭스")
        layout9["xaxis"]["tickangle"] = -30
        fig9.update_layout(**layout9)
        st.plotly_chart(fig9, use_container_width=True)
        show_appendix(corr_mat.reset_index().rename(columns={"index":"지표"}), "상관매트릭스")
        memo_block("btm_corr_mat")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 전체 지표")
    cmp_mode_co, _yoy_ref_co = cmp_selectbox("co", G["df"])
    df_cur_co, df_prev_co, lbl_cur_co, lbl_prev_co = get_compare_periods(G["df"], cmp_mode_co, _yoy_ref_co)
    if len(df_cur_co) and len(df_prev_co):
        tbl_co = compare_table(df_cur_co, df_prev_co, lbl_cur_co, lbl_prev_co)
        st.dataframe(styled_compare(tbl_co), use_container_width=True, hide_index=True)
        show_appendix(tbl_co.drop(columns=["_chg","_k"],errors="ignore"), f"전지표비교_{cmp_mode_co}")

# ══════════════════════════════════════════════════════
# PAGE: 요일별 패턴
# ══════════════════════════════════════════════════════
elif page == "06. 요일별 패턴":
    memo_block("top_06")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_dow", "요일별 패턴 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")

    sel_dow_m = st.selectbox("분석 지표", metric_select_list, index=metric_select_list.index("거래액"), key="dow_m")
    dow_k = key_from_label(sel_dow_m)
    DOW_KR = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}

    if dow_k and len(df_f):
        dow_agg = df_f.groupby("dow")[dow_k].mean().reset_index()
        dow_agg["요일"] = dow_agg["dow"].map(DOW_KR)

        fig10 = go.Figure(go.Bar(
            x=dow_agg["요일"].tolist(), y=dow_agg[dow_k].tolist(),
            marker_color=cbg("blue"), marker_line_color=clr("blue"), marker_line_width=1.5,
            text=[fmt_val(dow_k,v) for v in dow_agg[dow_k]], textposition="outside",
            textfont=dict(color="#64748b",size=11),
        ))
        layout10 = base_layout(280, title=f"요일별 {sel_dow_m} 평균")
        fig10.update_layout(**layout10)
        st.plotly_chart(fig10, use_container_width=True)
        memo_block("btm_dow_bar")

    # 요일 × 지표 히트맵
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("요일 × 지표 히트맵")
    heat_keys = ["perSend","revenue","rps","ctr","purchaseRate","avgOrderVal"]
    heat_labels = [METRIC_LABELS[k] for k in heat_keys]
    if len(df_f):
        heat_df = df_f.groupby("dow")[heat_keys].mean()
        heat_norm = (heat_df - heat_df.mean()) / (heat_df.std() + 1e-9)
        heat_norm.index = [DOW_KR[i] for i in heat_norm.index]
        heat_norm.columns = heat_labels
        fig11 = go.Figure(go.Heatmap(
            z=heat_norm.values.tolist(), x=heat_labels, y=list(heat_norm.index),
            colorscale=[[0,"rgba(245,101,101,0.8)"],[0.5,"rgba(248,249,252,1)"],[1,"rgba(79,143,255,0.8)"]],
            zmid=0, text=heat_df.round(2).values.tolist(), texttemplate="%{text}",
        ))
        layout11 = base_layout(280, title="요일 × 지표 (표준화)")
        fig11.update_layout(**layout11)
        st.plotly_chart(fig11, use_container_width=True)
        show_appendix(heat_df.reset_index().rename(columns={"dow":"요일"}), "요일별히트맵")
        memo_block("btm_dow_heat")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 요일별 성과")
    cmp_mode_dw, _yoy_ref_dw = cmp_selectbox("dw", G["df"])
    df_cur_dw, df_prev_dw, lbl_cur_dw, lbl_prev_dw = get_compare_periods(G["df"], cmp_mode_dw, _yoy_ref_dw)
    DOW_KR2 = {0:"월",1:"화",2:"수",3:"목",4:"금",5:"토",6:"일"}
    if len(df_cur_dw) and len(df_prev_dw) and dow_k:
        dw_rows = []
        for d in range(7):
            vc = df_cur_dw[df_cur_dw["dow"]==d][dow_k].mean()
            vp = df_prev_dw[df_prev_dw["dow"]==d][dow_k].mean()
            chg = pct(vp, vc) if not np.isnan(vp) and vp!=0 else np.nan
            dw_rows.append({"요일":DOW_KR2[d]+"요일", lbl_prev_dw:fmt_val(dow_k,vp),
                             lbl_cur_dw:fmt_val(dow_k,vc), "증감율":f"{chg:+.1f}%" if not np.isnan(chg) else "–"})
        st.dataframe(pd.DataFrame(dw_rows), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════
# PAGE: 발송 최적 구간
# ══════════════════════════════════════════════════════
elif page == "07. 발송 최적 구간":
    memo_block("top_07")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_opt", "발송 최적 구간 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("거래액·효율·구매율을 종합한 최적 발송 구간을 찾습니다.")

    if len(buckets_f):
        # 종합 점수: 각 지표를 0-1 정규화 후 가중 합산
        score_keys = ["revenue","rps","ctr","purchaseRate"]
        score_labels = [METRIC_LABELS[k] for k in score_keys]
        weights = {}
        st.subheader("가중치 설정")
        wcols = st.columns(len(score_keys))
        for i,(k,lbl) in enumerate(zip(score_keys,score_labels)):
            weights[k] = wcols[i].slider(lbl, 0.0, 1.0, 0.25, 0.05, key=f"w_{k}")

        # 정규화 점수
        scored = buckets_f.copy()
        for k in score_keys:
            mn, mx = scored[k].min(), scored[k].max()
            scored[f"norm_{k}"] = (scored[k]-mn)/(mx-mn+1e-9)
        scored["score"] = sum(scored[f"norm_{k}"]*weights[k] for k in score_keys)
        scored["score"] = scored["score"] / (sum(weights.values())+1e-9)

        best_idx = scored["score"].idxmax()
        best = scored.iloc[best_idx]

        fig12 = go.Figure()
        fig12.add_trace(go.Bar(
            x=scored["bucket_f"].astype(str).tolist(),
            y=scored["score"].round(3).tolist(),
            marker_color=[clr("green") if i==best_idx else cbg("blue") for i in range(len(scored))],
            marker_line_color=clr("blue"), marker_line_width=1.5,
            text=[f"{v:.3f}" for v in scored["score"]],
            textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout12 = base_layout(280, title="구간별 종합 점수 (가중 합산)")
        fig12.update_layout(**layout12)
        st.plotly_chart(fig12, use_container_width=True)
        show_appendix(scored[["bucket_f","n","revenue","rps","ctr","purchaseRate","score"]].rename(
            columns={"bucket_f":"구간","revenue":"거래액","rps":"건당거래","ctr":"CTR",
                     "purchaseRate":"구매율","score":"종합점수"}), "최적구간분석")
        memo_block("btm_score")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 구간 효과")
    cmp_mode_op, _yoy_ref_op = cmp_selectbox("op", G["df"])
    df_cur_op, df_prev_op, lbl_cur_op, lbl_prev_op = get_compare_periods(G["df"], cmp_mode_op, _yoy_ref_op)
    if len(df_cur_op) and len(df_prev_op):
        tbl_op = compare_table(df_cur_op, df_prev_op, lbl_cur_op, lbl_prev_op,
                               keys=["perSend","revenue","rps","ctr","purchaseRate","avgOrderVal"])
        st.dataframe(styled_compare(tbl_op), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════
# PAGE: 한계수익 분석
# ══════════════════════════════════════════════════════
elif page == "08. 한계수익 분석":
    memo_block("top_08")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    _t = editable_text("title_lm", "한계수익 분석", "h2", "font-size:1.8rem;font-weight:700;color:#1e293b")
    st.caption("발송 1건 추가 시 거래액·클릭이 얼마나 증가/감소하는지 분석합니다.")

    sel_lm = st.selectbox("분석 지표", metric_select_list, index=metric_select_list.index("발송건당 거래액"), key="lm_sel")
    lm_k   = key_from_label(sel_lm)

    if lm_k and len(buckets_f) >= 2:
        bk = buckets_f.copy()
        bk["marginal"] = bk[lm_k].diff()
        bk["bucket_str"] = bk["bucket_f"].astype(str)

        c1_lm, c2_lm = st.columns(2)
        # 절대값
        fig13 = go.Figure(go.Bar(
            x=bk["bucket_str"].tolist(), y=bk[lm_k].tolist(),
            marker_color=cbg("green"), marker_line_color=clr("green"), marker_line_width=1.5,
            text=[fmt_val(lm_k,v) for v in bk[lm_k]], textposition="outside",
            textfont=dict(color="#64748b",size=11),
        ))
        layout13 = base_layout(260, title=f"구간별 {sel_lm}")
        fig13.update_layout(**layout13)
        c1_lm.plotly_chart(fig13, use_container_width=True)

        # 한계 변화
        fig14 = go.Figure(go.Bar(
            x=bk["bucket_str"].tolist(), y=bk["marginal"].tolist(),
            marker_color=[cbg("green") if v>0 else cbg("red") for v in bk["marginal"].fillna(0)],
            marker_line_color=[clr("green") if v>0 else clr("red") for v in bk["marginal"].fillna(0)],
            marker_line_width=1.5,
            text=[f"{v:+.1f}" if not np.isnan(v) else "" for v in bk["marginal"]],
            textposition="outside", textfont=dict(color="#64748b",size=11),
        ))
        layout14 = base_layout(260, title=f"한계 변화량 (구간 간 차이)")
        layout14["shapes"] = [dict(type="line",x0=-0.5,x1=len(bk)-0.5,y0=0,y1=0,
                                    line=dict(color="#2a2d3a",width=1,dash="dot"))]
        fig14.update_layout(**layout14)
        c2_lm.plotly_chart(fig14, use_container_width=True)
        show_appendix(bk[["bucket_str","n",lm_k,"marginal"]].rename(
            columns={"bucket_str":"구간",lm_k:sel_lm,"marginal":"한계변화"}), "한계수익분석")
        memo_block("btm_marginal")


    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader("기간 비교 — 한계 효율")
    cmp_mode_lm, _yoy_ref_lm = cmp_selectbox("lm", G["df"])
    df_cur_lm, df_prev_lm, lbl_cur_lm, lbl_prev_lm = get_compare_periods(G["df"], cmp_mode_lm, _yoy_ref_lm)
    if len(df_cur_lm) and len(df_prev_lm):
        tbl_lm = compare_table(df_cur_lm, df_prev_lm, lbl_cur_lm, lbl_prev_lm,
                               keys=["perSend","rps","ctr","purchaseRate","revenue","rpc"])
        st.dataframe(styled_compare(tbl_lm), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════
# 동의 현황 공통 렌더러
# ══════════════════════════════════════════════════════
