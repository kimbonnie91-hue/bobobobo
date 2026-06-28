"""첫구매 채널별 분석 대시보드 — Plotly 인터랙티브 버전"""

import io, json, os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from scipy import stats

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="첫구매 채널 분석", page_icon="🛒",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "fp_insights.json"

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
# 색상 팔레트
# ══════════════════════════════════════════════════════
PALETTE = {
    "blue":   ("rgba(79,143,255,1)",  "rgba(79,143,255,0.15)"),
    "red":    ("rgba(245,101,101,1)", "rgba(245,101,101,0.15)"),
    "green":  ("rgba(72,187,120,1)",  "rgba(72,187,120,0.15)"),
    "amber":  ("rgba(237,137,54,1)",  "rgba(237,137,54,0.15)"),
    "purple": ("rgba(159,122,234,1)", "rgba(159,122,234,0.15)"),
    "teal":   ("rgba(56,178,172,1)",  "rgba(56,178,172,0.15)"),
    "orange": ("rgba(249,115,22,1)",  "rgba(249,115,22,0.15)"),
    "slate":  ("rgba(100,116,139,1)", "rgba(100,116,139,0.15)"),
    "pink":   ("rgba(236,72,153,1)",  "rgba(236,72,153,0.15)"),
}
def clr(n): return PALETTE.get(n, PALETTE["blue"])[0]
def cbg(n): return PALETTE.get(n, PALETTE["blue"])[1]

CHANNEL_PAL = {
    "직접":      "blue",
    "광고":      "amber",
    "EP":        "green",
    "PUSH":      "purple",
    "제휴":      "red",
    "브랜드광고": "teal",
    "미디어커머스": "orange",
    "*TOTAL":   "slate",
}
CHANNELS = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]

def ch_clr(ch): return clr(CHANNEL_PAL.get(ch, "blue"))
def ch_cbg(ch): return cbg(CHANNEL_PAL.get(ch, "blue"))

def base_layout(h=280, ysuffix="", title=""):
    return dict(
        paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
        font=dict(color="#475569", size=11), margin=dict(l=10, r=10, t=36, b=10),
        height=h, showlegend=False,
        title=dict(text=title, font=dict(color="#c0c8d8", size=13)),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=11), ticksuffix=ysuffix),
    )

# ══════════════════════════════════════════════════════
# 인사이트 저장소
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(data):
    with open(INSIGHT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "insights" not in st.session_state:
    st.session_state.insights = load_insights()

# ══════════════════════════════════════════════════════
# 텍스트 포맷 헬퍼
# ══════════════════════════════════════════════════════
FMT_DEFAULTS = {"bold": False, "color": "#334155", "size": 13}
SIZE_OPTS    = [11, 12, 13, 14, 15, 16, 18, 20, 24]

def _norm(val, base=None):
    d = (base or FMT_DEFAULTS).copy()
    if isinstance(val, str):
        d["text"] = val; return d
    if isinstance(val, dict):
        merged = d.copy()
        merged.update({k: v for k, v in val.items() if k != "editing"})
        merged.setdefault("text", "")
        return merged
    d["text"] = ""; return d

def _style(fmt):
    w = "700" if fmt.get("bold") else "400"
    c = fmt.get("color", "#334155")
    s = fmt.get("size", 13)
    return f"font-weight:{w};color:{c};font-size:{s}px;line-height:1.7"

def _fmt_ui(key, fmt):
    fc1, fc2, fc3 = st.columns([1, 2, 2])
    bold = fc1.checkbox("굵게", value=bool(fmt.get("bold", False)), key=f"_ck_b_{key}")
    raw_c = fmt.get("color", "#334155")
    safe_c = raw_c if (isinstance(raw_c, str) and raw_c.startswith("#") and len(raw_c) == 7) else "#334155"
    color = fc2.color_picker("색상", value=safe_c, key=f"_ck_c_{key}")
    sz  = fmt.get("size", 13)
    idx = SIZE_OPTS.index(sz) if sz in SIZE_OPTS else SIZE_OPTS.index(13)
    size = fc3.selectbox("크기(px)", SIZE_OPTS, index=idx, key=f"_ck_s_{key}")
    return bold, color, size

# ══════════════════════════════════════════════════════
# 인사이트 에디터
# ══════════════════════════════════════════════════════
def insight_editor(page_key, default_lines):
    store = st.session_state.insights
    if page_key not in store:
        store[page_key] = []
    items    = store[page_key]
    hide_key = f"__hide_memo_{page_key}__"
    if hide_key not in st.session_state:
        st.session_state[hide_key] = False

    h1, h2, h3 = st.columns([8, 1, 1])
    h1.markdown("**메모**", unsafe_allow_html=True)
    if h2.button("숨기기" if not st.session_state[hide_key] else "펼치기",
                 key=f"hide_{page_key}", use_container_width=True):
        st.session_state[hide_key] = not st.session_state[hide_key]; st.rerun()
    if h3.button("+ 추가", key=f"add_{page_key}", use_container_width=True):
        items.append({"text": "", "editing": True, "bold": False, "color": "#334155", "size": 13})
        store[page_key] = items; st.session_state.insights = store
        save_insights(store); st.rerun()

    if not st.session_state[hide_key]:
        if not items:
            st.markdown("<p style='color:#94a3b8;font-size:12px;margin:4px 0 0 2px'>+ 추가 버튼으로 메모를 입력하세요.</p>",
                        unsafe_allow_html=True)
        to_delete = []
        for i, raw_item in enumerate(items):
            item = _norm(raw_item) if not isinstance(raw_item, dict) else raw_item
            item.setdefault("bold", False); item.setdefault("color", "#334155"); item.setdefault("size", 13)
            edit_active = item.get("editing", False)
            c1, c2, c3 = st.columns([10, 1, 1])
            if edit_active:
                new_text = c1.text_area("", item.get("text", ""), key=f"memo_ta_{page_key}_{i}",
                                        height=72, label_visibility="collapsed", placeholder="내용을 입력하세요.")
                bold, color, size = _fmt_ui(f"memo_{page_key}_{i}", item)
                if c2.button("저장", key=f"memo_save_{page_key}_{i}", use_container_width=True):
                    items[i] = {"text": new_text, "editing": False, "bold": bold, "color": color, "size": size}
                    store[page_key] = items; st.session_state.insights = store
                    save_insights(store); st.rerun()
                if c3.button("삭제", key=f"memo_del_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)
            else:
                txt = item.get("text", "")
                c1.markdown(f"<p style='{_style(item)};margin:2px 0'>{txt}</p>", unsafe_allow_html=True)
                if c2.button("편집", key=f"memo_edit_{page_key}_{i}", use_container_width=True):
                    items[i]["editing"] = True; store[page_key] = items
                    st.session_state.insights = store; save_insights(store); st.rerun()
                if c3.button("삭제", key=f"memo_del2_{page_key}_{i}", use_container_width=True):
                    to_delete.append(i)
        for idx in sorted(to_delete, reverse=True):
            items.pop(idx)
        if to_delete:
            store[page_key] = items; st.session_state.insights = store
            save_insights(store); st.rerun()

    store[page_key] = items; st.session_state.insights = store; save_insights(store)

# ══════════════════════════════════════════════════════
# 편집 가능 텍스트
# ══════════════════════════════════════════════════════
if "editable_texts" not in st.session_state:
    st.session_state.editable_texts = load_insights().get("__fp_texts__", {})

TAG_FMT = {
    "h2": {"bold": True,  "color": "#1e293b", "size": 24},
    "h3": {"bold": True,  "color": "#1e293b", "size": 20},
    "h4": {"bold": True,  "color": "#1e293b", "size": 16},
    "p":  {"bold": False, "color": "#334155", "size": 13},
}

def editable_text(key, default, tag="p", style=""):
    base  = TAG_FMT.get(tag, FMT_DEFAULTS).copy()
    texts = st.session_state.editable_texts
    if key not in texts:
        texts[key] = {"text": default, **base}
    item     = _norm(texts[key], base)
    edit_key = f"__fpedit_{key}__"
    if edit_key not in st.session_state:
        st.session_state[edit_key] = False
    if st.session_state[edit_key]:
        col1, col2 = st.columns([10, 1])
        new_val = col1.text_area("", item["text"], key=f"fpta_{key}", height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(key, item)
        if col2.button("확인", key=f"fpsave_{key}"):
            new_item = {"text": new_val, "bold": bold, "color": color, "size": size}
            texts[key] = new_item; st.session_state[edit_key] = False
            all_data = load_insights(); all_data["__fp_texts__"] = texts
            save_insights(all_data); st.session_state.editable_texts = texts; st.rerun()
    else:
        item = _norm(texts[key], base)
        st.markdown(
            f'<{tag} style="{_style(item)};cursor:pointer;{style}" title="클릭하여 편집">{item["text"]}</{tag}>',
            unsafe_allow_html=True)
        if st.button("편집", key=f"fpedit_{key}", help="편집"):
            st.session_state[edit_key] = True; st.rerun()
    return item["text"]

def verdict_box(key, default_text, default_color="vg"):
    store = st.session_state.insights
    vkey  = f"__fpverdict_{key}__"
    ekey  = f"__fpediting_v_{key}__"
    if vkey not in store:
        store[vkey] = {"text": default_text, "bold": False, "color": "#334155", "size": 13}
    item = _norm(store[vkey])
    if ekey not in st.session_state:
        st.session_state[ekey] = False
    if st.session_state[ekey]:
        c1, c2 = st.columns([11, 1])
        new_text = c1.text_area("", item["text"], key=f"fpvta_{key}", height=100, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"fpv_{key}", item)
        if c2.button("확인", key=f"fpvsave_{key}"):
            new_item = {"text": new_text, "bold": bold, "color": color, "size": size}
            store[vkey] = new_item; st.session_state[ekey] = False
            all_data = load_insights(); all_data.update(store)
            save_insights(all_data); st.session_state.insights = store; st.rerun()
    else:
        item = _norm(store[vkey])
        txt  = item["text"].replace("\n", "<br>")
        col_a, col_b = st.columns([20, 1])
        col_a.markdown(f'<div class="{default_color}"><span style="{_style(item)}">{txt}</span></div>',
                       unsafe_allow_html=True)
        if col_b.button("편집", key=f"fpvedit_{key}"):
            st.session_state[ekey] = True; st.rerun()

# ══════════════════════════════════════════════════════
# 메모 블록
# ══════════════════════════════════════════════════════
FP_DEFAULTS = {
    "top_overview":  "첫구매 채널별 거래액·고객수·전환율(CR) 전반을 한눈에 파악합니다. 채널간 비중 차이와 전체 추이를 함께 보면 집중해야 할 채널이 드러납니다.",
    "btm_overview":  "채널별 평균값만으로는 계절성이나 추세를 파악하기 어렵습니다. 거래액 분석 또는 채널 효율 페이지에서 추이를 함께 확인하세요.",
    "top_tx":        "채널별 일평균거래액 추이를 비교합니다. 7일 이동평균으로 단기 노이즈를 제거하면 중기 방향성이 더 명확하게 나타납니다.",
    "btm_tx":        "월별 채널 비중 히트맵에서 특정 채널의 비중이 급변하는 시점을 확인하세요. 해당 시점의 프로모션 이력과 교차 확인이 필요합니다.",
    "top_cust":      "일평균고객수와 DAU 추이를 채널별로 비교합니다. 고객수가 늘어도 거래액이 정체된다면 객단가 하락을 의심해볼 수 있습니다.",
    "btm_cust":      "고객수 vs 거래액 산점도에서 우상향 클러스터가 나타나는 채널이 핵심 채널입니다. 이탈한 점들은 이상치 또는 특수 프로모션 기간일 수 있습니다.",
    "top_eff":       "CR(전환율)과 유입율은 채널 효율의 핵심 지표입니다. 유입율이 높아도 CR이 낮으면 랜딩 품질이나 상품 매칭을 점검해야 합니다.",
    "btm_eff":       "유입율 → CR 회귀분석에서 R²가 높은 채널은 유입 품질이 전환에 직결됩니다. R²가 낮으면 다른 요인(상품, 가격, UI)이 전환을 결정하고 있을 가능성이 높습니다.",
    "top_share":     "거래액·고객 비중의 시계열 변화를 누적 영역 차트로 확인합니다. 특정 채널의 비중이 지속적으로 줄어들고 있다면 경쟁력 약화 또는 예산 감소를 의미할 수 있습니다.",
    "btm_share":     "요일별 채널 CR 히트맵에서 색이 진한 셀이 '효율이 높은 요일+채널' 조합입니다. 발송/광고 일정을 이 조합에 맞게 조정하는 것을 고려해보세요.",
    "top_monthly":   "월별로 각 지표를 채널 단위로 요약합니다. MoM 증감을 기준으로 목표 달성 여부와 이탈 채널을 빠르게 파악할 수 있습니다.",
    "btm_monthly":   "전체 데이터 다운로드 후 피벗테이블을 활용해 임의 기간의 채널 비교 분석을 진행할 수 있습니다.",
}

def memo_block(key, location="bottom"):
    store = st.session_state.insights
    mkey  = f"__fpmemo_{key}__"
    ekey  = f"__fpmedit_{key}__"
    hkey  = f"__fpmhide_{key}__"
    default_text = FP_DEFAULTS.get(key, "")
    memo_base = {"bold": False, "color": "#475569", "size": 13}
    if mkey not in store:
        store[mkey] = {"text": default_text, **memo_base}
    item = _norm(store[mkey], memo_base)
    if ekey not in st.session_state: st.session_state[ekey] = False
    if hkey not in st.session_state: st.session_state[hkey] = False

    if st.session_state[hkey]:
        if st.button("펼치기", key=f"fpmshow_{key}", help="메모 표시"):
            st.session_state[hkey] = False; st.rerun()
        return

    if st.session_state[ekey]:
        c1, c2, c3 = st.columns([10, 1, 1])
        new_val = c1.text_area("", item["text"], key=f"fpmta_{key}",
                               height=80, label_visibility="collapsed")
        bold, color, size = _fmt_ui(f"fpm_{key}", item)
        if c2.button("저장", key=f"fpmsave_{key}", use_container_width=True):
            new_item = {"text": new_val, "bold": bold, "color": color, "size": size}
            store[mkey] = new_item; st.session_state[ekey] = False
            all_data = load_insights(); all_data.update(store)
            save_insights(all_data); st.session_state.insights = store; st.rerun()
        if c3.button("숨기기", key=f"fpmhide_{key}", use_container_width=True):
            st.session_state[hkey] = True; st.session_state[ekey] = False; st.rerun()
    else:
        item = _norm(store[mkey], memo_base)
        txt = item["text"]
        c1, c2 = st.columns([20, 1])
        if txt:
            c1.markdown(f'<p style="{_style(item)};margin:6px 0">{txt}</p>', unsafe_allow_html=True)
        else:
            c1.markdown('<p style="font-size:12px;color:#cbd5e1;margin:4px 0">메모를 입력하세요.</p>',
                        unsafe_allow_html=True)
        if c2.button("편집", key=f"fpmedit_{key}", use_container_width=True):
            st.session_state[ekey] = True; st.rerun()

    store[mkey] = store.get(mkey, {"text": "", **memo_base})
    st.session_state.insights = store

# ══════════════════════════════════════════════════════
# 통계 헬퍼
# ══════════════════════════════════════════════════════
def sig_label(p):
    if np.isnan(p): return "–"
    if p < 0.001: return "p<0.001 (우연일 확률 0.1% 미만 — 매우 확실)"
    if p < 0.01:  return "p<0.01 (우연일 확률 1% 미만 — 신뢰할 수 있음)"
    if p < 0.05:  return "p<0.05 (우연일 확률 5% 미만 — 유의함)"
    return "유의하지 않음 (우연일 수 있음)"

def r2_label(r2):
    if np.isnan(r2): return "–"
    if r2 >= 0.3:  return f"R²={r2:.3f} → 뚜렷한 경향 (변동의 {r2*100:.0f}% 설명)"
    if r2 >= 0.1:  return f"R²={r2:.3f} → 약한 경향 (변동의 {r2*100:.0f}% 설명)"
    return f"R²={r2:.3f} → 경향 거의 없음 ({r2*100:.0f}% 미만 설명)"

def corr_label(r):
    if np.isnan(r): return "–"
    ar   = abs(r)
    sign = "반대 방향" if r < 0 else "같은 방향"
    if ar >= 0.7:   strength = "매우 강하게"
    elif ar >= 0.5: strength = "강하게"
    elif ar >= 0.3: strength = "보통으로"
    elif ar >= 0.1: strength = "약하게"
    else:           strength = "거의 관계 없이"
    return f"{strength} {sign}으로 움직임"

def stat_explainer():
    with st.expander("통계 용어 해설 (클릭하여 펼치기)", expanded=False):
        verdict_box("fp_stat_r2",
            "R² (결정계수) — \"이 지표가 얼마나 일관된 패턴을 보이는가\"\n0~1 사이 숫자. 0.3 이상이면 뚜렷한 경향, 0.1~0.3은 약한 경향, 0.1 미만은 거의 패턴 없음.", "vb")
        verdict_box("fp_stat_p",
            "p값 (유의확률) — \"이 패턴이 우연일 확률\"\np<0.001 → 우연일 확률 0.1% 미만 = 거의 확실한 경향\np<0.05 → 우연일 확률 5% 미만 = 통계적으로 유의함\np≥0.05 ns → 우연일 수도 있음 = 근거 부족", "vg")
        verdict_box("fp_stat_corr",
            "상관계수 — \"두 지표가 얼마나 함께 움직이는가\"\n-1~+1. 음수=한쪽이 오르면 다른쪽 내려감, 양수=같이 움직임.\n±0.5 이상이면 꽤 강한 관계.", "va")

# ══════════════════════════════════════════════════════
# 기간 비교
# ══════════════════════════════════════════════════════
COMPARE_OPTS = ["전년비 (YoY)", "전분기비 (QoQ)", "전월비 (MoM)", "전주비 (WoW)", "전체 기간 처음↔끝"]

def cmp_selectbox(key):
    """비교 기준 selectbox + YoY일 때 기준일 date_input 반환 → (mode, yoy_ref)"""
    mode = st.selectbox("비교 기준", COMPARE_OPTS, key=f"fp_cmp_{key}")
    yoy_ref = None
    if mode == "전년비 (YoY)":
        avail_years = sorted(df_full["date"].dt.year.unique(), reverse=True)
        cur_year = avail_years[0]
        max_date = df_full[df_full["date"].dt.year == cur_year]["date"].max().date()
        ref_date = st.date_input(
            f"기준일 ({cur_year}년 기준)",
            value=max_date,
            min_value=df_full[df_full["date"].dt.year == cur_year]["date"].min().date(),
            max_value=max_date,
            key=f"fp_yoy_ref_{key}",
            help=f"선택한 날짜까지의 {cur_year}년 데이터와 전년도 같은 기간을 비교합니다",
        )
        yoy_ref = pd.Timestamp(ref_date)
    return mode, yoy_ref

def get_compare_periods(df_all, mode, yoy_ref=None):
    df_all = df_all.sort_values("date")
    last_date = df_all["date"].max()
    if mode == "전년비 (YoY)":
        # 동일 기간 비교: 현재 연도의 최대 날짜 기준 (또는 사용자 지정) 같은 월/일까지
        ref = yoy_ref if yoy_ref else last_date
        cur_year  = ref.year
        prev_year = cur_year - 1
        # 현재 연도: 1/1 ~ ref 날짜
        cur  = df_all[(df_all["date"].dt.year == cur_year) &
                      (df_all["date"] <= ref)]
        # 전년도: 1/1 ~ 전년도 같은 월/일
        try:
            prev_ref = ref.replace(year=prev_year)
        except ValueError:  # 2/29 같은 경우
            prev_ref = ref.replace(year=prev_year, day=28)
        prev = df_all[(df_all["date"].dt.year == prev_year) &
                      (df_all["date"] <= prev_ref)]
        lc = f"{cur_year}년 (~{ref.strftime('%m/%d')})"
        lp = f"{prev_year}년 (~{prev_ref.strftime('%m/%d')})"
        return cur, prev, lc, lp
    elif mode == "전분기비 (QoQ)":
        cur_q  = pd.Period(last_date, "Q"); prev_q = cur_q - 1
        cur  = df_all[df_all["date"].dt.to_period("Q") == cur_q]
        prev = df_all[df_all["date"].dt.to_period("Q") == prev_q]
        return cur, prev, str(cur_q), str(prev_q)
    elif mode == "전월비 (MoM)":
        cur_m  = pd.Period(last_date, "M"); prev_m = cur_m - 1
        cur  = df_all[df_all["date"].dt.to_period("M") == cur_m]
        prev = df_all[df_all["date"].dt.to_period("M") == prev_m]
        return cur, prev, str(cur_m), str(prev_m)
    elif mode == "전주비 (WoW)":
        cur_w = last_date.isocalendar()[1]; prev_w = cur_w - 1; cur_y = last_date.year
        cur  = df_all[(df_all["date"].dt.isocalendar().week == cur_w) & (df_all["date"].dt.year == cur_y)]
        prev = df_all[(df_all["date"].dt.isocalendar().week == prev_w) & (df_all["date"].dt.year == cur_y)]
        return cur, prev, f"{cur_y}W{cur_w}", f"{cur_y}W{prev_w}"
    else:
        mid  = df_all["date"].median()
        cur  = df_all[df_all["date"] >= mid]; prev = df_all[df_all["date"] < mid]
        return cur, prev, "후반기", "전반기"

def avail_segs(metric, df=None):
    """해당 지표에서 실제 non-NaN 값을 가진 세그먼트 목록 반환"""
    if df is None: df = df_full
    return (df[df["metric"] == metric]
            .groupby("segment")["value"]
            .apply(lambda s: s.notna().any())
            .pipe(lambda x: x[x].index.tolist()))

def compare_table(df_cur, df_prev, label_cur, label_prev, metric, channels):
    """롱포맷 기준 기간 비교 테이블 — 데이터 없는 세그먼트는 건너뜀"""
    rows = []
    for ch in channels:
        v_cur  = df_cur[(df_cur["metric"] == metric) & (df_cur["segment"] == ch)]["value"].mean()
        v_prev = df_prev[(df_prev["metric"] == metric) & (df_prev["segment"] == ch)]["value"].mean()
        if np.isnan(v_cur) and np.isnan(v_prev): continue  # 둘 다 없으면 스킵
        chg    = (v_cur - v_prev) / v_prev * 100 if (not np.isnan(v_prev) and v_prev != 0) else np.nan
        rows.append({
            "채널":    ch,
            label_prev: v_prev,
            label_cur:  v_cur,
            "증감율":  f"{chg:+.1f}%" if not np.isnan(chg) else "–",
            "_chg":    chg,
        })
    return pd.DataFrame(rows)

def styled_compare(df_tbl):
    def color_row(row):
        chg = row.get("_chg", np.nan)
        if np.isnan(chg): return [""] * len(row)
        color = "color: #16a34a; font-weight:600" if chg > 0 else "color: #dc2626; font-weight:600"
        result = [""] * len(row)
        if "증감율" in row.index:
            result[list(row.index).index("증감율")] = color
        return result
    display = df_tbl.drop(columns=["_chg"], errors="ignore")
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
# 포맷 헬퍼
# ══════════════════════════════════════════════════════
METRIC_KO = {
    "일평균거래액":  "일평균거래액",
    "거래액비중":    "거래액 비중",
    "일평균고객수":  "일평균고객수",
    "고객비중":      "고객 비중",
    "일평균객단가":  "일평균객단가",
    "유효회원수":    "유효회원수",
    "DAU":           "DAU",
    "유입율":        "유입율",
    "CR":            "전환율(CR)",
}

def fmt_num(v, suffix=""):
    if pd.isnull(v): return "–"
    av = abs(v)
    if av >= 1e8: return f"{v/1e8:.1f}억{suffix}"
    if av >= 1e4: return f"{v/1e4:.0f}만{suffix}"
    return f"{v:,.0f}{suffix}"

def fmt_pct(v):
    if pd.isnull(v): return "–"
    return f"{v*100:.2f}%"

# ══════════════════════════════════════════════════════
# 파싱
# ══════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def parse_xlsx(file_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="openpyxl")

    # row 0에서 연도 마커 위치 수집 (예: 2025→col6, 2026→col371)
    year_markers = {}
    for ci, v in enumerate(df_raw.iloc[0, :]):
        try:
            iv = int(float(v))
            if 2020 <= iv <= 2030:
                year_markers[ci] = iv
        except Exception:
            pass

    # 컬럼 인덱스별 연도 매핑 (각 마커 시작 컬럼부터 다음 마커 전까지 같은 연도)
    sorted_markers = sorted(year_markers.items())
    default_year = sorted_markers[0][1] if sorted_markers else 2026
    col_to_year = {}
    for idx, (col, yr) in enumerate(sorted_markers):
        end_col = sorted_markers[idx + 1][0] if idx + 1 < len(sorted_markers) else df_raw.shape[1]
        for c in range(col, end_col):
            col_to_year[c] = yr

    raw_dates = df_raw.iloc[2, 6:].tolist()
    dates = []
    for offset, ds in enumerate(raw_dates):
        ci = offset + 6
        yr = col_to_year.get(ci, default_year)
        if pd.isnull(ds):
            dates.append(pd.NaT)
        else:
            try:
                dates.append(pd.Timestamp(f"{yr}/{str(ds).strip()}"))
            except Exception:
                dates.append(pd.NaT)

    n = len(dates)
    records = []
    current_metric = None
    for ri in range(3, df_raw.shape[0]):
        m   = df_raw.iloc[ri, 0]
        seg = df_raw.iloc[ri, 1]
        if not pd.isnull(m) and str(m).strip() not in ("-", ""):
            current_metric = str(m).strip()
        if pd.isnull(seg) or str(seg).strip() in ("-", "") or current_metric is None:
            continue
        seg  = str(seg).strip()
        vals = pd.to_numeric(df_raw.iloc[ri, 6:6+n].values, errors="coerce")
        for dt, val in zip(dates, vals):
            if pd.isnull(dt): continue
            records.append({"date": dt, "metric": current_metric, "segment": seg, "value": val})

    df = pd.DataFrame(records)
    df["date"]  = pd.to_datetime(df["date"])
    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["quarter"] = df["date"].dt.to_period("Q").astype(str)
    df["year"]  = df["date"].dt.year
    df["dow"]   = df["date"].dt.dayofweek
    # 비율 지표 클리핑 (CR, 유입율, 거래액비중, 고객비중: 0~1 범위 강제)
    pct_metrics = {"CR", "유입율", "거래액비중", "고객비중"}
    mask = df["metric"].isin(pct_metrics)
    df.loc[mask, "value"] = df.loc[mask, "value"].clip(lower=0.0, upper=1.0)
    return df

# ══════════════════════════════════════════════════════
# 사이드바
# ══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🛒 첫구매 분석")
    uploaded = st.file_uploader("Excel 파일 업로드", type=["xlsx", "xls"], key="fp_up",
                                 accept_multiple_files=True,
                                 help="첫구매 채널별 지표 엑셀 파일을 업로드하세요. 여러 파일을 올리면 날짜 기준으로 이어붙입니다 (중복 날짜는 나중 파일 우선).")
    st.markdown("---")

    PAGE_LIST = [
        "01. 개요",
        "02. 거래액 분석",
        "03. 고객수 분석",
        "04. 채널 효율",
        "05. 채널 비중 추이",
        "06. 월별 요약",
    ]
    if "fp_page" not in st.session_state:
        st.session_state.fp_page = PAGE_LIST[0]
    page = st.radio("페이지", PAGE_LIST, key="fp_page")

# ── 데이터 없으면 정지
if not uploaded:
    st.info("👈 사이드바에서 Excel 파일을 업로드해주세요.")
    st.stop()

with st.spinner("데이터 파싱 중…"):
    # 여러 파일 이어붙이기 — 같은 (날짜, 지표, 채널)이 겹치면 나중 파일 값 사용
    _dfs = [parse_xlsx(f.read()) for f in uploaded]
    df_all = (pd.concat(_dfs, ignore_index=True)
                .drop_duplicates(subset=["date", "metric", "segment"], keep="last")
                .sort_values("date").reset_index(drop=True))
    if len(uploaded) > 1:
        st.caption(f"📎 파일 {len(uploaded)}개 병합 완료 — {df_all['date'].min().date()} ~ {df_all['date'].max().date()}")

if df_all.empty:
    st.error("데이터를 읽지 못했습니다. 엑셀 파일 형식을 확인해주세요.")
    st.stop()

# ── 사이드바 필터 (데이터 로드 후)
with st.sidebar:
    st.markdown("---")
    st.markdown("**기간 필터**")
    dmin = df_all["date"].min().date()
    dmax = df_all["date"].max().date()
    d_from, d_to = st.date_input(
        "날짜 범위", value=(dmin, dmax), min_value=dmin, max_value=dmax, key="fp_daterange"
    )

    st.markdown("**요일 필터**")
    DOW_KO = {0:"월", 1:"화", 2:"수", 3:"목", 4:"금", 5:"토", 6:"일"}
    dow_sel = st.multiselect("요일 선택", options=list(DOW_KO.values()),
                              default=list(DOW_KO.values()), key="fp_dow")
    dow_nums = [k for k, v in DOW_KO.items() if v in dow_sel]

    st.markdown("**채널 필터**")
    ch_sel = st.multiselect("채널 선택", options=CHANNELS, default=CHANNELS, key="fp_ch")

    st.markdown("---")
    if st.button("💾 전체 저장", key="fp_save_all", use_container_width=True,
                 help="편집한 모든 메모·텍스트를 fp_insights.json에 저장합니다"):
        all_data = load_insights()
        all_data.update(st.session_state.insights)
        all_data["__fp_texts__"] = st.session_state.get("editable_texts", {})
        save_insights(all_data)
        st.success("저장됐습니다 ✓")

    if st.button("메모 JSON 내보내기", key="fp_export_memo", use_container_width=True):
        j = json.dumps(st.session_state.insights, ensure_ascii=False, indent=2)
        st.download_button("다운로드", j.encode("utf-8"), "fp_insights.json", "application/json")

# ── 필터 적용
_date_mask = (df_all["date"].dt.date >= d_from) & (df_all["date"].dt.date <= d_to) & df_all["dow"].isin(dow_nums)
df_full = df_all[_date_mask].copy()                                    # 채널 필터 미적용 (TOTAL 포함)
df_f    = df_full[df_full["segment"].isin(ch_sel + ["*TOTAL"])].copy()  # 채널 필터 적용

# ── 공통 헬퍼
def get_pivot(metric, segs=None, df=None):
    if df is None: df = df_full
    mask = df["metric"] == metric
    if segs: mask &= df["segment"].isin(segs)
    return (df[mask]
            .pivot_table(index="date", columns="segment", values="value", aggfunc="first")
            .sort_index())

def get_series(metric, seg, df=None):
    if df is None: df = df_full
    return (df[(df["metric"] == metric) & (df["segment"] == seg)]
            .sort_values("date").set_index("date")["value"])

def monthly_avg(metric, segs, df=None):
    if df is None: df = df_full
    return (df[(df["metric"] == metric) & (df["segment"].isin(segs))]
            .groupby(["month", "segment"])["value"].mean()
            .unstack("segment").reindex(columns=segs))

def ch_list():
    return [c for c in ch_sel if c in CHANNELS] if ch_sel else CHANNELS

# ══════════════════════════════════════════════════════════════════════
# 01. 개요
# ══════════════════════════════════════════════════════════════════════
if page == "01. 개요":
    editable_text("fp_title_overview", "첫구매 채널 분석 — 개요", tag="h2")
    st.caption(f"데이터 기간: {df_full['date'].min().strftime('%Y-%m-%d')} ~ {df_full['date'].max().strftime('%Y-%m-%d')}  |  {df_full['date'].nunique()}일")

    memo_block("top_overview")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # KPI 카드
    last_d = df_full["date"].max()
    def week_avg(metric, seg):
        cur  = df_full[(df_full["metric"] == metric) & (df_full["segment"] == seg) &
                       (df_full["date"] >= last_d - pd.Timedelta(days=6))]["value"].mean()
        prev = df_full[(df_full["metric"] == metric) & (df_full["segment"] == seg) &
                       (df_full["date"] >= last_d - pd.Timedelta(days=13)) &
                       (df_full["date"] <  last_d - pd.Timedelta(days=6))]["value"].mean()
        delta = (cur - prev) / prev if (prev and not np.isnan(prev)) else None
        return cur, delta

    kpi_specs = [
        ("일평균거래액", "*TOTAL", "일평균거래액",   lambda v: fmt_num(v, "원")),
        ("일평균고객수", "*TOTAL", "일평균고객수",   lambda v: f"{v:,.0f}명"),
        ("일평균객단가", "*TOTAL", "일평균객단가",   lambda v: fmt_num(v, "원")),
        ("CR",          "*TOTAL", "전환율(CR)",      fmt_pct),
        ("유입율",       "*TOTAL", "유입율",         fmt_pct),
    ]
    cols = st.columns(5)
    for col, (metric, seg, label, fmtr) in zip(cols, kpi_specs):
        cur, delta = week_avg(metric, seg)
        dval = f"{delta*100:+.1f}% (전주비)" if delta is not None else None
        col.metric(label, fmtr(cur) if not pd.isnull(cur) else "–", delta=dval)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    cL, cR = st.columns(2)

    # 채널별 거래액 비중 (최근 30일)
    with cL:
        st.subheader("채널별 거래액 비중 (최근 30일)")
        share_30 = (
            df_full[(df_full["metric"] == "거래액비중") &
                    (df_full["segment"].isin(ch_list())) &
                    (df_full["date"] >= last_d - pd.Timedelta(days=29))]
            .groupby("segment")["value"].mean()
            .reindex(ch_list()).fillna(0)
        )
        fig = go.Figure(go.Bar(
            x=share_30.values * 100, y=share_30.index.tolist(),
            orientation="h",
            marker_color=[ch_cbg(c) for c in share_30.index],
            marker_line_color=[ch_clr(c) for c in share_30.index],
            marker_line_width=1.2,
            text=[f"{v*100:.1f}%" for v in share_30.values],
            textposition="outside",
        ))
        ly = base_layout(260, title="채널별 거래액 비중 (%)")
        ly["xaxis"]["ticksuffix"] = "%"
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    # 일평균거래액 전체 추이
    with cR:
        st.subheader("일평균거래액 전체 추이 (7일 이동평균)")
        total_tx = get_series("일평균거래액", "*TOTAL").rolling(7).mean().dropna()
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=total_tx.index.strftime("%m/%d").tolist(), y=(total_tx / 1e8).tolist(),
            mode="lines", line=dict(color=clr("blue"), width=2),
            fill="tozeroy", fillcolor=cbg("blue"), name="일평균거래액",
        ))
        ly = base_layout(260, ysuffix="억", title="일평균거래액 (7일 MA)")
        ly["xaxis"]["nticks"] = 15
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 채널별 요약표
    st.subheader("채널별 주요 지표 요약 (필터 기간 평균)")
    summary_rows = []
    for ch in ch_list():
        row = {"채널": ch}
        for m, fmtr in [("일평균거래액", lambda v: fmt_num(v, "원")),
                        ("일평균고객수", lambda v: f"{v:,.0f}명"),
                        ("일평균객단가", lambda v: fmt_num(v, "원")),
                        ("CR",   fmt_pct), ("유입율", fmt_pct)]:
            s = df_full[(df_full["metric"] == m) & (df_full["segment"] == ch)]["value"]
            row[METRIC_KO.get(m, m)] = fmtr(s.mean()) if len(s) else "–"
        summary_rows.append(row)
    st.dataframe(pd.DataFrame(summary_rows).set_index("채널"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 기간 비교
    st.subheader("기간 비교")
    cmp_mode, _yoy_ref = cmp_selectbox("ov")
    cmp_cur, cmp_prev, lc, lp = get_compare_periods(df_full, cmp_mode, _yoy_ref)
    for m in ["일평균거래액", "일평균고객수", "CR"]:
        tbl = compare_table(cmp_cur, cmp_prev, lc, lp, m, ch_list())
        if not tbl.empty:
            st.caption(METRIC_KO.get(m, m))
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    memo_block("btm_overview")
    insight_editor("fp_overview", ["첫구매 채널 전반 인사이트를 입력하세요."])

# ══════════════════════════════════════════════════════════════════════
# 02. 거래액 분석
# ══════════════════════════════════════════════════════════════════════
elif page == "02. 거래액 분석":
    editable_text("fp_title_tx", "거래액 분석", tag="h2")
    memo_block("top_tx")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 채널별 추이 (7일 MA)
    st.subheader("채널별 일평균거래액 추이 (7일 이동평균)")
    pivot_tx = get_pivot("일평균거래액", ch_list()).rolling(7).mean()
    fig = go.Figure()
    for ch in ch_list():
        if ch in pivot_tx.columns:
            s = (pivot_tx[ch].dropna() / 1e8)
            fig.add_trace(go.Scatter(
                x=s.index.strftime("%m/%d").tolist(), y=s.tolist(),
                mode="lines", name=ch,
                line=dict(color=ch_clr(ch), width=1.8),
            ))
    ly = base_layout(300, ysuffix="억", title="채널별 일평균거래액 (7일 MA)")
    ly["showlegend"] = True
    ly["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    ly["xaxis"]["nticks"] = 15
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    cL, cR = st.columns(2)

    # 월별 채널별 거래액 (grouped bar)
    with cL:
        st.subheader("월별 채널별 일평균거래액")
        m_avg = monthly_avg("일평균거래액", ch_list())
        if not m_avg.empty:
            fig = go.Figure()
            for ch in ch_list():
                if ch in m_avg.columns:
                    fig.add_trace(go.Bar(
                        x=m_avg.index.tolist(), y=(m_avg[ch] / 1e8).tolist(),
                        name=ch, marker_color=ch_cbg(ch),
                        marker_line_color=ch_clr(ch), marker_line_width=1,
                    ))
            ly = base_layout(280, ysuffix="억", title="월별 일평균거래액")
            ly["barmode"] = "group"; ly["showlegend"] = True
            ly["legend"] = dict(orientation="h", y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b", size=10))
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)

    # 거래액 비중 히트맵
    with cR:
        st.subheader("월별 채널 거래액 비중 히트맵")
        hm_tx = monthly_avg("거래액비중", ch_list()).fillna(0)
        if not hm_tx.empty:
            z_vals = (hm_tx.T * 100).values
            text_vals = [[f"{v:.1f}%" for v in row] for row in z_vals]
            fig = go.Figure(go.Heatmap(
                z=z_vals, x=hm_tx.index.tolist(), y=hm_tx.columns.tolist(),
                text=text_vals, texttemplate="%{text}",
                colorscale="YlOrRd", showscale=True,
                colorbar=dict(title="비중%", thickness=12),
            ))
            ly = base_layout(280, title="거래액 비중 (%)")
            ly["xaxis"]["tickangle"] = -30
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # Box (일별 분포 → violin/box via Plotly)
    st.subheader("채널별 일 거래액 분포")
    fig = go.Figure()
    for ch in ch_list():
        vals = df_full[(df_full["metric"] == "일평균거래액") & (df_full["segment"] == ch)]["value"].dropna()
        if len(vals):
            fig.add_trace(go.Box(
                y=(vals / 1e8).tolist(), name=ch,
                marker_color=ch_cbg(ch), line_color=ch_clr(ch),
                boxpoints="outliers",
            ))
    ly = base_layout(280, ysuffix="억", title="채널별 일평균거래액 분포")
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 기간 비교
    st.subheader("기간 비교")
    cmp_mode, _yoy_ref = cmp_selectbox("tx")
    cmp_cur, cmp_prev, lc, lp = get_compare_periods(df_full, cmp_mode, _yoy_ref)
    tbl = compare_table(cmp_cur, cmp_prev, lc, lp, "일평균거래액", ch_list())
    if not tbl.empty:
        st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    show_appendix(
        df_full[df_full["metric"] == "일평균거래액"][["date","segment","value"]].rename(
            columns={"date":"날짜","segment":"채널","value":"일평균거래액(원)"}),
        "거래액_분석")

    memo_block("btm_tx")
    insight_editor("fp_tx", ["거래액 분석 인사이트를 입력하세요."])

# ══════════════════════════════════════════════════════════════════════
# 03. 고객수 분석
# ══════════════════════════════════════════════════════════════════════
elif page == "03. 고객수 분석":
    editable_text("fp_title_cust", "고객수 분석", tag="h2")
    memo_block("top_cust")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    cL, cR = st.columns(2)

    def multi_line_chart(metric, title, ysuffix="", divisor=1.0, h=260):
        pivot = get_pivot(metric, ch_list()).rolling(7).mean()
        fig = go.Figure()
        for ch in ch_list():
            if ch in pivot.columns:
                s = pivot[ch].dropna() / divisor
                fig.add_trace(go.Scatter(
                    x=s.index.strftime("%m/%d").tolist(), y=s.tolist(),
                    mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
                ))
        ly = base_layout(h, ysuffix=ysuffix, title=title)
        ly["showlegend"] = True
        ly["legend"] = dict(orientation="h", y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b", size=10))
        ly["xaxis"]["nticks"] = 12
        fig.update_layout(**ly)
        return fig

    with cL:
        st.subheader("일평균고객수 추이 (7일 이동평균)")
        st.plotly_chart(multi_line_chart("일평균고객수", "일평균고객수 (7일 MA)", "명"), use_container_width=True)

    with cR:
        st.subheader("DAU 추이 (7일 이동평균)")
        st.plotly_chart(multi_line_chart("DAU", "DAU (7일 MA)", "명"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    st.subheader("유효회원수 추이 (7일 이동평균)")
    st.plotly_chart(multi_line_chart("유효회원수", "유효회원수 (7일 MA)", "만명", divisor=1e4, h=280),
                    use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    cL2, cR2 = st.columns(2)

    # 고객 비중 히트맵
    with cL2:
        st.subheader("월별 채널 고객 비중 히트맵")
        hm_c = monthly_avg("고객비중", ch_list()).fillna(0)
        if not hm_c.empty:
            z_vals = (hm_c.T * 100).values
            text_vals = [[f"{v:.1f}%" for v in row] for row in z_vals]
            fig = go.Figure(go.Heatmap(
                z=z_vals, x=hm_c.index.tolist(), y=hm_c.columns.tolist(),
                text=text_vals, texttemplate="%{text}",
                colorscale="Blues", showscale=True,
                colorbar=dict(title="비중%", thickness=12),
            ))
            ly = base_layout(260, title="고객 비중 (%)")
            ly["xaxis"]["tickangle"] = -30
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)

    # 고객수 vs 거래액 scatter
    with cR2:
        st.subheader("채널별 고객수 vs 거래액 (월 평균)")
        brows = []
        for ch in ch_list():
            c_m = df_full[(df_full["metric"] == "일평균고객수") & (df_full["segment"] == ch)].groupby("month")["value"].mean()
            t_m = df_full[(df_full["metric"] == "일평균거래액") & (df_full["segment"] == ch)].groupby("month")["value"].mean()
            for m_k in c_m.index.intersection(t_m.index):
                brows.append({"채널": ch, "month": m_k, "고객수": c_m[m_k], "거래액": t_m[m_k]})
        bdf = pd.DataFrame(brows)
        fig = go.Figure()
        for ch in ch_list():
            sub = bdf[bdf["채널"] == ch]
            if len(sub):
                fig.add_trace(go.Scatter(
                    x=sub["고객수"].tolist(), y=(sub["거래액"] / 1e8).tolist(),
                    mode="markers+text", name=ch,
                    text=sub["month"].tolist(), textposition="top center",
                    textfont=dict(size=9, color="#94a3b8"),
                    marker=dict(color=ch_cbg(ch), line=dict(color=ch_clr(ch), width=1.5), size=10),
                ))
        ly = base_layout(260, ysuffix="억", title="고객수 vs 거래액")
        ly["showlegend"] = True
        ly["legend"] = dict(orientation="h", y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b", size=10))
        ly["xaxis"]["title"] = dict(text="일평균 고객수", font=dict(size=11, color="#64748b"))
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 기간 비교
    st.subheader("기간 비교")
    cmp_mode, _yoy_ref = cmp_selectbox("cust")
    cmp_cur, cmp_prev, lc, lp = get_compare_periods(df_full, cmp_mode, _yoy_ref)
    tbl = compare_table(cmp_cur, cmp_prev, lc, lp, "일평균고객수", ch_list())
    if not tbl.empty:
        st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    memo_block("btm_cust")
    insight_editor("fp_cust", ["고객수 분석 인사이트를 입력하세요."])

# ══════════════════════════════════════════════════════════════════════
# 04. 채널 효율
# ══════════════════════════════════════════════════════════════════════
elif page == "04. 채널 효율":
    editable_text("fp_title_eff", "채널 효율 분석", tag="h2")
    memo_block("top_eff")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # CR 추이
    st.subheader("전환율(CR) 추이 (7일 이동평균)")
    pivot_cr = get_pivot("CR", ch_list()).rolling(7).mean()
    fig = go.Figure()
    for ch in ch_list():
        if ch in pivot_cr.columns:
            s = pivot_cr[ch].dropna() * 100
            fig.add_trace(go.Scatter(
                x=s.index.strftime("%m/%d").tolist(), y=s.tolist(),
                mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
            ))
    ly = base_layout(300, ysuffix="%", title="전환율(CR) (7일 MA)")
    ly["showlegend"] = True
    ly["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    ly["xaxis"]["nticks"] = 15
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    cL, cR = st.columns(2)

    # 평균 CR bar
    with cL:
        st.subheader("채널별 평균 CR (필터 기간)")
        avg_cr = (df_full[(df_full["metric"] == "CR") & (df_full["segment"].isin(ch_list()))]
                  .groupby("segment")["value"].mean().reindex(ch_list()).fillna(0))
        fig = go.Figure(go.Bar(
            x=avg_cr.index.tolist(), y=(avg_cr * 100).tolist(),
            marker_color=[ch_cbg(c) for c in avg_cr.index],
            marker_line_color=[ch_clr(c) for c in avg_cr.index], marker_line_width=1.2,
            text=[f"{v*100:.2f}%" for v in avg_cr.values], textposition="outside",
        ))
        ly = base_layout(260, ysuffix="%", title="평균 전환율(CR)")
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    # 평균 유입율 bar — 채널별 데이터 없으면 *TOTAL 표시
    with cR:
        st.subheader("평균 유입율 (필터 기간)")
        ir_ch = (df_full[(df_full["metric"] == "유입율") & (df_full["segment"].isin(ch_list()))]
                 .groupby("segment")["value"].mean().reindex(ch_list()))
        if ir_ch.notna().any():
            ir_data = ir_ch.fillna(0)
            x_labels = ir_data.index.tolist()
            colors_bg = [ch_cbg(c) for c in x_labels]
            colors_ln = [ch_clr(c) for c in x_labels]
        else:
            # 채널별 데이터 없음 → *TOTAL 단일 표시
            total_val = df_full[(df_full["metric"] == "유입율") & (df_full["segment"] == "*TOTAL")]["value"].mean()
            ir_data = pd.Series({"*TOTAL": total_val if not np.isnan(total_val) else 0})
            x_labels = ["*TOTAL"]; colors_bg = [ch_cbg("*TOTAL")]; colors_ln = [ch_clr("*TOTAL")]
            st.caption("채널별 유입율 데이터 없음 — 전체(*TOTAL) 기준 표시")
        fig = go.Figure(go.Bar(
            x=x_labels, y=(ir_data * 100).tolist(),
            marker_color=colors_bg, marker_line_color=colors_ln, marker_line_width=1.2,
            text=[f"{v*100:.2f}%" for v in ir_data.values], textposition="outside",
        ))
        ly = base_layout(260, ysuffix="%", title="평균 유입율")
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 객단가 추이
    st.subheader("일평균객단가 추이 (7일 이동평균)")
    pivot_aov = get_pivot("일평균객단가", ch_list()).rolling(7).mean()
    fig = go.Figure()
    for ch in ch_list():
        if ch in pivot_aov.columns:
            s = pivot_aov[ch].dropna() / 1e4
            fig.add_trace(go.Scatter(
                x=s.index.strftime("%m/%d").tolist(), y=s.tolist(),
                mode="lines", name=ch, line=dict(color=ch_clr(ch), width=1.8),
            ))
    ly = base_layout(280, ysuffix="만", title="일평균객단가 (7일 MA)")
    ly["showlegend"] = True
    ly["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
    ly["xaxis"]["nticks"] = 15
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 유입율 → CR 회귀 (채널별 유입율 데이터 있으면 채널별, 없으면 *TOTAL 폴백)
    ir_segs = avail_segs("유입율")
    reg_targets = [c for c in ch_list() if c in ir_segs] or (["*TOTAL"] if "*TOTAL" in ir_segs else [])
    st.subheader("유입율 → CR 선형회귀" + (" (전체 기준)" if reg_targets == ["*TOTAL"] else " (채널별)"))
    if reg_targets == ["*TOTAL"]:
        st.caption("채널별 유입율 데이터가 없어 전체(*TOTAL) 기준으로 분석합니다.")
    reg_rows = []
    for ch in reg_targets:
        ir = df_full[(df_full["metric"] == "유입율") & (df_full["segment"] == ch)].dropna(subset=["value"]).set_index("date")["value"]
        cr = df_full[(df_full["metric"] == "CR")     & (df_full["segment"] == ch)].dropna(subset=["value"]).set_index("date")["value"]
        common = ir.index.intersection(cr.index)
        if len(common) < 10: continue
        sl, _, r, p, _ = stats.linregress(ir[common].values.astype(float), cr[common].values.astype(float))
        reg_rows.append({
            "구분": ch, "기울기": f"{sl:.4f}", "R²": r2_label(r**2),
            "p값": sig_label(p), "판정": "✅ 유의" if p < 0.05 else "⬜ 비유의",
        })
    if reg_rows:
        st.dataframe(pd.DataFrame(reg_rows).set_index("구분"), use_container_width=True)
        st.caption("R² : 유입율이 CR 변동을 얼마나 설명하는지 (높을수록 강한 관계)")
    else:
        st.info("회귀 분석에 필요한 데이터가 부족합니다 (10일 이상 필요).")

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 효율 지표 상관 히트맵
    st.subheader("효율 지표 간 상관관계")
    eff_metrics = ["일평균거래액", "일평균고객수", "일평균객단가", "CR", "유입율"]
    corr_series = []
    for m in eff_metrics:
        s = df_full[(df_full["metric"] == m) & (df_full["segment"] == "*TOTAL")].set_index("date")["value"].sort_index()
        if len(s) > 5: corr_series.append(s.rename(METRIC_KO.get(m, m)))
    if len(corr_series) >= 2:
        corr_df = pd.concat(corr_series, axis=1).dropna()
        if len(corr_df) > 5:
            cm = corr_df.corr().values
            labels = corr_df.columns.tolist()
            text_vals = [[f"{v:.2f}" for v in row] for row in cm]
            fig = go.Figure(go.Heatmap(
                z=cm, x=labels, y=labels,
                text=text_vals, texttemplate="%{text}",
                colorscale="RdBu", zmid=0, zmin=-1, zmax=1,
                showscale=True, colorbar=dict(title="r", thickness=12),
            ))
            ly = base_layout(280, title="지표 간 상관관계 (Pearson r)")
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    stat_explainer()

    # 기간 비교
    st.subheader("기간 비교")
    cmp_mode, _yoy_ref = cmp_selectbox("eff")
    cmp_cur, cmp_prev, lc, lp = get_compare_periods(df_full, cmp_mode, _yoy_ref)
    for m in ["CR", "유입율", "일평균객단가"]:
        segs = avail_segs(m) if m == "유입율" else ch_list()
        tbl = compare_table(cmp_cur, cmp_prev, lc, lp, m, segs)
        if not tbl.empty:
            st.caption(METRIC_KO.get(m, m))
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    memo_block("btm_eff")
    insight_editor("fp_eff", ["채널 효율 인사이트를 입력하세요."])

# ══════════════════════════════════════════════════════════════════════
# 05. 채널 비중 추이
# ══════════════════════════════════════════════════════════════════════
elif page == "05. 채널 비중 추이":
    editable_text("fp_title_share", "채널 비중 추이", tag="h2")
    memo_block("top_share")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    def stacked_area_chart(metric, title, pct_label="비중 (%)"):
        pivot = get_pivot(metric, ch_list())
        pivot = pivot.ffill().fillna(0)
        # 날짜별 합계로 정규화 → 정확히 100% 합산
        row_sum = pivot.sum(axis=1).replace(0, np.nan)
        pivot = pivot.div(row_sum, axis=0).fillna(0)
        dates_str = pivot.index.strftime("%m/%d").tolist()
        fig = go.Figure()
        for ch in ch_list():
            if ch in pivot.columns:
                fig.add_trace(go.Scatter(
                    x=dates_str, y=(pivot[ch] * 100).tolist(),
                    mode="lines", name=ch, stackgroup="one",
                    line=dict(color=ch_clr(ch), width=0.5),
                    fillcolor=ch_cbg(ch),
                ))
        ly = base_layout(300, ysuffix="%", title=title)
        ly["showlegend"] = True
        ly["legend"] = dict(orientation="h", y=1.08, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b"))
        ly["xaxis"]["nticks"] = 15
        ly["yaxis"]["range"] = [0, 100]
        fig.update_layout(**ly)
        return fig

    st.subheader("거래액 비중 추이 (누적)")
    st.plotly_chart(stacked_area_chart("거래액비중", "거래액 비중 추이"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    st.subheader("고객 비중 추이 (누적)")
    st.plotly_chart(stacked_area_chart("고객비중", "고객 비중 추이"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 요일별 CR 히트맵
    st.subheader("요일별 × 채널별 전환율(CR) 패턴")
    dow_map = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}
    cr_df = df_full[(df_full["metric"] == "CR") & (df_full["segment"].isin(ch_list()))].copy()
    cr_df["요일"] = cr_df["dow"].map(dow_map)
    cr_hm = (cr_df.groupby(["요일", "segment"])["value"].mean()
             .unstack("segment")
             .reindex(index=["월","화","수","목","금","토","일"], columns=ch_list())
             .fillna(0))
    if not cr_hm.empty:
        z_vals = (cr_hm.T * 100).values
        text_vals = [[f"{v:.2f}%" for v in row] for row in z_vals]
        fig = go.Figure(go.Heatmap(
            z=z_vals, x=cr_hm.index.tolist(), y=cr_hm.columns.tolist(),
            text=text_vals, texttemplate="%{text}",
            colorscale="RdYlGn", showscale=True,
            colorbar=dict(title="CR%", thickness=12),
        ))
        ly = base_layout(280, title="요일×채널 CR (%)")
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 요일별 거래액 히트맵
    st.subheader("요일별 × 채널별 일평균거래액 패턴")
    tx_df = df_full[(df_full["metric"] == "일평균거래액") & (df_full["segment"].isin(ch_list()))].copy()
    tx_df["요일"] = tx_df["dow"].map(dow_map)
    tx_hm = (tx_df.groupby(["요일", "segment"])["value"].mean()
             .unstack("segment")
             .reindex(index=["월","화","수","목","금","토","일"], columns=ch_list())
             .fillna(0))
    if not tx_hm.empty:
        z_vals = (tx_hm.T / 1e8).values
        text_vals = [[f"{v:.2f}억" for v in row] for row in z_vals]
        fig = go.Figure(go.Heatmap(
            z=z_vals, x=tx_hm.index.tolist(), y=tx_hm.columns.tolist(),
            text=text_vals, texttemplate="%{text}",
            colorscale="YlOrRd", showscale=True,
            colorbar=dict(title="억원", thickness=12),
        ))
        ly = base_layout(280, title="요일×채널 일평균거래액 (억원)")
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

    memo_block("btm_share")
    insight_editor("fp_share", ["채널 비중 추이 인사이트를 입력하세요."])

# ══════════════════════════════════════════════════════════════════════
# 06. 월별 요약
# ══════════════════════════════════════════════════════════════════════
elif page == "06. 월별 요약":
    editable_text("fp_title_monthly", "월별 요약", tag="h2")
    memo_block("top_monthly")
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    SUMMARY_METRICS = [
        ("일평균거래액", lambda v: fmt_num(v, "원"),  1e8,  "억"),
        ("일평균고객수", lambda v: f"{v:,.0f}명",     1.0,  "명"),
        ("일평균객단가", lambda v: fmt_num(v, "원"),  1e4,  "만원"),
        ("CR",          fmt_pct,                    100.0, "%"),
    ]

    for metric, cell_fmt, divisor, y_unit in SUMMARY_METRICS:
        st.subheader(f"월별 {METRIC_KO.get(metric, metric)}")
        monthly = monthly_avg(metric, ["*TOTAL"] + ch_list())
        if monthly.empty:
            st.info("데이터 없음"); continue

        def safe_fmt(v, fmtr=cell_fmt):
            try: return fmtr(v) if not pd.isnull(v) else "–"
            except: return "–"

        st.dataframe(monthly.apply(lambda col: col.map(safe_fmt)), use_container_width=True)

        # 채널별 grouped bar
        monthly_ch = monthly.drop(columns=["*TOTAL"], errors="ignore")
        if not monthly_ch.empty:
            fig = go.Figure()
            for ch in ch_list():
                if ch in monthly_ch.columns:
                    y_vals = (monthly_ch[ch] / divisor).tolist()
                    fig.add_trace(go.Bar(
                        x=monthly_ch.index.tolist(), y=y_vals,
                        name=ch, marker_color=ch_cbg(ch),
                        marker_line_color=ch_clr(ch), marker_line_width=1,
                    ))
            ly = base_layout(260, ysuffix=y_unit, title=f"월별 {METRIC_KO.get(metric, metric)}")
            ly["barmode"] = "group"; ly["showlegend"] = True
            ly["legend"] = dict(orientation="h", y=1.1, bgcolor="rgba(0,0,0,0)", font=dict(color="#64748b", size=10))
            fig.update_layout(**ly)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 기간 비교 (전 지표)
    st.subheader("기간 비교 — 전 지표")
    cmp_mode, _yoy_ref = cmp_selectbox("monthly")
    cmp_cur, cmp_prev, lc, lp = get_compare_periods(df_full, cmp_mode, _yoy_ref)
    for m in ["일평균거래액", "일평균고객수", "일평균객단가", "CR", "유입율"]:
        segs = avail_segs(m) if m == "유입율" else ch_list()
        tbl = compare_table(cmp_cur, cmp_prev, lc, lp, m, segs)
        if not tbl.empty:
            st.caption(METRIC_KO.get(m, m))
            st.dataframe(styled_compare(tbl), use_container_width=True, hide_index=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 전체 데이터 다운로드
    st.subheader("📥 전체 데이터 다운로드")
    all_pivot = df_full.pivot_table(index=["metric","segment"], columns="date", values="value", aggfunc="first")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        all_pivot.to_excel(writer, sheet_name="전체데이터")
    st.download_button("⬇️ Excel 다운로드", data=buf.getvalue(),
                       file_name="첫구매_채널분석.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    show_appendix(df_full[["date","metric","segment","value"]].rename(
        columns={"date":"날짜","metric":"지표","segment":"채널","value":"값"}), "전체_데이터")

    memo_block("btm_monthly")
    insight_editor("fp_monthly", ["월별 요약 인사이트를 입력하세요."])
