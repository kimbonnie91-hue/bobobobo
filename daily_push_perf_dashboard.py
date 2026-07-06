# -*- coding: utf-8 -*-
"""
일일 PUSH 발송성과 대시보드
─────────────────────────────────────────────────────────────
쿼리로 추출해 매일 갱신하는 "일일 PUSH 실적.xlsx"(누적_소재별 / 소재별 실적(당주)_2 시트)를
업로드하면 AF코드 기준으로 발송모수/UV/주문건수/거래액을 집계해 시각화한다.

· BPU별 실적, 발송유형별(기본/우수) 실적, 주차별 누적 추이는 "BPU별 실적"/"발송유형별 실적"
  피벗 시트를 그대로 읽지 않고, 소재별 실적 원본 로우 데이터에서 매번 다시 집계한다.
  (원본 피벗 시트의 셀 레이아웃이 바뀌어도 항상 정확한 값을 보장하기 위함)
· 업로드할 때마다 (일자+AF코드+시간대+타겟구분) 키로 누적 저장 — 구글시트(설정 시) ↔ 로컬 CSV 폴백.

데이터 로직(파싱/집계 함수)은 Streamlit 비의존 순수 함수이며 모듈 import 만으로 테스트 가능하다.
앱 UI 는 main() 안에 있고 `python -m streamlit run daily_push_perf_dashboard.py` 시에만 실행된다.
"""
import io, os, re, datetime
import numpy as np
import pandas as pd

# ══════════════════════════════════════════════════════════════════════
# 1. 순수 데이터 로직 (Streamlit 비의존)
# ══════════════════════════════════════════════════════════════════════

# 시트/구글시트 헤더가 "일자(8자리)" "타겟구분" "기획전No." 처럼 표기가 흔들려도
# 매칭되도록 괄호 안 내용을 지우고 영문/숫자/한글만 남겨 비교한다.
def _norm_h(s):
    s = str(s if s is not None else "")
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^0-9A-Za-z가-힣]", "", s)
    return s


COLMAP_CANDIDATES = {
    "date":    ["일자(8자리)", "일자", "날짜"],
    "dow_k":   ["요일"],
    "hour":    ["시간대", "시간"],
    "target":  ["타겟구분", "타겟 구분", "타겟"],
    "stype":   ["발송유형"],
    "bpu":     ["BPU"],
    "prio":    ["우선순위"],
    "cat":     ["카테고리"],
    "attr":    ["속성"],
    "owner":   ["담당자"],
    "brand":   ["브랜드"],
    "af":      ["AF코드", "AF 코드"],
    "promo":   ["기획전No.", "기획전 No.", "기획전No", "기획전"],
    "send":    ["발송모수", "발송건수", "발송"],
    "uv":      ["UV"],
    "visit":   ["VISIT", "방문"],
    "cust":    ["고객수"],
    "oc":      ["주문건수"],
    "amt":     ["거래액", "주문금액", "GMV"],
    "infl_cr": ["유입전환율"],
    "ord_cr":  ["주문전환율"],
    "eff":     ["효율"],
}
NUM_COLS = ["prio", "hour", "send", "uv", "visit", "cust", "oc", "amt", "infl_cr", "ord_cr", "eff"]
TXT_COLS = ["dow_k", "target", "stype", "bpu", "cat", "attr", "owner", "brand", "promo"]
_ALL_CAND_NORMS = {_norm_h(c) for cands in COLMAP_CANDIDATES.values() for c in cands}


def _norm_date(v):
    """엑셀 셀 값 → 'YYYYMMDD' 문자열 또는 None. datetime/문자열/구분자 포함 문자열 모두 허용."""
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y%m%d")
    digits = re.sub(r"\D", "", str(v).strip())
    return digits if len(digits) == 8 else None


def _find_header_row(rows, max_scan=10, min_score=3):
    """rows 앞부분에서 COLMAP 후보와 가장 많이 일치하는 행을 헤더로 판정."""
    best_i, best_score = None, min_score - 1
    for i, row in enumerate(rows[:max_scan]):
        cells = [_norm_h(v) for v in (row or [])]
        score = sum(1 for c in cells if c and c in _ALL_CAND_NORMS)
        if score > best_score:
            best_score, best_i = score, i
    return best_i


def _map_columns(header_cells):
    """헤더 셀 리스트 → {표준키: 열 인덱스}."""
    normed = [_norm_h(c) for c in header_cells]
    idx = {}
    for key, cands in COLMAP_CANDIDATES.items():
        for cand in cands:
            cn = _norm_h(cand)
            if cn in normed:
                idx[key] = normed.index(cn)
                break
    return idx


def parse_material_rows(rows):
    """소재별 실적 시트(2차원 rows) → AF코드 단위 레코드 DataFrame."""
    if not rows:
        return pd.DataFrame()
    hdr_i = _find_header_row(rows)
    if hdr_i is None:
        return pd.DataFrame()
    header = [("" if v is None else str(v)) for v in rows[hdr_i]]
    idx = _map_columns(header)
    if "af" not in idx:
        return pd.DataFrame()
    af_i = idx["af"]
    recs = []
    for row in rows[hdr_i + 1:]:
        if row is None:
            continue
        n = len(row)
        af_raw = row[af_i] if af_i < n else None
        af = "" if af_raw is None else str(af_raw).strip()
        if not af or af.lower() in ("nan", "none", "-"):
            continue
        rec = {key: (row[i] if i < n else None) for key, i in idx.items()}
        rec["af"] = af
        rec["date"] = _norm_date(rec.get("date"))
        recs.append(rec)
    return pd.DataFrame(recs)


def discover_perf_sheets(sheetnames):
    """워크북 시트명 목록 → (누적_소재별 시트명, 소재별 실적(당주) 시트명). 없으면 None."""
    cum = week = None
    for s in sheetnames:
        sn = re.sub(r"\s+", "", s)
        if "누적" in sn and "소재" in sn:
            cum = cum or s
        elif "소재별" in sn and "실적" in sn:
            week = week or s
    return cum, week


def load_excel_perf(file_bytes):
    """업로드 엑셀 바이트 → (합쳐진 DataFrame, 실제로 읽은 시트명 리스트).

    우선 '누적_소재별'과 '소재별 실적(당주)_*' 시트를 찾아 읽는다.
    둘 다 없으면 전체 시트를 훑어 AF코드+발송/거래액 헤더가 있는 시트를 폴백으로 사용한다.
    (발송 건수/RawNew 같은 원본 쿼리 덤프 시트는 이미 소재별 실적 시트로 집계되어 있으므로 읽지 않는다.)
    """
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    try:
        names = wb.sheetnames
        cum_name, week_name = discover_perf_sheets(names)
        frames, used = [], []
        for sname in (cum_name, week_name):
            if not sname:
                continue
            rows = list(wb[sname].iter_rows(values_only=True))
            d = parse_material_rows(rows)
            if not d.empty:
                frames.append(d)
                used.append(sname)
        if not frames:
            for s in names:
                head_rows = list(wb[s].iter_rows(values_only=True, max_row=10))
                hdr_i = _find_header_row(head_rows)
                if hdr_i is None:
                    continue
                header = [("" if v is None else str(v)) for v in head_rows[hdr_i]]
                idx = _map_columns(header)
                if "af" in idx and ("amt" in idx or "send" in idx):
                    rows = list(wb[s].iter_rows(values_only=True))
                    d = parse_material_rows(rows)
                    if not d.empty:
                        frames.append(d)
                        used.append(s)
        if not frames:
            return pd.DataFrame(), []
        combined = pd.concat(frames, ignore_index=True)
        key_cols = [c for c in ("date", "af", "hour", "target") if c in combined.columns]
        if key_cols:
            combined = combined.drop_duplicates(subset=key_cols, keep="last")
        return combined, used
    finally:
        wb.close()


def _finalize(df):
    """숫자/텍스트 정리 + 파생지표(ctr/cvr/rps/aov) + 주차(월~일) 라벨 계산."""
    cols = list(COLMAP_CANDIDATES.keys())
    if df is None or df.empty:
        extra = ["dt", "week_start", "week_end", "week_label", "dow", "ctr", "cvr", "rps", "aov"]
        return pd.DataFrame(columns=cols + extra)
    df = df.copy()
    df["af"] = df["af"].astype(str).str.strip()
    df["date"] = df["date"].astype(str).str.replace(r"\.0$", "", regex=True)
    for c in NUM_COLS:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in TXT_COLS:
        if c in df:
            df[c] = df[c].apply(lambda v: "" if v is None else str(v).strip())
        else:
            df[c] = ""
    df["dt"] = pd.to_datetime(df["date"], format="%Y%m%d", errors="coerce")
    df = df.dropna(subset=["dt"]).reset_index(drop=True)
    wk = df["dt"].dt.to_period("W-SUN")
    df["week_start"] = wk.dt.start_time.dt.date
    df["week_end"] = wk.dt.end_time.dt.date
    df["week_label"] = df.apply(
        lambda r: f"{r['week_start'].month}/{r['week_start'].day}~{r['week_end'].month}/{r['week_end'].day}",
        axis=1)
    df["dow"] = df["dt"].dt.dayofweek
    send = df["send"].fillna(0); uv = df["uv"].fillna(0); oc = df["oc"].fillna(0); amt = df["amt"].fillna(0)
    df["ctr"] = np.where(send > 0, uv / send, 0.0)
    df["cvr"] = np.where(uv > 0, oc / uv, 0.0)
    df["rps"] = np.where(send > 0, amt / send, 0.0)
    df["aov"] = np.where(oc > 0, amt / oc, 0.0)
    return df


def agg_metrics(df, group_cols):
    """group_cols 기준 발송/UV/주문건수/거래액 합계 + 파생 효율지표."""
    if df is None or df.empty:
        return pd.DataFrame(columns=list(group_cols) + ["n", "send", "uv", "oc", "amt", "ctr", "cvr", "rps", "aov"])
    g = df.groupby(list(group_cols), dropna=False).agg(
        n=("af", "count"), send=("send", "sum"), uv=("uv", "sum"), oc=("oc", "sum"), amt=("amt", "sum"),
    ).reset_index()
    g["ctr"] = np.where(g["send"] > 0, g["uv"] / g["send"], 0.0)
    g["cvr"] = np.where(g["uv"] > 0, g["oc"] / g["uv"], 0.0)
    g["rps"] = np.where(g["send"] > 0, g["amt"] / g["send"], 0.0)
    g["aov"] = np.where(g["oc"] > 0, g["amt"] / g["oc"], 0.0)
    return g


# ── 누적 저장소 — 업로드할 때마다 (일자+AF코드+시간대+타겟구분) 키로 병합 ─────
DATA_STORE = "daily_push_perf_store.csv"
STORE_COLS = list(COLMAP_CANDIDATES.keys())
STORE_KEY_COLS = ["date", "af", "hour", "target"]


def store_key_frame(d):
    k = pd.DataFrame(index=d.index)
    for c in STORE_KEY_COLS:
        if c in d.columns:
            k[c] = d[c].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        else:
            k[c] = ""
    return k


def merge_store(old, new):
    """기존 누적 + 신규 업로드 병합 — 같은 키는 신규 우선."""
    def _pick(d):
        if d is None or len(d) == 0:
            return pd.DataFrame(columns=STORE_COLS)
        return d[[c for c in STORE_COLS if c in d]].copy()
    both = pd.concat([_pick(old), _pick(new)], ignore_index=True)
    if both.empty:
        return both
    both["date"] = both["date"].astype(str).str.replace(r"\.0$", "", regex=True)
    keep = ~store_key_frame(both).duplicated(keep="last")
    return both.loc[keep].reset_index(drop=True)


def load_store():
    if os.path.exists(DATA_STORE):
        try:
            return pd.read_csv(DATA_STORE, encoding="utf-8-sig", dtype={"date": str, "af": str})
        except Exception:
            pass
    return pd.DataFrame(columns=STORE_COLS)


def save_store(df):
    df[[c for c in STORE_COLS if c in df]].to_csv(DATA_STORE, index=False, encoding="utf-8-sig")


# ── 구글시트 영속 저장 (선택) — 미설정 시 로컬 CSV 폴백 ──────────────────
GS_TITLE = "daily_push_perf_store"


def _fix_pem(pk):
    """Secrets에서 깨진 private_key 복구 — 어떤 형태든 base64 본문을 추출해 표준 PEM으로 재구성."""
    if not isinstance(pk, str):
        return pk
    s = pk.strip().strip('"').strip("'")
    s = s.replace("\\r\\n", "\n").replace("\\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    mt = re.search(r"-----BEGIN ([A-Z0-9 ]+?)-----(.*?)-----END \1-----", s, re.S)
    if mt:
        header, body = mt.group(1).strip(), re.sub(r"\s+", "", mt.group(2))
    else:
        header = "PRIVATE KEY"
        body = re.sub(r"\s+", "", s)
    if not body:
        return s + "\n"
    wrapped = "\n".join(body[i:i + 64] for i in range(0, len(body), 64))
    return f"-----BEGIN {header}-----\n{wrapped}\n-----END {header}-----\n"


def _pem_diag(pk):
    if not isinstance(pk, str) or not pk.strip():
        return "키가 비어있음(미입력)"
    has_b = "-----BEGIN" in pk
    has_e = "-----END" in pk
    body = re.sub(r"\s+", "", re.sub(r"-----[A-Z0-9 ]+-----", "", pk))
    n = len(body)
    tip = "정상 길이" if n >= 1500 else "너무 짧음 → 잘린 듯"
    return f"BEGIN:{'있음' if has_b else '없음'} END:{'있음' if has_e else '없음'} 본문:{n}자({tip})"


def gs_open(creds_dict, spreadsheet):
    """서비스 계정 자격으로 스프레드시트 열기 (URL/키/제목 모두 허용)."""
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    info = dict(creds_dict)
    info["private_key"] = _fix_pem(info.get("private_key"))
    try:
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    except Exception as e:
        raise ValueError(f"서비스계정 private_key 문제 — {_pem_diag(info.get('private_key'))}. "
                         f"(원본오류: {str(e)[:50]})")
    gc = gspread.authorize(creds)
    sp = str(spreadsheet).strip()
    if sp.startswith("http"):
        return gc.open_by_url(sp)
    if "/" not in sp and " " not in sp and len(sp) >= 30:
        return gc.open_by_key(sp)
    try:
        return gc.open(sp)
    except Exception:
        return gc.open_by_key(sp)


def gs_read_ws(sh, title, cols):
    try:
        ws = sh.worksheet(title)
    except Exception:
        return pd.DataFrame(columns=cols)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]


def gs_write_ws(sh, title, df, cols):
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    out = out[cols].fillna("").astype(str)
    try:
        ws = sh.worksheet(title)
    except Exception:
        ws = sh.add_worksheet(title=title, rows=max(len(out) + 10, 100), cols=max(len(cols), 10))
    ws.clear()
    ws.update(values=[cols] + out.values.tolist(), range_name="A1")


def gs_clear_ws(sh, title, cols):
    try:
        ws = sh.worksheet(title)
        ws.clear()
        ws.update(values=[cols], range_name="A1")
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════
# 2. Streamlit 앱
# ══════════════════════════════════════════════════════════════════════
METRIC_LABELS = {"send": "발송모수", "uv": "UV", "oc": "주문건수", "amt": "거래액"}
RATE_LABELS = {"ctr": "CTR(UV/발송)", "cvr": "주문전환율(주문/UV)", "rps": "RPS(발송당거래액)", "aov": "객단가(거래액/주문)"}


def main():
    import streamlit as st
    import plotly.graph_objects as go

    st.set_page_config(page_title="일일 PUSH 발송성과 대시보드", layout="wide", initial_sidebar_state="expanded")
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"]{background:#f8f9fc}
    [data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
    [data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
    [data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
    [data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
    h1,h2,h3{color:#1e293b}
    </style>""", unsafe_allow_html=True)

    def won(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return "–"
        if abs(v) >= 1e8:
            return f"{v/1e8:.2f}억"
        if abs(v) >= 1e4:
            return f"{v/1e4:,.0f}만"
        return f"{v:,.0f}"

    def base_layout(h=340, title=""):
        return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#475569", size=11), height=h,
                    margin=dict(l=10, r=10, t=(72 if title else 20), b=10),
                    title=dict(text=title, font=dict(color="#94a3b8", size=13), x=0, xanchor="left",
                               y=0.99, yanchor="top"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.1, xanchor="left", x=0),
                    xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0"),
                    yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0"))

    # ── 저장소 백엔드: 구글시트(설정 시) ↔ 로컬 CSV(폴백) ──
    @st.cache_resource(show_spinner=False)
    def _get_sh(_email, spreadsheet):
        return gs_open(st.secrets["gcp_service_account"], spreadsheet)

    def init_storage():
        try:
            has = "gcp_service_account" in st.secrets
        except Exception:
            has = False
        if not has:
            return {"mode": "local", "status": "💾 로컬 CSV로 저장해요 (구글시트 미설정)"}
        try:
            sp = None
            if "gsheets" in st.secrets:
                sp = st.secrets["gsheets"].get("spreadsheet")
            sp = sp or st.secrets.get("gsheets_spreadsheet")
            if not sp:
                return {"mode": "local", "status": "⚠️ gsheets.spreadsheet 미설정 → 로컬에 저장해요"}
            sh = _get_sh(st.secrets["gcp_service_account"].get("client_email", ""), sp)
            return {"mode": "gsheets", "sh": sh, "status": "☁️ 구글시트에 연결됐어요"}
        except Exception as e:
            return {"mode": "local", "status": f"⚠️ 구글시트 연결 실패 → 로컬에 저장해요 ({str(e)[:50]})"}

    def storage_load(bk):
        if bk["mode"] == "gsheets":
            try:
                return gs_read_ws(bk["sh"], GS_TITLE, STORE_COLS)
            except Exception:
                return pd.DataFrame(columns=STORE_COLS)
        return load_store()

    def storage_save(bk, df):
        if bk["mode"] == "gsheets":
            gs_write_ws(bk["sh"], GS_TITLE, df, STORE_COLS)
        else:
            save_store(df)

    def storage_clear(bk):
        if bk["mode"] == "gsheets":
            gs_clear_ws(bk["sh"], GS_TITLE, STORE_COLS)
        elif os.path.exists(DATA_STORE):
            os.remove(DATA_STORE)

    BK = init_storage()

    @st.cache_data(show_spinner=False)
    def cached_load_excel(b):
        return load_excel_perf(b)

    # ══════════════════════════════════════════════════════════
    # 사이드바 — 업로드 / 필터
    # ══════════════════════════════════════════════════════════
    st.sidebar.markdown("### 📲 일일 PUSH 발송성과")
    st.sidebar.caption(BK["status"])

    uploaded = st.sidebar.file_uploader(
        "📂 일일 PUSH 실적.xlsx 업로드", type=["xlsx"], accept_multiple_files=True, key="perf_up",
        help="'누적_소재별' 또는 '소재별 실적(당주)_2' 시트를 자동으로 찾아 읽어요. "
             "'발송 건수'/'RawNew' 원본 쿼리 시트는 이미 저 시트로 집계되어 있어 읽지 않아요.")

    c1, c2 = st.sidebar.columns(2)
    apply_clicked = c1.button("📥 반영하기", use_container_width=True, disabled=not uploaded)
    clear_clicked = c2.button("🗑️ 초기화", use_container_width=True)

    if clear_clicked:
        storage_clear(BK)
        st.sidebar.success("저장된 데이터를 초기화했어요.")

    if apply_clicked and uploaded:
        old = storage_load(BK)
        frames, all_used = [], []
        for f in uploaded:
            d, used = cached_load_excel(f.getvalue())
            if not d.empty:
                frames.append(d)
                all_used.extend([f"{f.name} · {u}" for u in used])
        if not frames:
            st.sidebar.error("인식 가능한 시트를 찾지 못했어요. '누적_소재별'/'소재별 실적(당주)' 시트가 있는지 확인해 주세요.")
        else:
            new = pd.concat(frames, ignore_index=True)
            merged = merge_store(old, new)
            storage_save(BK, merged)
            st.sidebar.success(f"{len(new):,}건 반영 완료 (누적 {len(merged):,}건)")
            with st.sidebar.expander("읽은 시트"):
                for u in all_used:
                    st.caption(u)

    raw = storage_load(BK)
    df = _finalize(raw)

    if df.empty:
        st.title("일일 PUSH 발송성과 대시보드")
        st.info("왼쪽에서 '일일 PUSH 실적.xlsx'를 업로드하고 **반영하기**를 눌러 주세요.")
        with st.expander("기대하는 파일 형식"):
            st.markdown("""
- 시트명에 **'누적_소재별'** 또는 **'소재별 실적(당주)'**가 포함된 시트를 찾습니다 (`_2` 등 접미사 있어도 인식).
- 헤더 컬럼: 일자(8자리)/요일/시간대/타겟구분/발송유형/BPU/우선순위/카테고리/속성/담당자/브랜드/AF코드/기획전No. +
  발송(모수)/UV/VISIT/고객수/주문건수/거래액(주문금액)/유입전환율/주문전환율/효율
- AF코드가 비어있는 행은 자동으로 제외돼요.
            """)
        return

    st.sidebar.markdown("---")
    st.sidebar.markdown("**필터**")
    dmin, dmax = df["dt"].min().date(), df["dt"].max().date()
    date_range = st.sidebar.date_input("기간", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    bpu_sel = st.sidebar.multiselect("BPU", sorted([b for b in df["bpu"].unique() if b]))
    stype_sel = st.sidebar.multiselect("발송유형", sorted([s for s in df["stype"].unique() if s]))
    search = st.sidebar.text_input("검색 (브랜드/AF코드/담당자)", "")

    f = df.copy()
    if isinstance(date_range, tuple) and len(date_range) == 2:
        f = f[(f["dt"].dt.date >= date_range[0]) & (f["dt"].dt.date <= date_range[1])]
    if bpu_sel:
        f = f[f["bpu"].isin(bpu_sel)]
    if stype_sel:
        f = f[f["stype"].isin(stype_sel)]
    if search:
        q = search.lower()
        hay = (f["brand"] + " " + f["af"] + " " + f["owner"]).str.lower()
        f = f[hay.str.contains(re.escape(q), na=False)]

    st.sidebar.caption(f"필터 후 {len(f):,}건 · 전체 {len(df):,}건")

    page = st.sidebar.radio("페이지", ["종합요약", "BPU별 실적", "발송유형별 실적", "주차별 누적 추이", "AF코드별 리더보드", "데이터"])

    # ══════════════════════════════════════════════════════════
    # 종합요약
    # ══════════════════════════════════════════════════════════
    if page == "종합요약":
        st.title("종합 요약")
        tot = agg_metrics(f, ["stype"]).drop(columns=["stype"]).sum(numeric_only=True) if not f.empty else None
        send_t = f["send"].sum(); uv_t = f["uv"].sum(); oc_t = f["oc"].sum(); amt_t = f["amt"].sum()
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("발송모수", f"{send_t:,.0f}")
        c2.metric("UV", f"{uv_t:,.0f}")
        c3.metric("주문건수", f"{oc_t:,.0f}")
        c4.metric("거래액", won(amt_t))
        c5, c6, c7 = st.columns(3)
        c5.metric("CTR(UV/발송)", f"{(uv_t/send_t if send_t else 0):.2%}")
        c6.metric("주문전환율(주문/UV)", f"{(oc_t/uv_t if uv_t else 0):.2%}")
        c7.metric("RPS(발송당거래액)", f"{(amt_t/send_t if send_t else 0):,.0f}")

        st.markdown("---")
        metric = st.selectbox("추이 지표", list(METRIC_LABELS), format_func=lambda k: METRIC_LABELS[k], key="sum_metric")
        daily = f.groupby(f["dt"].dt.date)[metric].sum().reset_index()
        fig = go.Figure(go.Scatter(x=daily["dt"], y=daily[metric], mode="lines+markers", line=dict(color="#4f8fff", width=2)))
        fig.update_layout(**base_layout(title=f"일자별 {METRIC_LABELS[metric]} 추이"))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 전주 대비")
        last_end = f["dt"].max().normalize()
        this_start = last_end - pd.Timedelta(days=6)
        prev_start = this_start - pd.Timedelta(days=7)
        prev_end = this_start - pd.Timedelta(days=1)
        cur = f[(f["dt"] >= this_start) & (f["dt"] <= last_end)]
        prev = f[(f["dt"] >= prev_start) & (f["dt"] <= prev_end)]
        rows = []
        for k, label in METRIC_LABELS.items():
            cv, pv = cur[k].sum(), prev[k].sum()
            delta = (cv - pv) / pv * 100 if pv else np.nan
            rows.append({"지표": label, "이번주(최근7일)": cv, "전주": pv, "증감%": delta})
        st.dataframe(pd.DataFrame(rows).style.format({"이번주(최근7일)": "{:,.0f}", "전주": "{:,.0f}", "증감%": "{:+.1f}%"}),
                     use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════
    # BPU별 실적
    # ══════════════════════════════════════════════════════════
    elif page == "BPU별 실적":
        st.title("BPU별 실적")
        st.caption("소재별 실적 원본 데이터를 BPU 기준으로 실시간 집계한 화면이에요.")
        g = agg_metrics(f, ["bpu"])
        g = g[g["bpu"] != ""].sort_values("amt", ascending=False)
        metric = st.selectbox("정렬/차트 지표", list(METRIC_LABELS), format_func=lambda k: METRIC_LABELS[k], key="bpu_metric")
        gs = g.sort_values(metric, ascending=False)
        fig = go.Figure(go.Bar(x=gs[metric], y=gs["bpu"], orientation="h", marker_color="#4f8fff"))
        fig.update_layout(**base_layout(h=max(300, 28 * len(gs)), title=f"BPU별 {METRIC_LABELS[metric]}"))
        fig.update_yaxes(autorange="reversed")
        st.plotly_chart(fig, use_container_width=True)

        show = gs.rename(columns={**METRIC_LABELS, **RATE_LABELS, "bpu": "BPU", "n": "캠페인수"})
        st.dataframe(show.style.format({
            "발송모수": "{:,.0f}", "UV": "{:,.0f}", "주문건수": "{:,.0f}", "거래액": "{:,.0f}",
            "CTR(UV/발송)": "{:.2%}", "주문전환율(주문/UV)": "{:.2%}", "RPS(발송당거래액)": "{:,.0f}", "객단가(거래액/주문)": "{:,.0f}",
        }), use_container_width=True, hide_index=True)

        st.markdown("#### BPU 드릴다운")
        pick = st.selectbox("BPU 선택", gs["bpu"].tolist())
        drill = agg_metrics(f[f["bpu"] == pick], ["brand"]).sort_values("amt", ascending=False).head(15)
        st.dataframe(drill.rename(columns={**METRIC_LABELS, "brand": "브랜드", "n": "캠페인수"}),
                     use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════
    # 발송유형별 실적
    # ══════════════════════════════════════════════════════════
    elif page == "발송유형별 실적":
        st.title("발송유형별 실적")
        st.caption("소재별 실적 원본 데이터를 발송유형 기준으로 실시간 집계한 화면이에요.")
        has_basic = f["stype"].str.contains("기본", na=False).any()
        has_good = f["stype"].str.contains("우수", na=False).any()
        if has_basic and has_good:
            basic = f[f["stype"].str.contains("기본", na=False)]
            good = f[f["stype"].str.contains("우수", na=False)]
            st.markdown("#### 기본발송 vs 우수발송")
            cols = st.columns(4)
            for i, (k, label) in enumerate(METRIC_LABELS.items()):
                bv, gv = basic[k].sum(), good[k].sum()
                cols[i].metric(label, f"{bv:,.0f} / {gv:,.0f}", help="기본발송 / 우수발송")

        g = agg_metrics(f, ["stype"])
        g = g[g["stype"] != ""].sort_values("amt", ascending=False)
        metric = st.selectbox("정렬/차트 지표", list(METRIC_LABELS), format_func=lambda k: METRIC_LABELS[k], key="stype_metric")
        gs = g.sort_values(metric, ascending=False)
        fig = go.Figure(go.Bar(x=gs["stype"], y=gs[metric], marker_color="#48bb78"))
        fig.update_layout(**base_layout(title=f"발송유형별 {METRIC_LABELS[metric]}"))
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 발송유형 × BPU")
        pivot = f.pivot_table(index="bpu", columns="stype", values=metric, aggfunc="sum", fill_value=0)
        st.dataframe(pivot.style.format("{:,.0f}"), use_container_width=True)

    # ══════════════════════════════════════════════════════════
    # 주차별 누적 추이
    # ══════════════════════════════════════════════════════════
    elif page == "주차별 누적 추이":
        st.title("주차별 누적 추이")
        st.caption("월~일 기준 주차로 묶어 집계했어요 (누적_소재별 시트 기반).")
        g = f.groupby(["week_start", "week_label"], dropna=False).agg(
            send=("send", "sum"), uv=("uv", "sum"), oc=("oc", "sum"), amt=("amt", "sum")).reset_index()
        g = g.sort_values("week_start")
        metric = st.selectbox("지표", list(METRIC_LABELS), format_func=lambda k: METRIC_LABELS[k], key="week_metric")
        g["누적"] = g[metric].cumsum()
        g["WoW%"] = g[metric].pct_change() * 100

        fig = go.Figure()
        fig.add_bar(x=g["week_label"], y=g[metric], name=f"주간 {METRIC_LABELS[metric]}", marker_color="#4f8fff")
        fig.add_trace(go.Scatter(x=g["week_label"], y=g["누적"], name="누적", yaxis="y2",
                                 mode="lines+markers", line=dict(color="#ed8936", width=2)))
        layout = base_layout(title=f"주차별 {METRIC_LABELS[metric]} · 누적 추이")
        layout["yaxis2"] = dict(overlaying="y", side="right", showgrid=False)
        fig.update_layout(**layout)
        st.plotly_chart(fig, use_container_width=True)

        show = g.rename(columns={"week_label": "주차", **METRIC_LABELS})
        st.dataframe(show[["주차", "발송모수", "UV", "주문건수", "거래액", "누적", "WoW%"]]
                     .style.format({"발송모수": "{:,.0f}", "UV": "{:,.0f}", "주문건수": "{:,.0f}",
                                    "거래액": "{:,.0f}", "누적": "{:,.0f}", "WoW%": "{:+.1f}%"}),
                     use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════
    # AF코드별 리더보드
    # ══════════════════════════════════════════════════════════
    elif page == "AF코드별 리더보드":
        st.title("AF코드별 리더보드")
        g = f.groupby("af", dropna=False).agg(
            브랜드=("brand", "first"), BPU=("bpu", "first"), 발송유형=("stype", "first"),
            send=("send", "sum"), uv=("uv", "sum"), oc=("oc", "sum"), amt=("amt", "sum"),
        ).reset_index()
        g["ctr"] = np.where(g["send"] > 0, g["uv"] / g["send"], 0.0)
        g["cvr"] = np.where(g["uv"] > 0, g["oc"] / g["uv"], 0.0)
        g["rps"] = np.where(g["send"] > 0, g["amt"] / g["send"], 0.0)
        metric = st.selectbox("정렬 지표", list(METRIC_LABELS), format_func=lambda k: METRIC_LABELS[k], key="af_metric")
        g = g.sort_values(metric, ascending=False)

        top = g.head(20)
        fig = go.Figure(go.Bar(x=top[metric], y=top["af"] + " · " + top["브랜드"], orientation="h", marker_color="#7b5bc0"))
        fig.update_layout(**base_layout(h=max(320, 24 * len(top)), title=f"상위 20 · {METRIC_LABELS[metric]}"))
        fig.update_yaxes(autorange="reversed")
        st.plotly_chart(fig, use_container_width=True)

        show = g.rename(columns={"af": "AF코드", **METRIC_LABELS,
                                 "ctr": "CTR(UV/발송)", "cvr": "주문전환율(주문/UV)", "rps": "RPS(발송당거래액)"})
        st.dataframe(show.style.format({
            "발송모수": "{:,.0f}", "UV": "{:,.0f}", "주문건수": "{:,.0f}", "거래액": "{:,.0f}",
            "CTR(UV/발송)": "{:.2%}", "주문전환율(주문/UV)": "{:.2%}", "RPS(발송당거래액)": "{:,.0f}",
        }), use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════
    # 데이터
    # ══════════════════════════════════════════════════════════
    else:
        st.title("데이터 · 다운로드")
        st.caption(f"누적 저장 {len(df):,}건 · 필터 적용 {len(f):,}건")
        st.dataframe(f.drop(columns=["dt"]), use_container_width=True, height=500)
        st.download_button("⬇️ CSV 다운로드", f.drop(columns=["dt"]).to_csv(index=False).encode("utf-8-sig"),
                          file_name="daily_push_perf_filtered.csv", mime="text/csv")


if __name__ == "__main__":
    main()
