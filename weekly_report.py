"""주간보고 — 첫구매 통합 실적 대시보드
여러 원천 엑셀(전체관점 마스터 + 지표별 파일)을 업로드하면 하나의 통합 뷰로 종합하고,
전년(YoY)·전주(WoW) 증감현황과 보고란을 갖춘 주간보고 화면을 만든다.
통합 결과는 (월/주/첫구매_요약 + 차트) 엑셀 워크북으로 다운로드할 수 있다.
"""

import datetime
import io, json, os, re
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

try:
    from streamlit_quill import st_quill
    HAS_QUILL = True
except Exception:
    HAS_QUILL = False

# ══════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════
st.set_page_config(page_title="주간보고 — 첫구매 통합 실적", page_icon="📋",
                   layout="wide", initial_sidebar_state="expanded")

INSIGHT_FILE = "wr_insights.json"

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f8f9fc}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e2e8f0}
[data-testid="stMetric"]{background:#ffffff;border-radius:8px;padding:12px 16px;border:1px solid #e2e8f0;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
[data-testid="stMetricLabel"]{color:#64748b!important;font-size:12px!important}
[data-testid="stMetricValue"]{color:#1e293b!important;font-size:20px!important}
.sdiv{border-top:1px solid #e2e8f0;margin:22px 0}
.report-box{border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;margin:8px 0;line-height:1.7;background:#ffffff}
.report-box p{margin:0 0 4px}
.kpi-card{background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
.kpi-label{color:#64748b;font-size:12px}
.kpi-value{color:#1e293b;font-size:21px;font-weight:600;margin:2px 0 8px}
.kpi-delta{display:block;width:fit-content;font-size:12px;border-radius:6px;padding:2px 8px;margin:4px 0 0;font-weight:500;white-space:nowrap}
.kpi-delta.up{background:#ecfdf5;color:#15803d}
.kpi-delta.down{background:#fef2f2;color:#dc2626}
.kpi-delta.na{background:#f1f5f9;color:#94a3b8}
@media print {
  [data-testid="stSidebar"], [data-testid="stHeader"], [data-testid="stToolbar"],
  header, .stButton, .no-print, iframe { display:none !important; }
  [data-testid="stAppViewContainer"], .main, .block-container { background:#fff !important; }
  .block-container { max-width:100% !important; padding-top:0 !important; }
  .stPlotlyChart, .report-box, table { break-inside:avoid; }
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════
# 색상 팔레트 (기존 first_purchase 와 동일 계열)
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
}
def clr(n): return PALETTE.get(n, PALETTE["blue"])[0]
def cbg(n): return PALETTE.get(n, PALETTE["blue"])[1]

CHANNEL_PAL = {
    "직접": "blue", "광고": "amber", "EP": "green", "PUSH": "purple",
    "제휴": "red", "브랜드광고": "teal", "미디어커머스": "orange", "*TOTAL": "slate",
}
CHANNELS = ["직접", "광고", "EP", "PUSH", "제휴", "브랜드광고", "미디어커머스"]
YEAR_PAL = ["slate", "blue", "red", "green", "purple", "amber", "teal"]

# ══════════════════════════════════════════════════════
# 지표 정의
# ══════════════════════════════════════════════════════
METRICS7 = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가",
            "비회원트래픽", "가입자수", "가입율", "당일가입CR"]
PCT_METRICS = {"가입율", "당일가입CR", "유입율", "CR", "거래액비중", "고객비중"}
# 마스터 파일 지표 → 보고서 지표 매핑
MASTER_MAP = {"일평균거래액": "첫구매 거래액", "일평균고객수": "첫구매 고객수",
              "일평균객단가": "첫구매 객단가"}
# 지표별 파일명 → 보고서 지표 매핑 (공백 제거 후 매칭)
METRIC_FILE_MAP = {
    "가입율": "가입율", "가입률": "가입율",
    "가입자수": "가입자수",
    "당일가입첫구매율": "당일가입CR", "당일가입CR": "당일가입CR",
    "비회원트래픽": "비회원트래픽",
}

METRIC_UNIT = {
    "첫구매 거래액": ("백만원", 1e6), "첫구매 고객수": ("명", 1),
    "첫구매 객단가": ("원", 1), "비회원트래픽": ("명", 1),
    "가입자수": ("명", 1), "가입율": ("%", 1), "당일가입CR": ("%", 1),
}

def fmt_value(metric, v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "–"
    if metric in PCT_METRICS: return f"{v*100:.2f}%"
    if metric == "첫구매 거래액": return f"{v/1e6:,.1f}백만원"
    if metric == "첫구매 객단가": return f"{v:,.0f}원"
    return f"{int(v):,}명"  # 명 단위(고객수 등)는 소수점 버림

def fmt_delta(metric, cur, prev):
    """전년비/전주비 문자열: 비율 지표는 %p 차이, 그 외 증감율"""
    if cur is None or prev is None: return None
    if isinstance(cur, float) and np.isnan(cur): return None
    if isinstance(prev, float) and np.isnan(prev): return None
    if metric in PCT_METRICS:
        d = (cur - prev) * 100
        return f"△{abs(d):.2f}%p" if d < 0 else f"+{d:.2f}%p"
    if prev == 0: return None
    d = (cur - prev) / prev * 100
    return f"△{abs(d):.1f}%" if d < 0 else f"+{d:.1f}%"

def style_delta_cols(tbl):
    """증감 컬럼(△/+)에 빨강/초록 색상 적용한 Styler 반환"""
    delta_cols = [c for c in tbl.columns
                  if any(k in str(c) for k in ("전년비", "전주비", "전월비", "증감"))]
    def _color(v):
        s = str(v)
        if s.startswith("△"): return "color:#dc2626;font-weight:600"
        if s.startswith("+"): return "color:#16a34a;font-weight:600"
        return ""
    try:
        return tbl.style.map(_color, subset=delta_cols)
    except Exception:
        return tbl

# ══════════════════════════════════════════════════════
# 파싱 — 원천 파일 → 통합 long DataFrame
# ══════════════════════════════════════════════════════
YEAR_RE   = re.compile(r"^(20\d{2})(\.0)?$")
PERIOD_RE = re.compile(r"^\s*(\d{1,2}\s*월(\s*\d\s*주차)?|\d{1,2}/\d{1,2})\s*$")

def detect_file(name):
    """파일명에서 (kind, granularity 힌트, metric) 감지.
    단위(일/주/월)와 마스터 여부는 parse_file에서 내용으로 최종 판별하므로 여기선 힌트만."""
    base = os.path.basename(name)
    base = re.sub(r"\.(xlsx|xls|csv)$", "", base, flags=re.I)
    key = base.replace(" ", "")
    # 단위 힌트: 일자별/주별/월별 키워드 또는 일_/주_/월_ 접두사
    gran = None
    if "일자별" in key or "데일리" in key: gran = "일"
    elif "주별" in key or "주간" in key:   gran = "주"
    elif "월별" in key or "월간" in key:   gran = "월"
    else:
        m = re.match(r"^(일|주|월)[_\s]", base)
        if m: gran = m.group(1)
    if "전체관점" in key or "마스터" in key:
        return "master", gran, None
    # 지표 키워드는 파일명 어느 위치에 있어도 인식
    for k, v in METRIC_FILE_MAP.items():
        if k.replace(" ", "") in key:
            return "metric", gran, v
    # 접두사형(월_xxx)인데 모르는 지표면 정리된 이름 그대로 사용
    m = re.match(r"^(일|주|월)[_\s]+(.+)$", base)
    if m:
        rest = re.sub(r"\(.*?\)", "", m.group(2)).strip()
        return "metric", m.group(1), rest or "기타"
    return None, gran, None

def _decode_text(data: bytes) -> str:
    for enc in ("utf-16", "utf-8-sig", "cp949", "utf-8"):
        try: return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError): pass
    return data.decode("utf-8", "replace")

def read_grid(name, data: bytes):
    """csv(UTF-16/탭 포함)·xlsx 모두 2차원 셀 그리드로 읽는다"""
    if name.lower().endswith((".csv", ".txt", ".tsv")):
        text = _decode_text(data)
        return [line.split("\t") for line in text.splitlines()]
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]

def _cell(v):
    return "" if v is None else str(v).strip()

def _num(v):
    """셀 값 → float (콤마·% 처리, %는 비율로 변환)"""
    if v is None: return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "–"): return np.nan
    pct = s.endswith("%")
    if pct: s = s[:-1]
    try: f = float(s)
    except ValueError: return np.nan
    return f / 100 if pct else f

def period_parts(gran, year, plabel):
    """기간 라벨 → (표준라벨, 정렬키)"""
    p = plabel.replace(" ", "")
    if gran == "월":
        m = re.match(r"(\d{1,2})월", p)
        if not m: return None
        mo = int(m.group(1))
        return f"{mo}월", year * 10000 + mo * 100
    if gran == "주":
        m = re.match(r"(\d{1,2})월(\d)주차", p)
        if not m: return None
        mo, wk = int(m.group(1)), int(m.group(2))
        return f"{mo:02d}월 {wk}주차", year * 10000 + mo * 100 + wk
    m = re.match(r"(\d{1,2})/(\d{1,2})", p)
    if not m: return None
    mo, dd = int(m.group(1)), int(m.group(2))
    return f"{mo}/{dd}", year * 10000 + mo * 100 + dd

def parse_file(name, data: bytes) -> pd.DataFrame:
    # 누적 데이터 백업 CSV(long 포맷) 재업로드 시 그대로 복원
    if name.lower().endswith(".csv"):
        first = data[:400].decode("utf-8", "ignore").splitlines()[0] if data else ""
        if "gran" in first and "metric" in first and "sortkey" in first:
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if set(STORE_COLS) <= set(d.columns):
                    return d[STORE_COLS]
            except Exception:
                pass
    kind, gran, file_metric = detect_file(name)
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows: return pd.DataFrame()

    # 헤더 행 탐색 (앞 6행): 연도 / 기간라벨 / 마감구분
    year_row = period_row = close_row = None
    for ri in range(min(6, len(rows))):
        cells = [_cell(c) for c in rows[ri]]
        nY = sum(1 for c in cells if YEAR_RE.match(c))
        nP = sum(1 for c in cells if PERIOD_RE.match(c))
        nC = sum(1 for c in cells if "마감" in c)
        if year_row is None and nY >= 1 and nP == 0: year_row = ri
        if period_row is None and nP >= 2: period_row = ri
        if close_row is None and nC >= 2: close_row = ri
    if period_row is None: return pd.DataFrame()
    data_start = max(r for r in (year_row, period_row, close_row) if r is not None) + 1

    # ── 내용 기반 최종 판별 (파일명은 힌트일 뿐)
    # 단위: 기간 라벨 형태가 결정 (N월 N주차 → 주, N/N → 일, N월 → 월)
    plabels = [_cell(c) for c in rows[period_row] if PERIOD_RE.match(_cell(c))]
    if any("주차" in p for p in plabels):  gran = "주"
    elif any("/" in p for p in plabels):   gran = "일"
    else:                                   gran = "월"
    # 마스터 여부: 헤더에 '구분01'이 있으면 마스터(전체관점) 구조
    if any("구분01" in _cell(c) for ri in range(data_start) for c in rows[ri]):
        kind = "master"
    if kind is None or (kind == "metric" and not file_metric):
        return pd.DataFrame()

    ncols = max(len(r) for r in rows)
    def cell(ri, ci):
        return _cell(rows[ri][ci]) if ri is not None and ci < len(rows[ri]) else ""

    # 컬럼별 연도(좌측 ffill)·기간·마감
    col_year, cur_y = {}, None
    for ci in range(ncols):
        m = YEAR_RE.match(cell(year_row, ci)) if year_row is not None else None
        if m: cur_y = int(m.group(1))
        col_year[ci] = cur_y
    # 기간 라벨이 빈 셀(병합)인 일마감/MTD 컬럼은 직전 라벨을 이어받는다
    data_cols, col_label, last_lbl = [], {}, None
    for ci in range(ncols):
        lbl = cell(period_row, ci)
        if PERIOD_RE.match(lbl) and col_year[ci]:
            last_lbl = lbl
            col_label[ci] = lbl
            data_cols.append(ci)
        elif (close_row is not None and "마감" in cell(close_row, ci)
              and last_lbl and col_year[ci]):
            col_label[ci] = last_lbl
            data_cols.append(ci)

    seg_col = 1 if kind == "master" else 0
    records, cur_metric = [], None
    for ri in range(data_start, len(rows)):
        if kind == "master":
            m0 = cell(ri, 0)
            if m0 and m0 not in ("-", "–"): cur_metric = m0
            metric = MASTER_MAP.get(cur_metric, cur_metric)
        else:
            metric = file_metric
        seg = cell(ri, seg_col)
        if not seg or seg in ("-", "–") or metric is None: continue
        for ci in data_cols:
            pp = period_parts(gran, col_year[ci], col_label[ci])
            if pp is None: continue
            label, sortkey = pp
            close = cell(close_row, ci) if close_row is not None else ""
            records.append({
                "gran": gran, "metric": metric, "segment": seg,
                "year": col_year[ci], "label": label, "sortkey": sortkey,
                "close": "mtd" if "일마감" in close and gran != "일" else "final",
                "value": _num(rows[ri][ci] if ci < len(rows[ri]) else None),
            })
    return pd.DataFrame(records)

def _zip_entry_name(info):
    """zip 내 한글 파일명 복원 (UTF-8 플래그 없으면 cp437→cp949 재해석)"""
    if info.flag_bits & 0x800:
        return info.filename
    for enc in ("cp949", "utf-8", "euc-kr"):
        try: return info.filename.encode("cp437").decode(enc)
        except (UnicodeDecodeError, UnicodeEncodeError): pass
    return info.filename

def expand_uploads(uploads):
    """업로드 목록 → (이름, bytes) 목록. zip은 풀어서 내부 엑셀/CSV를 꺼낸다"""
    out = []
    for f in uploads:
        data = f.getvalue()
        if f.name.lower().endswith(".zip"):
            import zipfile
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
            except zipfile.BadZipFile:
                continue
            with zf:
                for info in zf.infolist():
                    if info.is_dir(): continue
                    name = _zip_entry_name(info)
                    if "__MACOSX" in name: continue
                    base = os.path.basename(name)
                    if base.startswith((".", "~")): continue
                    if not base.lower().endswith((".xlsx", ".xls", ".csv")): continue
                    out.append((base, zf.read(info)))
        else:
            out.append((f.name, data))
    return out

@st.cache_data(show_spinner=False)
def combine_files(file_tuples) -> pd.DataFrame:
    """업로드 파일들 → 통합 long DF. 동일 키는 마지막 파일 우선"""
    frames = [parse_file(n, b) for n, b in file_tuples]
    frames = [f for f in frames if not f.empty]
    if not frames: return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    df = df.drop_duplicates(subset=KEY_COLS, keep="last")
    return df

# ══════════════════════════════════════════════════════
# 데이터 누적 저장소 — 업로드할 때마다 병합·저장
# ══════════════════════════════════════════════════════
DATA_STORE = "wr_data_store.csv"
KEY_COLS = ["gran", "metric", "segment", "year", "label", "close"]
STORE_COLS = KEY_COLS + ["sortkey", "value"]

def load_store() -> pd.DataFrame:
    if os.path.exists(DATA_STORE):
        try:
            d = pd.read_csv(DATA_STORE, encoding="utf-8-sig")
            if set(STORE_COLS) <= set(d.columns):
                return d[STORE_COLS]
        except Exception:
            pass
    return pd.DataFrame()

def save_store(df: pd.DataFrame):
    df[STORE_COLS].to_csv(DATA_STORE, index=False, encoding="utf-8-sig")

def merge_store(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """기존 누적 + 신규 업로드 병합 — 같은 (단위·지표·채널·기간) 키는 신규 우선"""
    if old is None or old.empty: return new
    if new is None or new.empty: return old
    return (pd.concat([old[STORE_COLS], new[STORE_COLS]], ignore_index=True)
            .drop_duplicates(subset=KEY_COLS, keep="last"))

# ══════════════════════════════════════════════════════
# 회원 실적 데이터셋 (신규/기존 세그먼트 × 9지표, 월 단위)
# ══════════════════════════════════════════════════════
MEMBER_STORE = "wr_member_store.csv"
MEMBER_KEY = ["metric", "segment", "year", "label"]
MEMBER_COLS = MEMBER_KEY + ["sortkey", "value"]
MEMBER_SEGS = ["Total", "당월신규", "기가입신규", "기존"]
MEMBER_METRICS = ["유효회원수", "UV", "월방문율(%)", "구매고객수", "CR(%)",
                  "객단가", "거래액 (VAT제외)", "유효회원당 거래액", "활동고객수"]
MEMBER_PCT = {"월방문율(%)", "CR(%)"}

def is_member_grid(rows):
    """헤더에 회원구분/회원구분상세 + 유효회원수 지표가 있으면 회원 실적 파일"""
    head = " ".join(_cell(c) for ri in range(min(3, len(rows))) for c in rows[ri])
    col0 = " ".join(_cell(rows[ri][0]) for ri in range(len(rows)) if rows[ri])
    return ("회원구분" in head and "회원구분상세" in head
            and "유효회원수" in col0)

def fmt_member(metric, v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "–"
    if metric in MEMBER_PCT: return f"{v*100:.2f}%"
    if metric == "거래액 (VAT제외)": return f"{v/1e8:,.1f}억"
    if metric in ("객단가", "유효회원당 거래액"): return f"{v:,.0f}원"
    return f"{int(v):,}"

def parse_member_file(name, data: bytes) -> pd.DataFrame:
    # 회원 백업 CSV 복원
    if name.lower().endswith(".csv"):
        first = data[:300].decode("utf-8", "ignore").splitlines()[0] if data else ""
        if "metric" in first and "segment" in first and "sortkey" in first and "gran" not in first:
            try:
                d = pd.read_csv(io.BytesIO(data), encoding="utf-8-sig")
                if set(MEMBER_COLS) <= set(d.columns) and set(d["segment"]) <= set(MEMBER_SEGS) | {"Total"}:
                    return d[MEMBER_COLS]
            except Exception:
                pass
    try:
        rows = read_grid(name, data)
    except Exception:
        return pd.DataFrame()
    if not rows or not is_member_grid(rows): return pd.DataFrame()

    # 헤더: 0행 연도(ffill), 1행 차원명+월라벨. 데이터는 2행부터
    year_row, label_row = rows[0], rows[1]
    # 월 라벨이 시작되는 컬럼 = '가입연차' 다음 (보통 7열)
    data_c0 = 7
    col_year, cur = {}, None
    for ci in range(len(year_row)):
        m = re.search(r"(20\d{2})", _cell(year_row[ci]))
        if m: cur = int(m.group(1))
        col_year[ci] = cur
    data_cols = [ci for ci in range(data_c0, len(label_row))
                 if re.match(r"^\d{1,2}월$", _cell(label_row[ci])) and col_year.get(ci)]

    records, cur_metric = [], None
    for ri in range(2, len(rows)):
        row = rows[ri]
        m0 = _cell(row[0]) if len(row) > 0 else ""
        if m0: cur_metric = m0
        if cur_metric not in MEMBER_METRICS: continue
        seg = _cell(row[2]) if len(row) > 2 else ""   # 회원구분상세
        # 등급/성별/연령대/가입연차(3~6열)는 Total인 행만 (현재 분해 없음)
        dims = [_cell(row[ci]) if len(row) > ci else "" for ci in range(3, 7)]
        if any(d and d != "Total" for d in dims): continue
        if seg not in MEMBER_SEGS: continue
        for ci in data_cols:
            mo = int(re.match(r"(\d{1,2})월", _cell(label_row[ci])).group(1))
            yr = col_year[ci]
            records.append({"metric": cur_metric, "segment": seg, "year": yr,
                            "label": f"{mo}월", "sortkey": yr * 100 + mo,
                            "value": _num(row[ci] if ci < len(row) else None)})
    return pd.DataFrame(records)

def combine_member(file_tuples) -> pd.DataFrame:
    frames = [parse_member_file(n, b) for n, b in file_tuples]
    frames = [f for f in frames if not f.empty]
    if not frames: return pd.DataFrame()
    return (pd.concat(frames, ignore_index=True)
            .drop_duplicates(subset=MEMBER_KEY, keep="last"))

def load_member_store() -> pd.DataFrame:
    if os.path.exists(MEMBER_STORE):
        try:
            d = pd.read_csv(MEMBER_STORE, encoding="utf-8-sig")
            if set(MEMBER_COLS) <= set(d.columns): return d[MEMBER_COLS]
        except Exception:
            pass
    return pd.DataFrame()

def save_member_store(df): df[MEMBER_COLS].to_csv(MEMBER_STORE, index=False, encoding="utf-8-sig")

def merge_member(old, new):
    if old is None or old.empty: return new
    if new is None or new.empty: return old
    return (pd.concat([old[MEMBER_COLS], new[MEMBER_COLS]], ignore_index=True)
            .drop_duplicates(subset=MEMBER_KEY, keep="last"))

def _period_set(d, cols):
    if d is None or d.empty: return set()
    return set(map(tuple, d[cols].drop_duplicates().values.tolist()))

def upload_diff(stored, df_new, member_stored, member_new):
    """업로드 데이터가 기존 누적 대비 추가/갱신하는 기간 수 (저장 전 미리보기용)"""
    cols_c, cols_m = ["gran", "year", "label"], ["year", "label"]
    nc, oc = _period_set(df_new, cols_c), _period_set(stored, cols_c)
    nm, om = _period_set(member_new, cols_m), _period_set(member_stored, cols_m)
    added = len(nc - oc) + len(nm - om)
    updated = len(nc & oc) + len(nm & om)
    return added, updated


def member_pick(mdf, metric, seg, year, mo):
    s = mdf[(mdf["metric"] == metric) & (mdf["segment"] == seg) &
            (mdf["year"] == year) & (mdf["label"] == f"{mo}월")]["value"].dropna()
    return s.iloc[-1] if len(s) else np.nan

def member_series(mdf, metric, seg, year):
    s = mdf[(mdf["metric"] == metric) & (mdf["segment"] == seg) & (mdf["year"] == year)]
    return s.sort_values("sortkey").set_index("label")["value"]


# ══════════════════════════════════════════════════════
# 조회 헬퍼 — mtd(당월/당주 일마감) vs final 선택
# ══════════════════════════════════════════════════════
def pick(df, gran, metric, seg, year, label, prefer="final"):
    sub = df[(df["gran"] == gran) & (df["metric"] == metric) &
             (df["segment"] == seg) & (df["year"] == year) & (df["label"] == label)]
    if sub.empty: return np.nan
    order = ["final", "mtd"] if prefer == "final" else ["mtd", "final"]
    for c in order:
        s = sub[sub["close"] == c]["value"].dropna()
        if len(s): return s.iloc[-1]
    return np.nan

def series_by_label(df, gran, metric, seg, year, prefer="final"):
    """한 연도의 기간라벨 → 값 Series (sortkey 순)"""
    sub = df[(df["gran"] == gran) & (df["metric"] == metric) &
             (df["segment"] == seg) & (df["year"] == year)].copy()
    if sub.empty: return pd.Series(dtype=float)
    pref = {"final": 0, "mtd": 1} if prefer == "final" else {"mtd": 0, "final": 1}
    sub["_p"] = sub["close"].map(pref)
    sub = sub.sort_values(["sortkey", "_p"]).drop_duplicates("label", keep="first")
    return sub.set_index("label")["value"]

def labels_sorted(df, gran, years=None):
    sub = df[df["gran"] == gran]
    if years is not None: sub = sub[sub["year"].isin(years)]
    return (sub[["label", "sortkey"]].assign(k=lambda d: d["sortkey"] % 10000)
            .drop_duplicates("label").sort_values("k")["label"].tolist())

def latest_period(df, gran):
    """가장 최근 (year, label)"""
    sub = df[(df["gran"] == gran) & df["value"].notna()]
    if sub.empty: return None, None
    row = sub.loc[sub["sortkey"].idxmax()]
    return int(row["year"]), row["label"]

def prev_label(df, gran, year, label):
    """직전 기간 (연도 경계 포함)"""
    sub = (df[(df["gran"] == gran) & df["value"].notna()]
           [["year", "label", "sortkey"]].drop_duplicates().sort_values("sortkey"))
    keys = sub[["year", "label"]].apply(tuple, axis=1).tolist()
    try: i = keys.index((year, label))
    except ValueError: return None, None
    return keys[i - 1] if i > 0 else (None, None)

# ══════════════════════════════════════════════════════
# 보고란 영속화
# ══════════════════════════════════════════════════════
def load_insights():
    if os.path.exists(INSIGHT_FILE):
        with open(INSIGHT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_insights(d):
    with open(INSIGHT_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)

def report_text_block(key, title, default="", regen=None, ai_fn=None):
    """편집 가능한 보고란 텍스트 박스 (JSON 저장).
    regen 텍스트를 주면 '자동 생성'(템플릿) 버튼, ai_fn을 주면 'AI 생성'(Claude) 버튼이 뜬다.
    좁은 컬럼 안에서도 안 깨지도록 버튼은 박스 위/아래에 배치한다."""
    store = st.session_state.wr_texts
    if not store.get(key): store[key] = default
    ekey = f"__wr_edit_{key}__"
    if ekey not in st.session_state: st.session_state[ekey] = False

    # 제목 + 액션 버튼 (제목 한 줄, 버튼은 그 아래 정상 너비 컬럼)
    st.markdown(f"**{title}**")
    n = 2 + (1 if regen is not None else 0) + (1 if ai_fn is not None else 0)
    bcols = st.columns(n)
    bi = 0
    edit_on = st.session_state[ekey]
    if bcols[bi].button("편집" if not edit_on else "보기",
                        key=f"wr_edit_{key}", use_container_width=True):
        st.session_state[ekey] = not edit_on; st.rerun()
    bi += 1
    if regen is not None:
        if bcols[bi].button("자동 생성", key=f"wr_regen_{key}", use_container_width=True,
                            help="기준 주차 실적으로 템플릿 문구를 채웁니다 (기존 내용 대체)"):
            store[key] = regen
            all_d = load_insights(); all_d[key] = regen; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
        bi += 1
    if ai_fn is not None:
        if bcols[bi].button("AI 생성", key=f"wr_ai_{key}", use_container_width=True,
                            help="Claude가 데이터를 보고 인사이트 문구를 작성합니다 (기존 내용 대체)"):
            with st.spinner("AI가 인사이트를 작성 중…"):
                text, err = ai_fn()
            if err:
                st.error(err)
            else:
                store[key] = text
                all_d = load_insights(); all_d[key] = text; save_insights(all_d)
                st.session_state[ekey] = False; st.rerun()

    if st.session_state[ekey]:
        if HAS_QUILL:
            # Word 수준 리치 에디터 (글자 크기·색·굵게·기울임·밑줄·목록·정렬 등)
            toolbar = [
                [{"size": ["small", False, "large", "huge"]}],
                ["bold", "italic", "underline", "strike"],
                [{"color": []}, {"background": []}],
                [{"list": "ordered"}, {"list": "bullet"}],
                [{"align": []}], ["clean"],
            ]
            new = st_quill(value=store[key], html=True, toolbar=toolbar,
                           key=f"wr_quill_{key}")
        else:
            new = st.text_area("", store[key], key=f"wr_ta_{key}", height=180,
                               label_visibility="collapsed")
        if st.button("저장", key=f"wr_save_{key}", type="primary",
                     use_container_width=True):
            store[key] = new if new is not None else store[key]
            all_d = load_insights(); all_d[key] = store[key]; save_insights(all_d)
            st.session_state[ekey] = False; st.rerun()
    else:
        st.markdown(f"<div class='report-box'>{store[key] or '내용을 입력하세요.'}</div>",
                    unsafe_allow_html=True)
    return store[key]

# ══════════════════════════════════════════════════════
# YoY 요약표 / 추이표 빌더
# ══════════════════════════════════════════════════════
def month_label(n): return f"{n}월"

def week_disp(year, label):
    """주차 표시: 2026년 06월 2주차"""
    return f"{year}년 {label}" if label else "-"

def wow_summary_table(df, wy, wlabel, metrics):
    """전주비 요약표 (실적 요약과 동일 구조): 전주·기준주·전주비 + 전년동주·전년비"""
    py, plb = prev_label(df, "주", wy, wlabel)
    cols = [week_disp(py, plb), week_disp(wy, wlabel), "전주비",
            week_disp(wy - 1, wlabel), "전년비"]
    rows = []
    for met in metrics:
        cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
        prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
        yoy = pick(df, "주", met, "*TOTAL", wy - 1, wlabel, "final")
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, prv), cols[1]: fmt_value(met, cur),
            cols[2]: fmt_delta(met, cur, prv) or "–",
            cols[3]: fmt_value(met, yoy), cols[4]: fmt_delta(met, cur, yoy) or "–",
        })
    return pd.DataFrame(rows).set_index("구분")

def yoy_summary_table(df, ref_year, ref_month, metrics):
    """참고본 '실적 요약' 표: 전월·당월 × (전년, 당년, 전년비)"""
    rows = []
    pm_y, pm_m = (ref_year, ref_month - 1) if ref_month > 1 else (ref_year - 1, 12)
    cols = [f"{pm_y-1}년 {pm_m}월", f"{pm_y}년 {pm_m}월", "전년비(전월)",
            f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월", "전년비(당월)"]
    for met in metrics:
        pm_prev = pick(df, "월", met, "*TOTAL", pm_y - 1, month_label(pm_m), "final")
        pm_cur  = pick(df, "월", met, "*TOTAL", pm_y,     month_label(pm_m), "final")
        cm_prev = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
        cm_cur  = pick(df, "월", met, "*TOTAL", ref_year,     month_label(ref_month), "mtd")
        rows.append({
            "구분": met,
            cols[0]: fmt_value(met, pm_prev), cols[1]: fmt_value(met, pm_cur),
            cols[2]: fmt_delta(met, pm_cur, pm_prev) or "–",
            cols[3]: fmt_value(met, cm_prev), cols[4]: fmt_value(met, cm_cur),
            cols[5]: fmt_delta(met, cm_cur, cm_prev) or "–",
        })
    return pd.DataFrame(rows).set_index("구분"), (pm_y, pm_m)

def trend_table(df, gran, metrics, years, seg="*TOTAL"):
    """추이표: 행=지표, 열=(연도, 기간)"""
    out, columns = {}, []
    for y in years:
        for lb in labels_sorted(df, gran, [y]):
            sub = df[(df["gran"] == gran) & (df["year"] == y) & (df["label"] == lb) &
                     df["value"].notna()]
            if sub.empty: continue
            columns.append((y, lb))
    if not columns:
        return pd.DataFrame()
    for met in metrics:
        vals = []
        for y, lb in columns:
            vals.append(pick(df, gran, met, seg, y, lb, "final"))
        out[met] = vals
    tbl = pd.DataFrame(out, index=pd.MultiIndex.from_tuples(columns, names=["연도", "기간"])).T
    return tbl

def style_trend(tbl, metrics):
    # 최신 pandas는 float 컬럼에 문자열 대입을 금지하므로 object로 변환 후 포맷
    disp = tbl.astype(object).copy()
    for met in disp.index:
        disp.loc[met] = [fmt_value(met, v) for v in tbl.loc[met]]
    return disp

# ══════════════════════════════════════════════════════
# YoY 라인차트 (Plotly)
# ══════════════════════════════════════════════════════
def base_layout(h=300, ysuffix="", title=""):
    return dict(
        paper_bgcolor="rgba(248,249,252,0)", plot_bgcolor="rgba(248,249,252,0)",
        font=dict(color="#475569", size=11), margin=dict(l=10, r=10, t=40, b=10),
        height=h, showlegend=True,
        legend=dict(orientation="h", y=1.12, bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#64748b", size=10)),
        title=dict(text=title, font=dict(color="#94a3b8", size=13)),
        xaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10)),
        yaxis=dict(gridcolor="#f1f5f9", linecolor="#e2e8f0",
                   tickfont=dict(color="#64748b", size=10), ticksuffix=ysuffix),
    )

def yoy_chart(df, gran, metric, years, seg="*TOTAL", h=300):
    unit, div = METRIC_UNIT.get(metric, ("", 1))
    if metric in PCT_METRICS: div, unit = 0.01, "%"
    if gran == "월":
        x_all = [month_label(i) for i in range(1, 13)]
    else:
        x_all = labels_sorted(df, gran, years)
    fig = go.Figure()
    for i, y in enumerate(sorted(years)):
        # 연도마다 없는 주차(5주차 등)는 건너뛰고 선을 잇는다
        s = series_by_label(df, gran, metric, seg, y, prefer="final").reindex(x_all).dropna()
        fig.add_trace(go.Scatter(
            x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=str(y),
            line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
            marker=dict(size=5),
        ))
    gname = "월별" if gran == "월" else "주차별"
    ly = base_layout(h, ysuffix=unit if unit == "%" else "",
                     title=f"{metric} {gname} 추이 ({unit})")
    ly["xaxis"]["categoryorder"] = "array"
    ly["xaxis"]["categoryarray"] = x_all
    if gran == "주": ly["xaxis"]["tickangle"] = -45; ly["xaxis"]["nticks"] = 20
    fig.update_layout(**ly)
    return fig

# ══════════════════════════════════════════════════════
# 엑셀 워크북 내보내기 (월/주/첫구매_요약 + 차트)
# ══════════════════════════════════════════════════════
def build_workbook(df, texts, ref_year, ref_month, chart_years):
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    head_font = Font(bold=True, size=11)
    head_fill = PatternFill("solid", fgColor="EEF3FA")
    title_font = Font(bold=True, size=13)

    # ── 데이터 시트 (월/주): 지표 블록 스택
    extra_metrics = [m for m in df["metric"].unique() if m not in METRICS7]
    for gran, sheet in (("월", "월"), ("주", "주")):
        ws = wb.create_sheet(sheet)
        sub_g = df[df["gran"] == gran]
        if sub_g.empty: continue
        cols = (sub_g[["year", "label", "sortkey"]].drop_duplicates()
                .sort_values("sortkey")[["year", "label"]].apply(tuple, axis=1).tolist())
        r = 1
        for met in METRICS7 + sorted(extra_metrics):
            sub_m = sub_g[sub_g["metric"] == met]
            if sub_m.empty: continue
            ws.cell(r, 1, met).font = title_font
            for j, (y, lb) in enumerate(cols):
                ws.cell(r, 2 + j, y).font = head_font
                ws.cell(r + 1, 2 + j, lb).font = head_font
                ws.cell(r + 1, 2 + j).fill = head_fill
            segs = ["*TOTAL"] + [s for s in CHANNELS if s in set(sub_m["segment"])]
            for i, seg in enumerate(segs):
                ws.cell(r + 2 + i, 1, seg).font = head_font
                for j, (y, lb) in enumerate(cols):
                    v = pick(df, gran, met, seg, y, lb, "final")
                    c = ws.cell(r + 2 + i, 2 + j)
                    if not (isinstance(v, float) and np.isnan(v)):
                        c.value = v
                        c.number_format = "0.00%" if met in PCT_METRICS else "#,##0"
            r += 2 + len(segs) + 2

    # ── 요약 시트
    ws = wb.create_sheet("첫구매_요약", 0)
    wb.remove(wb["Sheet"])
    r = 1
    ws.cell(r, 1, f"첫구매 주간보고 — {ref_year}년 {ref_month}월 기준").font = Font(bold=True, size=15)
    r += 2

    # 실적 요약 YoY 표
    ws.cell(r, 1, "실적 요약 (일평균)").font = title_font; r += 1
    tbl, _ = yoy_summary_table(df, ref_year, ref_month, METRICS7)
    ws.cell(r, 1, "구분").font = head_font
    for j, cname in enumerate(tbl.columns):
        c = ws.cell(r, 2 + j, cname); c.font = head_font; c.fill = head_fill
    red_font = Font(color="DC2626")
    green_font = Font(color="16A34A")
    for i, met in enumerate(tbl.index):
        ws.cell(r + 1 + i, 1, met).font = head_font
        for j, cname in enumerate(tbl.columns):
            c = ws.cell(r + 1 + i, 2 + j, tbl.loc[met, cname])
            s = str(c.value)
            if "전년비" in str(cname):
                if s.startswith("△"): c.font = red_font
                elif s.startswith("+"): c.font = green_font
    r += len(tbl) + 3

    # 보고란
    ws.cell(r, 1, "전주 주요 지표 현황").font = title_font
    ws.cell(r, 8, "금주 집행 내용 요약").font = title_font
    c1 = ws.cell(r + 1, 1, texts.get("wr_metrics_summary", ""))
    c2 = ws.cell(r + 1, 8, texts.get("wr_exec_summary", ""))
    for c in (c1, c2): c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 8, end_column=6)
    ws.merge_cells(start_row=r + 1, start_column=8, end_row=r + 8, end_column=13)
    r += 11

    # 차트 데이터 블록 + 라인차트 (월별/주차별 × 거래액·고객수·객단가)
    chart_metrics = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]
    for gran, gname in (("월", "월별"), ("주", "주차별")):
        x_all = ([month_label(i) for i in range(1, 13)] if gran == "월"
                 else labels_sorted(df, gran, chart_years))
        if not x_all: continue
        anchor_row = r
        for k, met in enumerate(chart_metrics):
            ws.cell(r, 1, f"{met} {gname} (차트 데이터)").font = head_font
            for j, lb in enumerate(x_all):
                ws.cell(r, 2 + j, lb).fill = head_fill
            nrow = 0
            for y in sorted(chart_years):
                s = series_by_label(df, gran, met, "*TOTAL", y).reindex(x_all)
                if s.dropna().empty: continue
                nrow += 1
                ws.cell(r + nrow, 1, y)
                for j, v in enumerate(s.tolist()):
                    if not (isinstance(v, float) and np.isnan(v)):
                        c = ws.cell(r + nrow, 2 + j, v)
                        c.number_format = "#,##0"
            if nrow:
                ch = LineChart()
                ch.title = f"{met} {gname} 추이"
                ch.height, ch.width = 7.5, 13
                data = Reference(ws, min_col=1, min_row=r + 1,
                                 max_col=1 + len(x_all), max_row=r + nrow)
                cats = Reference(ws, min_col=2, min_row=r, max_col=1 + len(x_all))
                ch.add_data(data, titles_from_data=True, from_rows=True)
                ch.set_categories(cats)
                ws.add_chart(ch, f"{get_column_letter(2 + len(x_all) + 1 + (k % 3) * 8)}{anchor_row}")
            r += nrow + 2
        r += 14

    # 채널별 실적 (당월 YoY)
    ws.cell(r, 1, f"채널별 실적 — {ref_year}년 {ref_month}월 (전년비)").font = title_font; r += 1
    for met in chart_metrics:
        ws.cell(r, 1, met).font = head_font
        heads = [f"{ref_year-1}년 {ref_month}월", f"{ref_year}년 {ref_month}월", "전년비"]
        for j, hd in enumerate(heads):
            c = ws.cell(r, 2 + j, hd); c.font = head_font; c.fill = head_fill
        segs = ["*TOTAL"] + CHANNELS
        for i, seg in enumerate(segs):
            pv = pick(df, "월", met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "월", met, seg, ref_year, month_label(ref_month), "mtd")
            ws.cell(r + 1 + i, 1, seg)
            ws.cell(r + 1 + i, 2, None if np.isnan(pv) else pv).number_format = "#,##0"
            ws.cell(r + 1 + i, 3, None if np.isnan(cv) else cv).number_format = "#,##0"
            dcell = ws.cell(r + 1 + i, 4, fmt_delta(met, cv, pv) or "–")
            ds = str(dcell.value)
            if ds.startswith("△"): dcell.font = red_font
            elif ds.startswith("+"): dcell.font = green_font
        r += len(segs) + 3

    ws.column_dimensions["A"].width = 16
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════
# 자동 보고 초안
# ══════════════════════════════════════════════════════
def _delta_html(d):
    """증감 문자열 → 색상 span (역신장 △ 빨강 / 신장 + 초록)"""
    if not d: return ""
    if d.startswith("△"): return f'<span style="color:#dc2626;font-weight:700">{d}</span>'
    return f'<span style="color:#16a34a;font-weight:700">{d}</span>'

def auto_draft(df, ref_year, ref_month, ref_week=None):
    """기준 주차(없으면 기준 월) 실적으로 보고 문구 자동 생성 (HTML, 역신장 빨강)"""
    lines = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "주", met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "주", ref_year, ref_week)
            prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "주", met, "*TOTAL", ref_year - 1, ref_week, "final")
            parts = [f" - {met} — {fmt_value(met, cur)}"]
            d_w = fmt_delta(met, cur, prv)
            d_y = fmt_delta(met, cur, yoy)
            if d_w: parts.append(f"전주비 {_delta_html(d_w)}")
            if d_y: parts.append(f"전년비 {_delta_html(d_y)}")
            lines.append(", ".join(parts))
        else:
            cur = pick(df, "월", met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            d = fmt_delta(met, cur, prv)
            tail = f", 전년비 {_delta_html(d)}" if d else ""
            lines.append(f" - {met} — {fmt_value(met, cur)}" + tail)
    return "<br>".join(lines)

# ══════════════════════════════════════════════════════
# AI 인사이트 생성 (Claude API) — 모델 교체 가능
# ══════════════════════════════════════════════════════
AI_MODELS = {
    "Claude Sonnet 4.6 (균형·기본)": "claude-sonnet-4-6",
    "Claude Opus 4.8 (최고 품질)": "claude-opus-4-8",
    "Claude Haiku 4.5 (빠름·저렴)": "claude-haiku-4-5",
}

def _anthropic_key():
    """Streamlit secrets 또는 환경변수에서 API 키 조회"""
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY")

def _ai_metric_facts(df, ref_year, ref_month, ref_week=None):
    """모델에 넘길 지표·증감 요약 텍스트"""
    rows = []
    for met in METRICS7:
        if ref_week:
            cur = pick(df, "주", met, "*TOTAL", ref_year, ref_week, "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            py, plb = prev_label(df, "주", ref_year, ref_week)
            prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
            yoy = pick(df, "주", met, "*TOTAL", ref_year - 1, ref_week, "final")
            wm = re.match(r"(\d{1,2})월 (\d)주차", ref_week)
            mom = np.nan
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                my, mm = (ref_year, mo - 1) if mo > 1 else (ref_year - 1, 12)
                mom = pick(df, "주", met, "*TOTAL", my, f"{mm:02d}월 {wk}주차", "final")
            rows.append(f"- {met}: {fmt_value(met, cur)} "
                        f"(전주비 {fmt_delta(met, cur, prv) or '–'}, "
                        f"전월비 {fmt_delta(met, cur, mom) or '–'}, "
                        f"전년비 {fmt_delta(met, cur, yoy) or '–'})")
        else:
            cur = pick(df, "월", met, "*TOTAL", ref_year, month_label(ref_month), "mtd")
            prv = pick(df, "월", met, "*TOTAL", ref_year - 1, month_label(ref_month), "mtd")
            if isinstance(cur, float) and np.isnan(cur): continue
            rows.append(f"- {met}: {fmt_value(met, cur)} (전년비 {fmt_delta(met, cur, prv) or '–'})")
    return "\n".join(rows)

def ai_generate_insight(df, ref_year, ref_month, ref_week, model, focus="전주 주요 지표 현황"):
    """Claude API로 보고 인사이트 생성 → HTML(역신장 빨강) 반환. (텍스트, 에러) 튜플"""
    key = _anthropic_key()
    if not key:
        return None, ("ANTHROPIC_API_KEY가 설정되지 않았습니다. "
                      "Streamlit Cloud → Settings → Secrets에 ANTHROPIC_API_KEY를 추가하세요.")
    try:
        import anthropic
    except ImportError:
        return None, "anthropic 패키지가 설치되지 않았습니다. requirements.txt 반영 후 재배포하세요."

    period = f"{ref_year}년 {ref_week}" if ref_week else f"{ref_year}년 {ref_month}월"
    facts = _ai_metric_facts(df, ref_year, ref_month, ref_week)
    system = (
        "당신은 LF몰 CRM 첫구매 보고서를 작성하는 데이터 분석가입니다. "
        "주어진 지표 수치만 근거로 한국어 실무 보고 문구를 작성하세요. "
        "수치를 지어내지 말고, 추세·특이점·점검 포인트를 간결한 불릿으로 정리하세요. "
        "출력은 HTML로만 (불릿은 <br>로 구분). "
        "증감 수치 중 역신장(감소)은 <span style=\"color:#dc2626;font-weight:700\">…</span>, "
        "신장(증가)은 <span style=\"color:#16a34a;font-weight:700\">…</span>로 감싸세요. "
        "서론·맺음말 없이 불릿만 출력하세요."
    )
    user = f"[기준: {period}]\n다음 지표로 '{focus}' 문구를 작성하세요:\n{facts}"
    try:
        client = anthropic.Anthropic(api_key=key)
        resp = client.messages.create(
            model=model, max_tokens=2000, system=system,
            messages=[{"role": "user", "content": user}],
        )
        text = "".join(b.text for b in resp.content if b.type == "text").strip()
        return (text or None), (None if text else "빈 응답이 반환됐습니다.")
    except anthropic.AuthenticationError:
        return None, "API 키 인증에 실패했습니다. 키를 확인하세요."
    except anthropic.RateLimitError:
        return None, "요청이 많아 일시적으로 제한됐습니다. 잠시 후 다시 시도하세요."
    except Exception as e:
        return None, f"생성 중 오류: {e}"

# ══════════════════════════════════════════════════════
# 회원 실적 페이지
# ══════════════════════════════════════════════════════
def member_delta(metric, cur, prev):
    if cur is None or prev is None: return None
    if isinstance(cur, float) and np.isnan(cur): return None
    if isinstance(prev, float) and np.isnan(prev): return None
    if metric in MEMBER_PCT:
        d = (cur - prev) * 100
        return f"△{abs(d):.2f}%p" if d < 0 else f"+{d:.2f}%p"
    if prev == 0: return None
    d = (cur - prev) / prev * 100
    return f"△{abs(d):.1f}%" if d < 0 else f"+{d:.1f}%"

def render_member_page(mdf, chart_years=None):
    st.markdown("## 회원 실적 — 신규/기존 세그먼트")
    if mdf.empty:
        st.info("회원 실적 데이터가 없습니다. 회원 실적 엑셀을 업로드해주세요.")
        return
    years = sorted(mdf["year"].dropna().unique().astype(int))
    last_y = years[-1]
    last_mo = int(mdf[mdf["year"] == last_y]["sortkey"].max() % 100)
    st.caption(f"최신: {last_y}년 {last_mo}월 · {years[0]}–{years[-1]}년")

    # KPI 카드 — 유효회원수/구매고객수/거래액/CR (Total, 전월·전년비)
    pm_y, pm_m = (last_y, last_mo - 1) if last_mo > 1 else (last_y - 1, 12)
    kpis = ["유효회원수", "구매고객수", "거래액 (VAT제외)", "CR(%)"]
    cols = st.columns(4)
    for col, met in zip(cols, kpis):
        cur = member_pick(mdf, met, "Total", last_y, last_mo)
        mom = member_pick(mdf, met, "Total", pm_y, pm_m)
        yoy = member_pick(mdf, met, "Total", last_y - 1, last_mo)
        pills = ""
        for d, lab in [(member_delta(met, cur, mom), "전월비"),
                       (member_delta(met, cur, yoy), "전년비")]:
            if d:
                neg = d.startswith("△")
                pills += (f'<div class="kpi-delta {"down" if neg else "up"}">'
                          f'{"" if neg else "↑ "}{d} ({lab})</div>')
            else:
                pills += f'<div class="kpi-delta na">– ({lab})</div>'
        col.markdown(f'<div class="kpi-card"><div class="kpi-label">{met} ({last_mo}월)</div>'
                     f'<div class="kpi-value">{fmt_member(met, cur)}</div>{pills}</div>',
                     unsafe_allow_html=True)
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 신규/기존 구성비 (유효회원수 최신월)
    st.subheader(f"회원 구성 — {last_y}년 {last_mo}월 (유효회원수)")
    comp = {s: member_pick(mdf, "유효회원수", s, last_y, last_mo)
            for s in ["당월신규", "기가입신규", "기존"]}
    total = member_pick(mdf, "유효회원수", "Total", last_y, last_mo)
    c1, c2 = st.columns([1, 1])
    with c1:
        fig = go.Figure(go.Pie(
            labels=list(comp.keys()), values=[v for v in comp.values()],
            hole=0.55, marker=dict(colors=[clr("blue"), clr("teal"), clr("slate")]),
            textinfo="label+percent"))
        fig.update_layout(**{**base_layout(280, title="유효회원 구성비"), "showlegend": False})
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        rows = []
        for s in ["Total", "당월신규", "기가입신규", "기존"]:
            v = member_pick(mdf, "유효회원수", s, last_y, last_mo)
            share = (v / total * 100) if (total and not np.isnan(v)) else np.nan
            rows.append({"세그먼트": s, "유효회원수": fmt_member("유효회원수", v),
                         "비중": "–" if np.isnan(share) else f"{share:.1f}%"})
        st.dataframe(pd.DataFrame(rows).set_index("세그먼트"), use_container_width=True)

    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

    # 지표 선택 → 세그먼트별 YoY 추이
    met = st.selectbox("지표 선택", MEMBER_METRICS, key="wr_mem_metric")
    cyears = chart_years or years[-2:]
    cyears = [y for y in cyears if y in years] or years[-2:]
    div = 1e8 if met == "거래액 (VAT제외)" else 1
    unit = "억" if met == "거래액 (VAT제외)" else ("%" if met in MEMBER_PCT else "")
    if met in MEMBER_PCT: div = 0.01

    st.subheader(f"{met} — 세그먼트별 월 추이 ({last_y}년)")
    x_all = [f"{i}월" for i in range(1, 13)]
    fig = go.Figure()
    seg_color = {"Total": "slate", "당월신규": "blue", "기가입신규": "teal", "기존": "amber"}
    for s in MEMBER_SEGS:
        ser = member_series(mdf, met, s, last_y).reindex(x_all).dropna()
        if ser.empty: continue
        fig.add_trace(go.Scatter(x=ser.index.tolist(), y=(ser / div).tolist(),
                                 mode="lines+markers", name=s,
                                 line=dict(color=clr(seg_color[s]), width=2), marker=dict(size=5)))
    ly = base_layout(320, ysuffix=unit if unit == "%" else "", title=f"{met} ({unit})")
    ly["xaxis"]["categoryorder"] = "array"; ly["xaxis"]["categoryarray"] = x_all
    fig.update_layout(**ly)
    st.plotly_chart(fig, use_container_width=True)

    # Total 전년 비교
    st.subheader(f"{met} — Total 전년 비교")
    fig2 = go.Figure()
    for i, y in enumerate(sorted(cyears)):
        ser = member_series(mdf, met, "Total", y).reindex(x_all).dropna()
        if ser.empty: continue
        fig2.add_trace(go.Scatter(x=ser.index.tolist(), y=(ser / div).tolist(),
                                  mode="lines+markers", name=str(y),
                                  line=dict(color=clr(YEAR_PAL[i % len(YEAR_PAL)]), width=2),
                                  marker=dict(size=5)))
    ly2 = base_layout(320, ysuffix=unit if unit == "%" else "", title=f"{met} Total ({unit})")
    ly2["xaxis"]["categoryorder"] = "array"; ly2["xaxis"]["categoryarray"] = x_all
    fig2.update_layout(**ly2)
    st.plotly_chart(fig2, use_container_width=True)

    # 세그먼트 × 월 표 (최신 연도)
    st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
    st.subheader(f"{met} — 세그먼트 × 월 ({last_y}년)")
    rows = []
    for s in MEMBER_SEGS:
        ser = member_series(mdf, met, s, last_y)
        row = {"세그먼트": s}
        for lb in x_all:
            if lb in ser.index and not (isinstance(ser[lb], float) and np.isnan(ser[lb])):
                row[lb] = fmt_member(met, ser[lb])
        rows.append(row)
    st.dataframe(pd.DataFrame(rows).set_index("세그먼트"), use_container_width=True)

# ══════════════════════════════════════════════════════
# 페이지 PDF 저장 (브라우저 인쇄 → PDF, 차트 포함)
# ══════════════════════════════════════════════════════
def print_button(label="이 페이지 PDF 저장 / 인쇄"):
    components.html(
        f"""<button onclick="window.parent.print()"
        style="float:right;background:#2E68B0;color:#fff;border:0;border-radius:6px;
        padding:7px 14px;font-size:13px;font-weight:600;cursor:pointer;
        font-family:'Pretendard',-apple-system,sans-serif">{label}</button>
        <div style="clear:both"></div>""", height=44)
    st.caption("버튼이 동작하지 않으면 Ctrl+P(Mac ⌘+P) → 대상을 'PDF로 저장'으로 인쇄하세요.")

# ══════════════════════════════════════════════════════
# 메인 앱
# ══════════════════════════════════════════════════════
def main():
    if "wr_texts" not in st.session_state:
        st.session_state.wr_texts = load_insights()

    with st.sidebar:
        st.markdown("## 📋 주간보고 통합")
        files = st.file_uploader(
            "원천 엑셀/CSV/ZIP 업로드 (복수 선택)",
            type=["xlsx", "xls", "csv", "zip"], accept_multiple_files=True, key="wr_up",
            help="주간 폴더를 zip으로 묶어 통째로 올려도 됩니다. "
                 "전체관점 마스터(일/주/월) + 지표별 파일(가입율·가입자수·당일가입 첫구매율·비회원 트래픽)을 자동 인식합니다.")
        st.markdown("---")
        PAGES = ["01. 주간보고 요약", "02. 월별 추이", "03. 주차별 추이",
                 "04. 채널별 실적", "05. 회원 실적", "06. 통합 데이터·다운로드"]
        page = st.radio("페이지", PAGES, key="wr_page")

    stored = load_store()
    member_stored = load_member_store()
    expanded = expand_uploads(files) if files else []
    df_new = combine_files(tuple(expanded)) if expanded else pd.DataFrame()
    member_new = combine_member(tuple(expanded)) if expanded else pd.DataFrame()

    has_any = (not stored.empty or not member_stored.empty
               or not df_new.empty or not member_new.empty)
    if files and df_new.empty and member_new.empty and stored.empty and member_stored.empty:
        st.error("업로드한 파일에서 데이터를 읽지 못했습니다. 파일명 형식을 확인해주세요.")
        st.stop()
    if not has_any:
        st.info("👈 사이드바에서 주간 폴더의 엑셀/CSV/ZIP 파일들을 업로드해주세요.")
        st.markdown("""
- **마스터**: `전체관점 - 일자별/주별/월별 실적 (기본)`
- **지표별**: `월_가입율(일평균)`, `주_가입자수(일평균)`, `일_비회원 트래픽(일평균)`, `월_당일가입 첫구매율 (일평균)` …
- **회원 실적**: 신규/기존 회원별 9개 지표 월별 파일 (유효회원수·UV·구매고객수·CR 등)
- 주간 폴더를 **zip으로 묶어 통째로** 올려도 됩니다. 파일 내용으로 단위·지표를 자동 감지합니다.
- 업로드 후 **저장**을 누르면 누적됩니다. 다음에 기간이 다른 파일을 올리면 **겹치는 기간은 최신값으로 갱신**되고 나머지는 이어붙습니다.
""")
        st.stop()

    # 누적 저장소와 병합 — 미리보기(저장 전까지 영구 반영 안 함)
    df = merge_store(stored, df_new)
    mdf = merge_member(member_stored, member_new)

    has_new = (not df_new.empty) or (not member_new.empty)
    sig = tuple(sorted((n, len(b)) for n, b in expanded))
    with st.sidebar:
        if has_new:
            added, updated = upload_diff(stored, df_new, member_stored, member_new)
            saved = st.session_state.get("wr_saved_sig") == sig
            if saved:
                st.success("저장됨 ✓ (누적 반영 완료)")
            else:
                st.warning(f"새 데이터 감지 — 추가 {added}기간 · 갱신(겹침) {updated}기간\n\n"
                           "**저장** 눌러야 누적에 반영됩니다.")
                if st.button("💾 저장 (누적 반영)", key="wr_commit",
                             type="primary", use_container_width=True):
                    if not df_new.empty: save_store(df)
                    if not member_new.empty: save_member_store(mdf)
                    st.session_state["wr_saved_sig"] = sig
                    st.rerun()

    # 코어 데이터가 비어 회원 데이터만 있을 때: 회원 페이지로 안내
    if df.empty:
        st.warning("첫구매(전체관점/지표별) 데이터가 없습니다. **05. 회원 실적** 페이지를 이용하세요.")
        if not mdf.empty:
            print_button()
            render_member_page(mdf)
        st.stop()

    # ── 인식 결과 + 필터
    years_all = sorted(df["year"].dropna().unique().astype(int))
    ly, llabel = latest_period(df, "월")
    ref_year_default = ly or years_all[-1]
    with st.sidebar:
        st.markdown("---")
        src = f"인식된 파일 {len(expanded)}개" if expanded else "누적 데이터 사용 중"
        st.caption(f"{src} · 지표 {df['metric'].nunique()}종 · "
                   f"{years_all[0]}–{years_all[-1]}년")
        st.markdown("**기준 기간**")
        ref_year = st.selectbox("기준 연도", years_all[::-1],
                                index=years_all[::-1].index(ref_year_default), key="wr_refy")
        months_avail = sorted({int(re.match(r"(\d+)월", l).group(1))
                               for l in df[(df["gran"] == "월") & (df["year"] == ref_year)]["label"]})
        ref_month = st.selectbox("기준 월", months_avail[::-1], key="wr_refm")
        weeks_avail = (df[(df["gran"] == "주") & (df["year"] == ref_year) & df["value"].notna()]
                       [["label", "sortkey"]].drop_duplicates()
                       .sort_values("sortkey")["label"].tolist())
        if weeks_avail:
            # 최신 주차가 진행 중(일마감 데이터이거나 오늘이 속한 주차)이면 직전 주를 기본값으로
            latest_w = weeks_avail[-1]
            today = datetime.date.today()
            cur_week_lbl = f"{today.month:02d}월 {(today.day - 1) // 7 + 1}주차"
            is_partial = not df[(df["gran"] == "주") & (df["year"] == ref_year) &
                                (df["label"] == latest_w) & (df["close"] == "mtd")].empty
            in_progress = is_partial or (ref_year == today.year and latest_w == cur_week_lbl)
            default_week = (weeks_avail[-2] if in_progress and len(weeks_avail) >= 2
                            else latest_w)
            ref_week = st.selectbox("기준 주차", weeks_avail[::-1],
                                    index=weeks_avail[::-1].index(default_week),
                                    key="wr_refw",
                                    help="주간보고 대상 주차. 최신 주차가 진행 중이면 직전 완료 주차가 기본값입니다.")
        else:
            ref_week = None
        st.markdown("**차트 연도**")
        default_yrs = years_all[-2:] if len(years_all) >= 2 else years_all
        chart_years = st.multiselect("비교 연도", years_all, default=default_yrs, key="wr_cyrs")
        st.markdown("**채널**")
        ch_sel = st.multiselect("채널 선택", CHANNELS, default=CHANNELS, key="wr_ch")
        if not chart_years: chart_years = default_yrs

        st.markdown("---")
        st.markdown("**AI 인사이트 모델**")
        ai_label = st.selectbox("모델 선택", list(AI_MODELS.keys()), key="wr_ai_model_label")
        st.session_state["wr_ai_model"] = AI_MODELS[ai_label]
        st.caption("✅ API 키 설정됨" if _anthropic_key()
                   else "⚠ ANTHROPIC_API_KEY 미설정 — Secrets에 추가하세요")

        st.markdown("---")
        st.markdown("**누적 데이터**")
        saved_rows = len(load_store())
        pend = " · 저장 시 반영" if (has_new and not st.session_state.get("wr_saved_sig") == sig) else ""
        st.caption(f"저장됨 {saved_rows:,}행 / 현재 보기 {len(df):,}행{pend}")
        st.download_button("⬇ 누적 데이터 백업 (CSV)",
                           df[STORE_COLS].to_csv(index=False).encode("utf-8-sig"),
                           "wr_data_store.csv", "text/csv", use_container_width=True,
                           help="앱 재배포 시 누적 데이터가 초기화될 수 있으니 주기적으로 백업하세요. 이 CSV를 다시 업로드하면 복원됩니다.")
        if not mdf.empty:
            st.caption(f"회원 실적 {len(mdf):,}행")
            st.download_button("💾 회원 실적 백업 (CSV)",
                               mdf[MEMBER_COLS].to_csv(index=False).encode("utf-8-sig"),
                               "wr_member_store.csv", "text/csv", use_container_width=True)
        if st.button("🗑 누적 데이터 초기화", key="wr_clear_store", use_container_width=True):
            if os.path.exists(MEMBER_STORE): os.remove(MEMBER_STORE)
            if os.path.exists(DATA_STORE): os.remove(DATA_STORE)
            st.cache_data.clear()
            st.rerun()

        st.markdown("---")
        st.markdown("**보고란·메모**")
        st.download_button(
            "⬇ 보고란·메모 백업 (JSON)",
            json.dumps(st.session_state.wr_texts, ensure_ascii=False, indent=2).encode("utf-8"),
            "wr_insights.json", "application/json", use_container_width=True,
            help="모든 보고란·인사이트 메모를 백업합니다. 재배포로 초기화돼도 이 파일을 복원하면 되살아납니다.")
        restore = st.file_uploader("메모 복원 (JSON 업로드)", type=["json"], key="wr_restore_memo")
        if restore is not None:
            try:
                data = json.loads(restore.getvalue().decode("utf-8"))
                if isinstance(data, dict) and st.session_state.get("wr_restored") != restore.name:
                    merged = {**st.session_state.wr_texts, **data}  # 업로드 값 우선
                    st.session_state.wr_texts = merged
                    all_d = load_insights(); all_d.update(merged); save_insights(all_d)
                    st.session_state["wr_restored"] = restore.name
                    st.success(f"{len(data)}개 메모 복원됨 ✓"); st.rerun()
            except Exception as e:
                st.error(f"복원 실패: {e}")

    texts = st.session_state.wr_texts
    print_button()

    # ════════════ 01. 주간보고 요약 ════════════
    if page == "01. 주간보고 요약":
        st.markdown(f"## 첫구매 주간보고 — {ref_year}년 {ref_month}월")
        wy, wlabel = (ref_year, ref_week) if ref_week else latest_period(df, "주")
        if wlabel:
            st.caption(f"기준 주차: {week_disp(wy, wlabel)}")

        # KPI 카드 (기준 주차 — 전주비·전월비·전년비 모두 주차 기준)
        if wlabel:
            cols = st.columns(4)
            kpi_metrics = ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가", "가입자수"]
            py, plb = prev_label(df, "주", wy, wlabel)
            # 전월 동일 주차 (예: 06월 1주차 → 05월 1주차)
            mom_y = mom_lbl = None
            wm = re.match(r"(\d{1,2})월 (\d)주차", wlabel)
            if wm:
                mo, wk = int(wm.group(1)), int(wm.group(2))
                mom_y, mom_m = (wy, mo - 1) if mo > 1 else (wy - 1, 12)
                mom_lbl = f"{mom_m:02d}월 {wk}주차"
            for col, met in zip(cols, kpi_metrics):
                cur = pick(df, "주", met, "*TOTAL", wy, wlabel, "mtd")
                deltas = []
                prv = pick(df, "주", met, "*TOTAL", py, plb, "final") if plb else np.nan
                deltas.append((fmt_delta(met, cur, prv), "전주비"))
                if mom_lbl:
                    mom = pick(df, "주", met, "*TOTAL", mom_y, mom_lbl, "final")
                    deltas.append((fmt_delta(met, cur, mom), "전월비"))
                yoy = pick(df, "주", met, "*TOTAL", wy - 1, wlabel, "final")
                deltas.append((fmt_delta(met, cur, yoy), "전년비"))
                pills = ""
                for d, lab in deltas:
                    if d:
                        neg = d.startswith("△")
                        cls = "down" if neg else "up"
                        prefix = "" if neg else "↑ "
                        pills += f'<div class="kpi-delta {cls}">{prefix}{d} ({lab})</div>'
                    else:
                        pills += f'<div class="kpi-delta na">– ({lab})</div>'
                col.markdown(
                    f'<div class="kpi-card"><div class="kpi-label">{met} ({week_disp(wy, wlabel)})</div>'
                    f'<div class="kpi-value">{fmt_value(met, cur)}</div>{pills}</div>',
                    unsafe_allow_html=True)
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 — 전주비 (주차 기준)
        if wlabel:
            st.subheader("실적 요약 (일평균 · 전주비)")
            st.caption(f"기준 주차: {week_disp(wy, wlabel)} — 전주·전년 동주 대비")
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         use_container_width=True)
            st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 실적 요약 YoY 표
        st.subheader("실적 요약 (일평균 · 전년비)")
        tbl, (pm_y, pm_m) = yoy_summary_table(df, ref_year, ref_month, METRICS7)
        st.caption(f"전월({pm_m}월)은 월마감, 당월({ref_month}월)은 일마감(MTD) 기준 동일기간 비교")
        st.dataframe(style_delta_cols(tbl), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 보고란
        draft = auto_draft(df, ref_year, ref_month, ref_week=wlabel)
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        cL, cR = st.columns(2)
        with cL:
            report_text_block("wr_metrics_summary", "전주 주요 지표 현황",
                              default=draft, regen=draft,
                              ai_fn=lambda: ai_generate_insight(df, ref_year, ref_month,
                                                                wlabel, ai_model))
        with cR:
            report_text_block("wr_exec_summary", "금주 집행 내용 요약")

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)

        # 핵심 차트 (주차별 YoY 3종)
        st.subheader("주차별 추이 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "주", met, chart_years, h=280),
                                use_container_width=True)

    # ════════════ 02. 월별 추이 ════════════
    elif page == "02. 월별 추이":
        st.markdown("## 월별 추이")
        st.subheader("월별 추이 차트 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "월", met, chart_years, h=280),
                                use_container_width=True)
        c4, c5, c6 = st.columns(3)
        for col, met in zip((c4, c5, c6), ["비회원트래픽", "가입자수", "가입율"]):
            with col:
                st.plotly_chart(yoy_chart(df, "월", met, chart_years, h=280),
                                use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("월별 추이표 (일평균)")
        tbl = trend_table(df, "월", METRICS7, chart_years)
        st.dataframe(style_trend(tbl, METRICS7), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        report_text_block(
            f"wr_month_memo_{ref_year}_{ref_month}",
            f"{ref_year}년 {ref_month}월 액션·이슈사항",
            ai_fn=lambda: ai_generate_insight(df, ref_year, ref_month, None, ai_model,
                                              focus=f"{ref_year}년 {ref_month}월 액션·이슈 및 인사이트"))

    # ════════════ 03. 주차별 추이 ════════════
    elif page == "03. 주차별 추이":
        st.markdown("## 주차별 추이")
        st.subheader("주차별 추이 차트 — 전년 비교")
        c1, c2, c3 = st.columns(3)
        for col, met in zip((c1, c2, c3), ["첫구매 거래액", "첫구매 고객수", "첫구매 객단가"]):
            with col:
                st.plotly_chart(yoy_chart(df, "주", met, chart_years, h=280),
                                use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"주차별 추이표 — {ref_year}년")
        tbl = trend_table(df, "주", METRICS7, [ref_year])
        if not tbl.empty:
            recent = tbl.columns[-16:]
            st.dataframe(style_trend(tbl[recent], METRICS7), use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("전주비(WoW)·전년비(YoY) 증감")
        wy, wlabel = (ref_year, ref_week) if ref_week else latest_period(df, "주")
        if wlabel:
            st.caption(f"기준 주차: {week_disp(wy, wlabel)}")
            st.dataframe(style_delta_cols(wow_summary_table(df, wy, wlabel, METRICS7)),
                         use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        wy2, wlabel2 = (ref_year, ref_week) if ref_week else latest_period(df, "주")
        ai_model = st.session_state.get("wr_ai_model", "claude-sonnet-4-6")
        report_text_block(
            f"wr_week_memo_{wy2}_{wlabel2}",
            f"{wy2}년 {wlabel2} 액션·이슈사항" if wlabel2 else "주차별 액션·이슈사항",
            ai_fn=lambda: ai_generate_insight(df, ref_year, ref_month, wlabel2, ai_model,
                                              focus=f"{wlabel2} 주차 액션·이슈 및 인사이트"))

    # ════════════ 04. 채널별 실적 ════════════
    elif page == "04. 채널별 실적":
        st.markdown("## 채널(BPU)별 실적")
        avail = [m for m in METRICS7 if (df["metric"] == m).any()]
        met = st.selectbox("지표 선택", avail, key="wr_chmet")

        st.subheader(f"{met} — {ref_year}년 {ref_month}월 채널별 전년비")
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            pv = pick(df, "월", met, seg, ref_year - 1, month_label(ref_month), "mtd")
            cv = pick(df, "월", met, seg, ref_year, month_label(ref_month), "mtd")
            rows.append({"채널": seg,
                         f"{ref_year-1}년 {ref_month}월": fmt_value(met, pv),
                         f"{ref_year}년 {ref_month}월": fmt_value(met, cv),
                         "전년비": fmt_delta(met, cv, pv) or "–"})
        st.dataframe(style_delta_cols(pd.DataFrame(rows).set_index("채널")),
                     use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} — 채널별 월 추이 ({ref_year}년)")
        unit, div = METRIC_UNIT.get(met, ("", 1))
        if met in PCT_METRICS: div, unit = 0.01, "%"
        fig = go.Figure()
        x = [month_label(i) for i in range(1, 13)]
        for seg in [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "월", met, seg, ref_year).reindex(x).dropna()
            if s.empty: continue
            fig.add_trace(go.Scatter(
                x=s.index.tolist(), y=(s / div).tolist(), mode="lines+markers", name=seg,
                line=dict(color=clr(CHANNEL_PAL.get(seg, "blue")), width=1.8),
                marker=dict(size=4)))
        ly = base_layout(340, ysuffix=unit if unit == "%" else "",
                         title=f"{met} 채널별 ({unit})")
        fig.update_layout(**ly)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader(f"{met} — 채널 × 월 표 ({ref_year}년)")
        rows = []
        for seg in ["*TOTAL"] + [c for c in CHANNELS if c in ch_sel]:
            s = series_by_label(df, "월", met, seg, ref_year)
            row = {"채널": seg}
            for lb in [month_label(i) for i in range(1, 13)]:
                if lb in s.index and not np.isnan(s[lb]):
                    row[lb] = fmt_value(met, s[lb])
            rows.append(row)
        st.dataframe(pd.DataFrame(rows).set_index("채널"), use_container_width=True)

    # ════════════ 05. 회원 실적 ════════════
    elif page == "05. 회원 실적":
        render_member_page(mdf, chart_years)

    # ════════════ 06. 통합 데이터·다운로드 ════════════
    else:
        st.markdown("## 통합 데이터 · 다운로드")
        st.caption("업로드한 모든 파일을 합친 통합 long 데이터입니다.")
        st.dataframe(df.sort_values(["gran", "metric", "segment", "sortkey"]).head(2000),
                     use_container_width=True, height=420)

        st.markdown('<div class="sdiv"></div>', unsafe_allow_html=True)
        st.subheader("통합 워크북 다운로드")
        st.caption("`첫구매_요약`(요약표·보고란·YoY 차트·채널표) + `월`·`주`(통합 데이터) 3개 시트")
        if st.button("📥 엑셀 워크북 생성", key="wr_build"):
            with st.spinner("워크북 생성 중…"):
                xls = build_workbook(df, st.session_state.wr_texts,
                                     ref_year, ref_month, chart_years)
            st.download_button(
                "다운로드 — 첫구매_주간보고.xlsx", xls,
                file_name=f"첫구매_주간보고_{ref_year}{ref_month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("통합 long 데이터 CSV", csv, "통합데이터.csv", "text/csv")
        if not mdf.empty:
            st.download_button("회원 실적 데이터 CSV",
                               mdf.to_csv(index=False).encode("utf-8-sig"),
                               "회원실적데이터.csv", "text/csv")


if st.runtime.exists():
    main()
