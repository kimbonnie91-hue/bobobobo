"""네이버 키워드 지표 수집기 — 검색광고 API + 데이터랩 + 쇼핑인사이트

왜? Semrush는 구글 기준이라 한국 실수요(네이버)를 못 잡는다.
이 스크립트가 우리 920개 키워드에 네이버 월간 검색량/추이/쇼핑수요를 붙인다.

────────────────────────────────────────────────────────────────
키 발급 (개인계정 가능)
────────────────────────────────────────────────────────────────
A) 검색광고 API  ─ https://searchad.naver.com  →  로그인 → 도구 → API 사용 관리
   · NAVER_AD_API_KEY      (액세스라이선스)
   · NAVER_AD_SECRET       (비밀키)
   · NAVER_AD_CUSTOMER_ID  (내 정보 → CUSTOMER_ID / 광고주 ID, 숫자)
B) 데이터랩/쇼핑인사이트 ─ https://developers.naver.com → Application → 애플리케이션 등록
   · NAVER_CLIENT_ID
   · NAVER_CLIENT_SECRET
   (사용 API에 '데이터랩(검색어트렌드)'·'데이터랩(쇼핑인사이트)' 체크)

환경변수로 넣거나 프로젝트 루트 .env 파일에 작성(.env 는 git 무시됨):
   NAVER_AD_API_KEY=...
   NAVER_AD_SECRET=...
   NAVER_AD_CUSTOMER_ID=...
   NAVER_CLIENT_ID=...
   NAVER_CLIENT_SECRET=...

실행:
   python naver_keyword_data.py            # 전체(검색량+추이)
   python naver_keyword_data.py --limit 30 # 테스트로 30개만

출력: data/naver_keyword_metrics.csv  (키워드, 네이버검색량, PC, 모바일, 경쟁정도, 추이, 추이라벨)
이후 `python build_keyword_data.py` 가 이 파일을 자동 병합한다.
"""

import argparse
import base64
import csv
import hashlib
import hmac
import json
import os
import re
import sys
import time
from datetime import date, timedelta

try:
    import requests
except ImportError:
    sys.exit("requests 가 필요합니다:  pip install requests")

ROOT = os.path.dirname(os.path.abspath(__file__))


# ── .env 간단 로더 (python-dotenv 없이) ──
def load_env():
    path = os.path.join(ROOT, ".env")
    if os.path.exists(path):
        for line in open(path, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())


def norm(s):
    return re.sub(r"\s+", "", str(s)).lower()


# ══════════════════════════════════════════════════════
# A) 검색광고 API — 키워드도구 (월간 검색수 절대값)
# ══════════════════════════════════════════════════════
AD_BASE = "https://api.searchad.naver.com"


def _ad_headers(method, path):
    api_key = os.environ["NAVER_AD_API_KEY"].strip()
    secret = os.environ["NAVER_AD_SECRET"].strip()
    customer = os.environ["NAVER_AD_CUSTOMER_ID"].strip()
    ts = str(round(time.time() * 1000))
    msg = f"{ts}.{method}.{path}"
    sig = base64.b64encode(
        hmac.new(secret.encode(), msg.encode(), hashlib.sha256).digest()
    ).decode()
    return {
        "X-Timestamp": ts, "X-API-KEY": api_key,
        "X-Customer": str(customer), "X-Signature": sig,
    }


def _to_int(v):
    """monthlyPcQcCnt 등은 정수 또는 '< 10' 문자열로 옴."""
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).replace(",", "").strip()
    if s.startswith("<"):
        return 5            # '< 10' → 보수적으로 5
    try:
        return int(s)
    except ValueError:
        return 0


def fetch_ad_volumes(keywords):
    """검색광고 키워드도구로 월간 검색량 조회. {norm(kw): {...}} 반환."""
    path = "/keywordstool"
    out = {}
    batches = [keywords[i:i + 5] for i in range(0, len(keywords), 5)]
    for bi, batch in enumerate(batches, 1):
        # 힌트키워드는 공백 제거 권장
        hint = ",".join(k.replace(" ", "") for k in batch)
        for attempt in range(4):
            try:
                r = requests.get(AD_BASE + path,
                                 params={"hintKeywords": hint, "showDetail": "1"},
                                 headers=_ad_headers("GET", path), timeout=15)
                if r.status_code == 429:           # rate limit
                    time.sleep(2 ** attempt)
                    continue
                r.raise_for_status()
                for row in r.json().get("keywordList", []):
                    pc = _to_int(row.get("monthlyPcQcCnt", 0))
                    mo = _to_int(row.get("monthlyMobileQcCnt", 0))
                    out[norm(row.get("relKeyword", ""))] = {
                        "pc": pc, "mobile": mo, "total": pc + mo,
                        "comp": row.get("compIdx", ""),
                    }
                break
            except requests.RequestException as e:
                if attempt == 3:
                    detail = ""
                    if hasattr(e, "response") and e.response is not None:
                        detail = f" body={e.response.text[:200]}"
                    print(f"  [경고] 배치 {bi} 실패: {e}{detail}")
                time.sleep(2 ** attempt)
        if bi == 1 and not out:
            print("  ⚠ 첫 배치부터 실패 — API 키 확인 필요")
        if bi % 10 == 0:
            print(f"  검색광고 {bi}/{len(batches)} 배치…")
        time.sleep(0.3)                            # QPS 보호
    return out


# ══════════════════════════════════════════════════════
# B) 데이터랩 검색어트렌드 (최근 12개월 추이)
# ══════════════════════════════════════════════════════
DATALAB_URL = "https://openapi.naver.com/v1/datalab/search"


def _datalab_headers():
    return {
        "X-Naver-Client-Id": os.environ["NAVER_CLIENT_ID"].strip(),
        "X-Naver-Client-Secret": os.environ["NAVER_CLIENT_SECRET"].strip(),
        "Content-Type": "application/json",
    }


def fetch_datalab_trends(keywords):
    """최근 12개월 월별 상대추이 → {norm(kw): (최근지수, 성장%)}.
    그룹 5개/요청. 상승=초기3개월 대비 최근3개월 증가율."""
    end = date.today().replace(day=1) - timedelta(days=1)
    start = (end.replace(day=1) - timedelta(days=365)).replace(day=1)
    out = {}
    groups = [keywords[i:i + 5] for i in range(0, len(keywords), 5)]
    for gi, g in enumerate(groups, 1):
        body = {
            "startDate": start.isoformat(), "endDate": end.isoformat(),
            "timeUnit": "month",
            "keywordGroups": [{"groupName": k, "keywords": [k]} for k in g],
        }
        for attempt in range(4):
            try:
                r = requests.post(DATALAB_URL, data=json.dumps(body),
                                  headers=_datalab_headers(), timeout=15)
                if r.status_code == 429:
                    time.sleep(2 ** attempt)
                    continue
                r.raise_for_status()
                for res in r.json().get("results", []):
                    pts = [p["ratio"] for p in res.get("data", [])]
                    if not pts:
                        continue
                    early = sum(pts[:3]) / max(1, len(pts[:3]))
                    late = sum(pts[-3:]) / max(1, len(pts[-3:]))
                    growth = round((late - early) / early * 100) if early else 0
                    out[norm(res["title"])] = (round(pts[-1], 1), growth,
                                               [round(p, 1) for p in pts])
                break
            except requests.RequestException as e:
                if attempt == 3:
                    detail = ""
                    if hasattr(e, "response") and e.response is not None:
                        detail = f" body={e.response.text[:200]}"
                    print(f"  [경고] 데이터랩 그룹 {gi} 실패: {e}{detail}")
                time.sleep(2 ** attempt)
        if gi % 10 == 0:
            print(f"  데이터랩 {gi}/{len(groups)} 그룹…")
        time.sleep(0.3)
    return out


def trend_label(growth):
    if growth is None:
        return ""
    if growth >= 20:
        return "급상승"
    if growth >= 5:
        return "상승"
    if growth <= -20:
        return "급하락"
    if growth <= -5:
        return "하락"
    return "유지"


# ══════════════════════════════════════════════════════
# 메인
# ══════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="테스트용 상위 N개만")
    ap.add_argument("--source", choices=["category", "cep", "combo"], default="category",
                    help="category=.semrush_keywords_uniq.json / cep=.cep_keywords.json / "
                         "combo=data/combo_candidates.csv(조합키워드)")
    args = ap.parse_args()
    load_env()

    if args.source == "cep":
        cep = json.load(open(os.path.join(ROOT, ".cep_keywords.json"), encoding="utf-8"))
        keywords = [k for v in cep.values() for k in v]
        out_name = "naver_cep_metrics.csv"
        next_build = "python build_cep_data.py"
    elif args.source == "combo":
        # 조합 후보 CSV의 '조합키워드' 컬럼 (조합점수 순으로 이미 정렬돼 있음).
        # combo_candidates.csv는 생성물이라 git에 없을 수 있음 → 없으면 즉석 빌드.
        path = os.path.join(ROOT, "data", "combo_candidates.csv")
        if not os.path.exists(path):
            print("combo_candidates.csv 없음 → build_combos로 생성")
            import build_combos
            build_combos.main()
        with open(path, encoding="utf-8-sig") as f:
            keywords = [r["조합키워드"] for r in csv.DictReader(f) if r.get("조합키워드")]
        out_name = "naver_combo_metrics.csv"
        next_build = "python build_combos.py"
    else:
        keywords = json.load(open(os.path.join(ROOT, ".semrush_keywords_uniq.json"),
                                  encoding="utf-8"))
        out_name = "naver_keyword_metrics.csv"
        next_build = "python build_keyword_data.py"
    if args.limit:
        keywords = keywords[:args.limit]
    print(f"[{args.source}] 대상 키워드 {len(keywords)}개")

    have_ad = all(os.environ.get(k) for k in
                  ("NAVER_AD_API_KEY", "NAVER_AD_SECRET", "NAVER_AD_CUSTOMER_ID"))
    have_lab = all(os.environ.get(k) for k in
                   ("NAVER_CLIENT_ID", "NAVER_CLIENT_SECRET"))
    if not have_ad:
        sys.exit("검색광고 키(NAVER_AD_API_KEY/SECRET/CUSTOMER_ID)가 없습니다. "
                 "README 참고해 .env에 넣어주세요.")

    print("① 검색광고 API 월간 검색량 수집…")
    vols = fetch_ad_volumes(keywords)
    trends = {}
    if have_lab:
        print("② 데이터랩 12개월 추이 수집…")
        trends = fetch_datalab_trends(keywords)
    else:
        print("② 데이터랩 키 없음 → 추이 생략(검색량만 수집)")

    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    out = os.path.join(ROOT, "data", out_name)
    cols = ["키워드", "네이버검색량", "PC", "모바일", "경쟁정도", "추이지수", "추이", "월별추이"]
    n_hit = 0
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for kw in keywords:
            v = vols.get(norm(kw))
            t = trends.get(norm(kw))
            if v:
                n_hit += 1
            idx, growth, series = (t if t else (None, None, []))
            w.writerow({
                "키워드": kw,
                "네이버검색량": v["total"] if v else 0,
                "PC": v["pc"] if v else 0,
                "모바일": v["mobile"] if v else 0,
                "경쟁정도": v["comp"] if v else "",
                "추이지수": idx if idx is not None else "",
                "추이": trend_label(growth),
                "월별추이": "|".join(str(x) for x in series),   # 12개월 상대지수
            })
    print(f"완료 → {out}  (검색량 매칭 {n_hit}/{len(keywords)})")
    print(f"다음: {next_build}  (네이버 데이터 자동 병합)")


if __name__ == "__main__":
    main()
