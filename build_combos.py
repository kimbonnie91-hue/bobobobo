"""③단계 — [CEP]×[카테고리] 롱테일 조합 후보 생성.

왜? CEP(상황·맥락)와 카테고리(상품)를 곱해 롱테일 키워드를 만든다.
   예: [결혼식 하객] × [원피스] → "결혼식 하객 원피스"
이 후보가 ④단계(롱테일 수요조사)의 입력이 된다.

전략: 각 CEP 축은 '의미가 통하는' 카테고리 섹션하고만 조합한다.
   (선물 축은 가방·주얼리와, 골프 축은 골프 카테고리와…)
   고검색량 우선: CEP 상위 × 카테고리 상위만 뽑아 폭발(n×m)을 억제.

입력:  data/cep_keyword_research.csv, data/lfmall_keyword_research.csv
출력:  data/combo_candidates.csv
   CEP, CEP축, CEP검색량, 카테고리, 섹션, 카테고리검색량, 조합키워드, 조합점수
"""
import csv
import os
import re

import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))

# CEP 축 → 조합할 카테고리 섹션 (착장 의미에 맞게 정밀화)
#   '코디/착장' 성격 축(시즌·스타일·체형)은 의류·신발만 → 시계·양산 등 노이즈 제거
#   이벤트/연령은 가방까지, 선물은 잡화(주얼리·가방·지갑·뷰티), 여행은 수영/비치, 활동은 골프
AXIS_TO_SECTIONS = {
    "이벤트/TPO": ["의류", "신발", "가방"],          # 하객·면접·소개팅 = 착장
    "시즌/날씨":  ["의류", "신발"],                  # 계절 코디
    "휴가/여행":  ["의류", "신발", "가방", "수영/비치"],  # 여행 = 옷·캐리어·수영복
    "선물/관계":  ["주얼리/시계", "가방", "지갑/벨트", "뷰티/향수"],  # 선물 = 잡화
    "활동/스포츠": ["골프", "의류", "신발"],
    "체형/핏":    ["의류", "신발"],                  # 핏 = 옷
    "연령/세대":  ["의류", "신발", "가방"],
    "스타일/무드": ["의류", "신발"],                 # 스타일 = 코디
}

# CEP 키워드 → '수식어 핵심'만 추출 (접미 제거). 카테고리를 붙여 자연스러운 검색어로.
CEP_SUFFIX = ["룩", "코디", "패션", "옷차림", "옷", "웨어", "복장", "정장",
              "원피스", "선물", "신발", "복"]


def cep_core(kw):
    s = kw.strip()
    for suf in CEP_SUFFIX:
        if s.endswith(suf) and len(s) > len(suf):
            s = s[:-len(suf)].strip()
            break
    return s


TOP_CEP = 8     # 축별 상위 CEP
TOP_CAT = 12    # 섹션별 상위 카테고리


def main():
    cep = pd.read_csv(os.path.join(ROOT, "data", "cep_keyword_research.csv"),
                      encoding="utf-8-sig")
    cat = pd.read_csv(os.path.join(ROOT, "data", "lfmall_keyword_research.csv"),
                      encoding="utf-8-sig")
    for df in (cep, cat):
        df["대표검색량"] = pd.to_numeric(df["대표검색량"], errors="coerce").fillna(0).astype(int)

    rows = []
    for axis, sections in AXIS_TO_SECTIONS.items():
        top_cep = (cep[(cep["섹션"] == axis) & (cep["대표검색량"] > 0)]
                   .nlargest(TOP_CEP, "대표검색량"))
        # 매핑 섹션의 상위 카테고리 키워드 풀 (섹션별 상위 N)
        cat_pool = cat[(cat["섹션"].isin(sections)) & (cat["대표검색량"] > 0)]
        cat_top = pd.concat(
            [g.nlargest(TOP_CAT, "대표검색량") for _, g in cat_pool.groupby("섹션")]
        ) if len(cat_pool) else cat_pool
        for _, c in top_cep.iterrows():
            core = cep_core(c["키워드"])
            core_tokens = core.split()
            for _, k in cat_top.iterrows():
                cat_tokens = str(k["키워드"]).split()
                # 흡수 제거: CEP 핵심어가 카테고리에 이미 다 들어있으면 조합=카테고리 → 스킵
                if set(core_tokens) <= set(cat_tokens):
                    continue
                # 토큰 단위 중복 제거 (여름 + 여름샌들 류 방지)
                words = []
                for w in core_tokens + cat_tokens:
                    if w not in words:
                        words.append(w)
                combo = re.sub(r"\s+", " ", " ".join(words)).strip()
                # 조합점수 = √(CEP검색량) × √(카테고리검색량) — 둘 다 큰 쌍 우대
                score = round((c["대표검색량"] ** 0.5) * (k["대표검색량"] ** 0.5))
                rows.append({
                    "CEP": c["키워드"], "CEP축": axis, "CEP검색량": c["대표검색량"],
                    "카테고리": k["키워드"], "섹션": k["섹션"],
                    "카테고리검색량": k["대표검색량"],
                    "조합키워드": combo, "조합점수": score,
                })

    out_df = pd.DataFrame(rows).drop_duplicates("조합키워드")

    # ④단계 실제 조회분 병합 (.combo_vol.csv: 키워드;검색량 / .combo_kd.csv: 키워드;KD)
    def _load(path, conv):
        d = {}
        p = os.path.join(ROOT, path)
        if os.path.exists(p):
            for r in csv.reader(open(p, encoding="utf-8"), delimiter=";"):
                if len(r) >= 2 and r[0] not in ("키워드", "Keyword"):
                    try:
                        d[r[0]] = conv(r[1])
                    except ValueError:
                        pass
        return d

    real = _load(".combo_vol.csv", int)        # 구글(Semrush) 실제검색량
    kd = _load(".combo_kd.csv", float)         # 난이도(KD)

    # 네이버 조합 검색량 + 경쟁정도 — 별도 배치(naver-combo 워크플로)가 생성
    #   경쟁정도(높음/중간/낮음)는 네이버 검색광고 API가 함께 주는 KD 대용 지표.
    #   Semrush KD가 구글 기준이라 한국 시장에선 이쪽이 더 정확하다.
    naver, naver_comp = {}, {}
    np_ = os.path.join(ROOT, "data", "naver_combo_metrics.csv")
    if os.path.exists(np_):
        for r in csv.DictReader(open(np_, encoding="utf-8-sig")):
            try:
                naver[r["키워드"]] = int(r.get("네이버검색량") or 0)
            except ValueError:
                pass
            naver_comp[r["키워드"]] = r.get("경쟁정도", "")
        print(f"네이버 조합 검색량 병합: {np_} ({len(naver)}개)")

    out_df["실제검색량"] = out_df["조합키워드"].map(real)       # 구글
    out_df["네이버검색량"] = out_df["조합키워드"].map(naver)     # 네이버
    out_df["네이버경쟁"] = out_df["조합키워드"].map(naver_comp)  # KD 대용(높음/중간/낮음)
    out_df["KD"] = out_df["조합키워드"].map(kd)
    # 대표실제검색량 = 네이버 우선(없으면 구글) — CEP·카테고리와 동일 규칙
    out_df["대표실제검색량"] = out_df["조합키워드"].map(
        lambda k: naver[k] if naver.get(k, 0) > 0 else real.get(k))

    def _ver(k):
        g, n = real.get(k, 0) or 0, naver.get(k, 0) or 0
        if g > 0 or n > 0:
            return "검증됨"
        if k in real or k in naver:
            return "수요없음"
        return "미조회"
    out_df["검증"] = out_df["조합키워드"].apply(_ver)

    # ── 키워드 난이도(경쟁) 산출식 ────────────────────────────────────────
    # Semrush KD(구글 0~100)는 API 유닛 필요 + 구글 기준이라 한국에 부정확.
    # → 네이버 검색광고가 함께 주는 '경쟁정도'(높음/중간/낮음)를 난이도로 사용.
    # 이를 가중치로 환산 (낮을수록 노리기 쉬움 = 가중치 큼):
    #     경쟁가중 = {낮음 1.0, 중간 0.6, 높음 0.3}, 정보없음 0.5(중립)
    # 최종 공략점수 = 대표검색량 × 경쟁가중
    #     → 수요(검색량) 크고 난이도(경쟁) 낮은 키워드일수록 높음 = pSEO 우선순위.
    COMP_W = {"낮음": 1.0, "중간": 0.6, "높음": 0.3}
    rep = pd.to_numeric(out_df["대표실제검색량"], errors="coerce").fillna(0)
    w = out_df["네이버경쟁"].map(COMP_W).fillna(0.5)
    out_df["공략점수"] = (rep * w).round().astype(int)

    # ── 공략 그룹 3분할 (검증된 조합 대상) ────────────────────────────────
    # 기준 = 수요(대표검색량) × 난이도(경쟁). pSEO 페이지 제작 우선순위 티어.
    #   1군 즉시공략 : 검색량 ≥ 500  &  경쟁 낮음/중간
    #                  → 수요 충분 + 경쟁 낮음 = ROI 최고. 최우선 제작.
    #   2군 성장공략 : (검색량 ≥ 500 & 경쟁 높음)  또는  (100 ≤ 검색량 < 500)
    #                  → 수요 크나 경쟁 세거나, 수요 중간. 콘텐츠 보강 시 가치.
    #   3군 롱테일풀 : 검색량 < 100
    #                  → 개별 수요는 작지만 수가 많음. 템플릿 대량생성(AirOps) 후보.
    def _group(r):
        if r["검증"] != "검증됨":
            return ""
        sv = r["대표실제검색량"] if pd.notna(r["대표실제검색량"]) else 0
        comp = r["네이버경쟁"]
        if sv >= 500 and comp in ("낮음", "중간"):
            return "1군 즉시공략"
        if sv >= 100:
            return "2군 성장공략"
        return "3군 롱테일풀"
    out_df["공략그룹"] = out_df.apply(_group, axis=1)

    # 검증된 건 공략점수 순, 나머지는 조합점수(예측) 순
    out_df["_rep"] = pd.to_numeric(out_df["대표실제검색량"], errors="coerce").fillna(-1)
    out_df["_atk"] = out_df["공략점수"].where(out_df["검증"] == "검증됨", -1)
    out_df = out_df.sort_values(["_atk", "_rep", "조합점수"], ascending=False)
    out_df = out_df.drop(columns=["_rep", "_atk"]).reset_index(drop=True)
    out_df["우선순위"] = out_df.index + 1

    out = os.path.join(ROOT, "data", "combo_candidates.csv")
    out_df.to_csv(out, index=False, encoding="utf-8-sig")
    nver = int((out_df["검증"] == "검증됨").sum())
    ng = int((out_df["실제검색량"].fillna(0) > 0).sum())
    nn = int((out_df["네이버검색량"].fillna(0) > 0).sum())
    print(f"조합 후보 {len(out_df):,}개 → {out}  (검증됨 {nver}개 · 구글 {ng} · 네이버 {nn})")
    print(f"CEP축 {len(AXIS_TO_SECTIONS)}개 기준 · 축별 상위 {TOP_CEP} × 섹션별 상위 {TOP_CAT}")
    grp = out_df[out_df["공략그룹"] != ""]["공략그룹"].value_counts()
    for g in ["1군 즉시공략", "2군 성장공략", "3군 롱테일풀"]:
        print(f"  {g}: {int(grp.get(g, 0))}개")
    print("\n=== 조합점수 TOP 15 (④단계 우선 조사 대상) ===")
    for _, r in out_df.head(15).iterrows():
        print(f"  {r['우선순위']:>3}. {r['조합키워드']}  "
              f"(CEP {r['CEP']} {r['CEP검색량']:,} × {r['카테고리']} {r['카테고리검색량']:,})")


if __name__ == "__main__":
    main()
