"""CEP(Category Entry Points) 키워드 수요조사 빌더.

왜? 카테고리 키워드(명사)와 별개로, 소비자가 '상황·맥락'으로 검색하는
CEP 키워드(결혼식 하객룩, 여름 휴가룩 등)의 수요를 측정해
[CEP]×[카테고리] 롱테일 pSEO 페이지 우선순위를 정하기 위함.

입력:
  .cep_keywords.json        축별 CEP 키워드 (이벤트/시즌/휴가/선물/활동/체형/연령/스타일)
  .cep_semrush_vol.csv      Semrush(구글) 검색량  (Keyword;Search Volume)
  data/naver_cep_metrics.csv  네이버 검색량(있으면 병합, 워크플로가 생성)

출력:
  data/cep_keyword_research.csv
    키워드, 검색량(구글), 네이버검색량, 대표검색량(네이버 우선), 추이,
    섹션(=CEP축), 순위, 우선순위, Status
"""
import csv
import json
import os
import re
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))


def norm(s):
    return re.sub(r"\s+", "", str(s)).lower()


def main():
    cep = json.load(open(os.path.join(ROOT, ".cep_keywords.json"), encoding="utf-8"))

    # Semrush(구글) 볼륨 — 띄어쓰기/붙임 모두 norm 키로 매칭
    vol = {}
    sp = os.path.join(ROOT, ".cep_semrush_vol.csv")
    if os.path.exists(sp):
        with open(sp, encoding="utf-8") as f:
            for row in csv.reader(f, delimiter=";"):
                if row and row[0] != "Keyword":
                    try:
                        vol[norm(row[0])] = int(row[1])
                    except (ValueError, IndexError):
                        pass

    # 네이버 볼륨(있으면)
    naver = {}
    npath = os.path.join(ROOT, "data", "naver_cep_metrics.csv")
    if os.path.exists(npath):
        with open(npath, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                try:
                    nv = int(r.get("네이버검색량") or 0)
                except ValueError:
                    nv = 0
                naver[norm(r["키워드"])] = (nv, r.get("추이", ""))
        print(f"네이버 CEP 데이터 병합: {npath} ({len(naver)}개)")

    rows = []
    by_axis = defaultdict(list)
    for axis, kws in cep.items():
        for kw in kws:
            msv = vol.get(norm(kw), 0)
            nv, ntrend = naver.get(norm(kw), (0, ""))
            rep = nv if nv > 0 else msv          # 대표검색량 = 네이버 우선
            r = {
                "키워드": kw, "검색량": msv, "네이버검색량": nv,
                "대표검색량": rep, "추이": ntrend, "섹션": axis,
                "Status": "신규" if rep > 0 else "미수집",
            }
            rows.append(r)
            by_axis[axis].append(r)

    # 축 내부에서 대표검색량 내림차순 순위
    for axis, items in by_axis.items():
        for i, r in enumerate(sorted(items, key=lambda x: x["대표검색량"], reverse=True), 1):
            r["순위"] = i
            r["우선순위"] = f"{i}순위"

    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    out = os.path.join(ROOT, "data", "cep_keyword_research.csv")
    cols = ["키워드", "검색량", "네이버검색량", "대표검색량", "추이",
            "섹션", "Status", "순위", "우선순위"]
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    g = sum(1 for r in rows if r["검색량"] > 0)
    n = sum(1 for r in rows if r["네이버검색량"] > 0)
    print(f"총 {len(rows)}개 CEP 키워드 → {out}")
    print(f"구글 보유 {g}개 · 네이버 보유 {n}개 · 대표검색량 보유 "
          f"{sum(1 for r in rows if r['대표검색량'] > 0)}개")
    print("축별 1순위:")
    for axis in cep:
        top = next((r for r in rows if r["섹션"] == axis and r["순위"] == 1), None)
        if top:
            print(f"  {axis} 1순위: {top['키워드']}({top['대표검색량']:,})")


if __name__ == "__main__":
    main()
