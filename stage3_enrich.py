"""Stage 3 — 초근master 보강 (A~M열).

Stage 1이 만든 초근master(N~BN)에 왼쪽 A~M을 채운다. 실제 셀 수식 기준:
  · B 코드_변환(인사자료) = VLOOKUP(사번, 재직추출 A:AN, 40)   → 재직추출 AN(26년 조직코드)
  · C 근무부서명          = VLOOKUP(B, 26년조직 B:C, 2)
  · D 검증_변환=소속      = (B == R)   R=정보검색 부서코드  → 명부부서 vs 신청서부서 일치 검증
  · H 담당1코드 = VLOOKUP(B,…E)  I 담당1 = …D
  · J 담당2코드 = …G        K 담당2 = …F
  · L 담당3코드 = …I        M 담당3 = …H
(E·F·G는 소속≠근무 보조 주석열로, 샘플에서도 비어있어 그대로 비워둔다.)

예외 자동검출:
  ① 사번_미발견 : 초근 기록은 있으나 재직추출에 없는 사번
  ② 담당_매핑누락 : 조직코드가 26년조직 표에 없음
  ③ 근무≠부서   : D==False (명부부서 ≠ 신청서부서)
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import ROSTER_AN, ORG26_CODE, ORG26


@dataclass
class Stage3Result:
    master: pd.DataFrame
    matched_people: int = 0
    sabun_missing: pd.DataFrame = field(default_factory=pd.DataFrame)   # 예외①
    dam_unmapped: pd.DataFrame = field(default_factory=pd.DataFrame)    # 예외②
    dept_mismatch: pd.DataFrame = field(default_factory=pd.DataFrame)   # 예외③
    report: dict = field(default_factory=dict)


def enrich_chogeun_master(master: pd.DataFrame,
                          roster: pd.DataFrame,
                          org26: pd.DataFrame) -> Stage3Result:
    """master : Stage1 결과 (사번/부서코드/근무구분/조정계 등 채워짐, A~M 비어있음)
       roster : Stage2 결과 (사번, ROSTER_AN=26년 조직코드 보유)
       org26  : 26년조직_특수문자제외 (위치로 읽음)
    """
    m = master.copy()
    m["사번"] = m["사번"].astype(str).str.strip()

    # 사번 → 조직코드 (재직추출 AN)
    rs = roster.copy()
    rs["사번"] = rs["사번"].astype(str).str.strip()
    sabun2org = dict(zip(rs["사번"], rs[ROSTER_AN]))

    # 조직코드 → 26년조직 행 (담당계층/부서명)
    org = org26.copy()
    orgtbl = {}
    for _, r in org.iterrows():
        code = r.iloc[ORG26_CODE]
        if code is None or str(code).strip() == "":
            continue
        orgtbl[str(code).strip()] = r

    # ── B: 사번 → 조직코드 ──
    m["코드_변환(인사자료)"] = m["사번"].map(sabun2org)

    def org_get(code, field):
        if code is None or (isinstance(code, float) and pd.isna(code)):
            return None
        row = orgtbl.get(str(code).strip())
        if row is None:
            return None
        return row.iloc[ORG26[field]]

    B = m["코드_변환(인사자료)"]
    # ── C, 담당계층 ──
    m["근무부서명"] = B.map(lambda c: org_get(c, "근무부서명"))
    m["담당1코드"]  = B.map(lambda c: org_get(c, "담당1코드"))
    m["담당1"]      = B.map(lambda c: org_get(c, "담당1"))
    m["담당2코드"]  = B.map(lambda c: org_get(c, "담당2코드"))
    m["담당2"]      = B.map(lambda c: org_get(c, "담당2"))
    m["담당3코드"]  = B.map(lambda c: org_get(c, "담당3코드"))
    m["담당3"]      = B.map(lambda c: org_get(c, "담당3"))

    # ── D: B == R(부서코드) ──
    def eq(b, r):
        if b is None or (isinstance(b, float) and pd.isna(b)):
            return False
        return str(b).strip() == str(r).strip()
    m["검증_변환=소속"] = [eq(b, r) for b, r in zip(B, m["부서코드"])]

    # ── 예외 검출 ──
    miss_mask = B.isna()
    sabun_missing = (m[miss_mask].groupby(["사번", "성명"], dropna=False)
                     .size().reset_index(name="초근건수")) if miss_mask.any() \
        else pd.DataFrame(columns=["사번", "성명", "초근건수"])

    # B는 있는데 조직코드가 org 표에 없음
    unmapped_mask = (~B.isna()) & (~B.astype(str).str.strip().isin(orgtbl.keys()))
    dam_unmapped = (m[unmapped_mask].groupby(["코드_변환(인사자료)"], dropna=False)
                    .agg(초근건수=("사번", "count"),
                         사번목록=("사번", lambda s: ", ".join(s.astype(str)[:5])))
                    .reset_index().rename(columns={"코드_변환(인사자료)": "조직코드"})) \
        if unmapped_mask.any() else pd.DataFrame(columns=["조직코드", "초근건수", "사번목록"])

    mismatch_mask = (~B.isna()) & (~pd.Series(m["검증_변환=소속"].values, index=m.index))
    dept_mismatch = m[mismatch_mask][["사번", "성명", "코드_변환(인사자료)", "부서코드", "부서명"]] \
        .rename(columns={"코드_변환(인사자료)": "명부_조직코드", "부서코드": "신청서_부서코드"})

    report = {
        "초근행": len(m),
        "사번매칭_건수": int((~B.isna()).sum()),
        "사번_미발견_명": len(sabun_missing),
        "담당매핑누락_조직수": len(dam_unmapped),
        "근무≠부서_건수": int(mismatch_mask.sum()),
        "초근master_고유인원": int(m.loc[~B.isna(), "사번"].nunique()),
    }

    return Stage3Result(
        master=m,
        matched_people=report["초근master_고유인원"],
        sabun_missing=sabun_missing.reset_index(drop=True),
        dam_unmapped=dam_unmapped.reset_index(drop=True),
        dept_mismatch=dept_mismatch.reset_index(drop=True),
        report=report,
    )
