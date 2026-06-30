"""Stage 2 — 인사자료추출 → 재직추출.

매뉴얼 나·라 단계:
  · 나-1 : 사번(A) 오름차순 후 '9천번대~끝' 삭제  →  사번이 '9'로 시작하는 행 삭제
  · 나-2 : 현원여부(C) == 'N' 행 삭제
  · 라   : AN '근무부서' = VLOOKUP(U=근무부서코드, 연중부서코드매핑 B:D, 3)  (=26년 조직코드)
  · 확인 : AO = (U == S)  →  FALSE인 행 = 소속≠근무 (참고)소속근무다름 후보)

사람이 판단해야 하는 두 가지를 '예외 목록'으로 자동 추출한다:
  ① 매핑 안 된 근무부서(코드)  → 연중부서코드매핑에 추가해야 함
  ② 소속≠근무 인원              → 어느 쪽으로 집계할지 확인 필요
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import (
    ROSTER_HEADERS, ROSTER_AN, ROSTER_AO,
    YEONJUNG_KEY, YEONJUNG_VAL, YEONJUNG_NAME,
)


@dataclass
class Stage2Result:
    roster: pd.DataFrame                       # 필터·매핑 완료된 재직추출 (A~AO)
    headcount: int = 0                         # 집계 기준 재직 인원수
    unmapped_depts: pd.DataFrame = field(default_factory=pd.DataFrame)  # 예외①
    mismatch_rows: pd.DataFrame = field(default_factory=pd.DataFrame)   # 예외②
    report: dict = field(default_factory=dict)


def build_jaejik(insa: pd.DataFrame, yeonjung: pd.DataFrame) -> Stage2Result:
    """insa     : 인사자료추출 (A~AM, 39열). 컬럼명 무시·위치로 읽음. 사번/코드는 문자열.
       yeonjung : 연중부서코드매핑 (A~E). 위치로 읽음.
    """
    df = insa.copy()
    df.columns = ROSTER_HEADERS[: df.shape[1]]
    for c in ["사번", "소속부서", "근무부서"]:
        df[c] = df[c].astype(str).str.strip()

    n0 = len(df)
    report = {"입력행": n0}

    # ── 나-1: 집계 대상은 '숫자로 시작하는 사번'만.
    #   · 알파벳 시작(임원·지원선수·임원기사·외부 등) 삭제
    #   · 9천번대(9 시작)도 제외  (9는 숫자라 위 규칙에 안 걸리므로 별도 처리)
    first = df["사번"].str[:1]
    mask_nondigit = ~first.str.isdigit().fillna(False)
    mask9 = first.eq("9")
    report["사번_비숫자시작_삭제"] = int(mask_nondigit.sum())
    report["사번9시작_삭제"] = int(mask9.sum())
    df = df[~(mask_nondigit | mask9)]

    # ── 나-2: 현원여부 'N' 삭제 ──
    maskN = df["현원여부"].astype(str).str.strip().eq("N")
    report["현원N_삭제"] = int(maskN.sum())
    df = df[~maskN]

    headcount = len(df)
    report["재직인원"] = headcount

    # ── 라: AN 근무부서(26년 조직코드) = VLOOKUP(U, 연중부서코드매핑 B:D, 3) ──
    ym = yeonjung.copy()
    code2org, code2name = {}, {}
    for _, r in ym.iterrows():
        key = r.iloc[YEONJUNG_KEY]
        if key is None or str(key).strip() == "":
            continue
        key = str(key).strip()
        code2org[key] = r.iloc[YEONJUNG_VAL]
        code2name[key] = r.iloc[YEONJUNG_NAME] if ym.shape[1] > YEONJUNG_NAME else None

    df[ROSTER_AN] = df["근무부서"].map(code2org)
    # identity-default: 연중부서코드매핑에 없는 코드 = '연중 변경 없음'으로 보고 그대로 사용.
    #   (재직추출은 매달 새로 뽑으므로, 실제 코드변경은 D검증(근무≠부서)에서 드러남)
    미등록 = df[ROSTER_AN].isna()
    df.loc[미등록, ROSTER_AN] = df.loc[미등록, "근무부서"]

    # ── AO: 소속(S)==근무(U) 확인 ──
    df[ROSTER_AO] = df["근무부서"].eq(df["소속부서"])

    # ── 정보: 연중부서코드매핑에 없어 그대로 사용된 코드(신규/미변경) ──
    unmapped = df[미등록]
    if len(unmapped):
        ud = (unmapped.groupby(["근무부서", "근무부서명"], dropna=False)
              .agg(인원=("사번", "count"),
                   사번목록=("사번", lambda s: ", ".join(s.astype(str)[:5])))
              .reset_index()
              .rename(columns={"근무부서": "근무부서코드"}))
    else:
        ud = pd.DataFrame(columns=["근무부서코드", "근무부서명", "인원", "사번목록"])
    report["미등록코드_그대로사용_부서수"] = len(ud)
    report["미등록코드_그대로사용_인원"] = int(len(unmapped))

    # ── 예외②: 소속≠근무 ──
    mm = df[~df[ROSTER_AO]][
        ["사번", "성명", "소속부서명", "소속부서", "근무부서명", "근무부서"]
    ].copy()
    report["소속≠근무_인원"] = len(mm)

    return Stage2Result(
        roster=df.reset_index(drop=True),
        headcount=headcount,
        unmapped_depts=ud.reset_index(drop=True),
        mismatch_rows=mm.reset_index(drop=True),
        report=report,
    )