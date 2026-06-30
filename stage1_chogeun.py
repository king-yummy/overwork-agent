"""Stage 1 — 정보검색 → 초근master 행 생성.

매뉴얼 다·마·바 단계를 결정론적으로 구현한다:
  · 다 : 업무내용 '교대근무' 행 삭제, 근무구분 '비대상' → '초근'
  · 마 : 사번이 P 또는 Y로 시작하는 행 삭제
  · 바 : 정보검색 B~BA → 초근master N~BM 으로 재배치, BN(조정계) 계산

A~M열(월/부서/담당 보강)은 명부·매핑표가 필요하므로 다음 단계(enrich)에서 채운다.
여기서는 '근무월'만 입력받아 A열에 채운다.

모든 처리 결과와 함께 '예외/판단 목록'에 들어갈 정보를 같이 반환한다.
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import (
    CHOGEUN_MASTER_HEADERS, JEONGBO_HEADERS,
    BN_MAX_COLS, BN_DIV60_COLS, BN_SUB_COLS,
)

# 정보검색에서 숫자로 다뤄야 하는(=BN 계산에 쓰는) 열
_NUMERIC_COLS = BN_MAX_COLS + BN_DIV60_COLS + BN_SUB_COLS


@dataclass
class Stage1Result:
    master: pd.DataFrame                 # 초근master N~BN 부분이 채워진 표 (A~M은 비어있음)
    report: dict = field(default_factory=dict)   # 무엇이 몇 건 처리/삭제됐는지 (QA용)


def _to_num(series: pd.Series) -> pd.Series:
    """빈칸·문자 → 0. Excel 산술에서 빈칸은 0으로 취급되는 동작을 재현."""
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def build_chogeun_master(
    jeongbo: pd.DataFrame,
    geunmu_month: str,
    *,
    gyodae_match: str = "exact",   # '교대근무' 판정: exact | contains
) -> Stage1Result:
    """정보검색 DataFrame → 초근master 행.

    jeongbo : 정보검색 Sheet1 (헤더 제외, 53개 열 A~BA). 컬럼명은 무시하고
              위치로 읽으므로 어떤 헤더든 상관없다. 사번은 문자열이어야 한다.
    geunmu_month : 'A열 근무월'에 채울 값 (예: '4월').
    """
    df = jeongbo.copy()
    df.columns = JEONGBO_HEADERS[: df.shape[1]]   # 위치 기반으로 표준 헤더 부여
    n0 = len(df)
    report = {"입력행": n0}

    # 사번을 leading-zero 보존 문자열로
    df["사번"] = df["사번"].astype(str).str.strip()

    # --- 다: 업무내용 '교대근무' 삭제 ---
    work = df["업무내용"].astype(str).str.strip()
    if gyodae_match == "contains":
        mask_gyodae = work.str.contains("교대근무", na=False)
    else:
        mask_gyodae = work.eq("교대근무")
    report["교대근무_삭제"] = int(mask_gyodae.sum())
    df = df[~mask_gyodae]

    # --- 마: 집계 대상은 '숫자로 시작하는 사번'만.
    #         알파벳으로 시작(임원 M·지원선수 PF·임원기사 PH·외부 P/Y 등)은 비대상 → 삭제 ---
    mask_nondigit = ~df["사번"].str[:1].str.isdigit().fillna(False)
    report["사번_비숫자시작_삭제"] = int(mask_nondigit.sum())
    df = df[~mask_nondigit]

    # --- 다: 근무구분 '비대상' → '초근' ---
    gubun = df["근무구분"].astype(str).str.strip()
    mask_bidae = gubun.eq("비대상")
    report["비대상→초근_변환"] = int(mask_bidae.sum())
    df.loc[mask_bidae, "근무구분"] = "초근"

    report["출력행"] = len(df)

    # --- 바: 초근master 틀에 옮겨담기 ---
    master = pd.DataFrame(index=df.index, columns=CHOGEUN_MASTER_HEADERS)
    # N~BM : 정보검색 B~BA (구분자 제외) 를 같은 헤더명으로 복사
    for col in JEONGBO_HEADERS[1:]:            # '구분자' 제외
        master[col] = df[col].values
    # A : 근무월
    master["근무월"] = geunmu_month

    # --- BN(조정계) 계산 ---
    maxpart = pd.concat([_to_num(master[c]) for c in BN_MAX_COLS], axis=1).max(axis=1)
    div60 = sum(_to_num(master[c]) / 60.0 for c in BN_DIV60_COLS)
    subpart = sum(_to_num(master[c]) for c in BN_SUB_COLS)
    master["조정계"] = maxpart - div60 - subpart

    master = master.reset_index(drop=True)
    return Stage1Result(master=master, report=report)