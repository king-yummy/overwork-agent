# -*- coding: utf-8 -*-
"""26년조직(팀→담당계층) 자동 생성 — 집계대상 재직추출(Stage2 출력)에서 도출.

★ 반드시 'Stage 2를 거친 재직추출'(9천번대·알파벳·현원N 제거 완료)을 입력으로 받는다.
  그래야 팀 목록 모집단이 집계 모집단과 일치하고, 9천번대 1명이 만드는 유령팀
  (예: SCFX0)이 생기지 않는다.

키:  ROSTER_AN(=근무부서의 연중 조직코드)로 그룹 → 보고 집계 키와 동일.
계층 도출:
  · 담당Ⅲ(차차상위)·담당Ⅱ(차상위) = 멤버 최빈값 (실측 99.6~100% 일치)
  · 담당Ⅰ(직상위):
      - native(소속==근무, ROSTER_AO=True)가 있으면 → native 직상위 최빈값
        (일반팀=상위담당 / native TF=자기자신, 데이터가 그렇게 줌)
      - native가 없는 근무전용 TF → 담당Ⅰ = 팀 자신 (기존 TF 매핑 방식과 동일)
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import ROSTER_POS, ROSTER_AN, ROSTER_AO

P = ROSTER_POS


@dataclass
class OrgTableResult:
    org_table: pd.DataFrame
    pure_work_tf: pd.DataFrame = field(default_factory=pd.DataFrame)
    report: dict = field(default_factory=dict)


def _mode(series):
    s = series.astype(str).str.strip()
    s = s[s != ""]
    return s.mode().iloc[0] if len(s) else None


def build_org_table(roster: pd.DataFrame) -> OrgTableResult:
    """roster : Stage2 출력 재직추출 (원본 A~AM 위치 유지 + ROSTER_AN, ROSTER_AO)."""
    df = roster.copy().reset_index(drop=True)
    an = df[ROSTER_AN].astype(str).str.strip()
    ao = df[ROSTER_AO] == True

    rows, pure = [], []
    # 명부 표준 헤더 이름으로 직접 접근 (위치 의존 제거)
    NM = {"담당3코드": "차차상위", "담당3명": "차차상위명",
          "담당2코드": "차상위", "담당2명": "차상위명",
          "담당1코드": "직상위", "담당1명": "직상위명", "근무부서명": "근무부서명"}
    for code in sorted(set(an) - {"", "nan"}):
        m = an == code
        grp = df[m]
        nat = df[m & ao]

        d3c, d3n = _mode(grp[NM["담당3코드"]]), _mode(grp[NM["담당3명"]])
        d2c, d2n = _mode(grp[NM["담당2코드"]]), _mode(grp[NM["담당2명"]])
        부서명 = _mode(grp[NM["근무부서명"]])

        if len(nat):
            d1c, d1n = _mode(nat[NM["담당1코드"]]), _mode(nat[NM["담당1명"]])
            is_pure = False
        else:
            d1c, d1n = code, 부서명           # 근무전용 TF: 자기 자신 = 담당Ⅰ
            is_pure = True
            pure.append({"조직코드": code, "근무부서명": 부서명, "인원": int(m.sum()),
                         "담당2명": d2n, "담당3명": d3n})

        rows.append({"조직코드": code, "근무부서명": 부서명,
                     "담당1명": d1n, "담당1코드": d1c,
                     "담당2명": d2n, "담당2코드": d2c,
                     "담당3명": d3n, "담당3코드": d3c,
                     "근무전용TF": is_pure})

    org_table = pd.DataFrame(rows, columns=[
        "조직코드", "근무부서명", "담당1명", "담당1코드",
        "담당2명", "담당2코드", "담당3명", "담당3코드", "근무전용TF"])
    pure_work_tf = pd.DataFrame(pure, columns=["조직코드", "근무부서명", "인원", "담당2명", "담당3명"])
    report = {"팀수": len(org_table), "근무전용TF수": len(pure_work_tf)}
    return OrgTableResult(org_table=org_table, pure_work_tf=pure_work_tf, report=report)


def org_table_to_26form(org_table: pd.DataFrame) -> pd.DataFrame:
    """자동생성 org_table → 26년조직_특수문자제외 위치형식(A~I)으로 변환.
    (B=조직코드 키, C=부서명, D=담당1명,E=담당1코드,F=담당2명,G=담당2코드,H=담당3명,I=담당3코드)"""
    return pd.DataFrame({
        0: org_table["근무부서명"],
        1: org_table["조직코드"],
        2: org_table["근무부서명"],
        3: org_table["담당1명"], 4: org_table["담당1코드"],
        5: org_table["담당2명"], 6: org_table["담당2코드"],
        7: org_table["담당3명"], 8: org_table["담당3코드"],
    })