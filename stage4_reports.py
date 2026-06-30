"""Stage 4 — 리포트 집계 (인별 / 팀별 / 담당1 / 담당2).

실제 셀 수식 기준:
  · 인별 월값  = SUMIFS(초근master BN, 근무구분="초근", 사번 일치, 근무월 일치)
  · 팀별 인원  = COUNTIFS(재직추출 AN = 팀코드)                  (월 무관, 명부 스냅샷)
  · 팀별 총초근 = SUMIFS(초근master BN, 근무구분="초근", B=팀코드, 근무월)
  · 팀별 평균  = 총초근 / 인원
  · 담당1     = 팀별을 담당Ⅰ코드로 롤업, 담당2 = 담당Ⅱ코드로 롤업

팀 목록·담당계층은 26년조직 표에서 가져온다(엑셀이 팀별 시트를 다시 합산하는 것과
동일하게, 담당 롤업은 '팀별 결과'를 그룹합산 → 담당 합 = 팀 합 자동 보장).

자동 정합성 검증:
  · 팀 인원 합 == 재직추출 인원
  · 팀 총초근 합 == 초근master(초근) BN 합 중 '팀에 매핑된' 분
  · 어느 팀에도 안 잡힌 초근시간(담당 매핑누락분)을 잔차로 표면화
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import ROSTER_AN, ORG26_CODE, ORG26

CHOGEUN = "초근"


@dataclass
class Stage4Result:
    인별: pd.DataFrame
    팀별: dict          # {'인원':df,'총초근':df,'평균':df}
    담당1: dict
    담당2: dict
    qa: dict = field(default_factory=dict)


def _team_table(org26: pd.DataFrame) -> pd.DataFrame:
    """26년조직 → 팀 목록 (담당Ⅲ/Ⅱ/Ⅰ 코드·명 + 근무부서코드·명)."""
    rows = []
    seen = set()
    for _, r in org26.iterrows():
        code = r.iloc[ORG26_CODE]
        if code is None or str(code).strip() == "":
            continue
        code = str(code).strip()
        if code in seen:
            continue
        seen.add(code)
        rows.append({
            "담당3코드": r.iloc[ORG26["담당3코드"]], "담당3": r.iloc[ORG26["담당3"]],
            "담당2코드": r.iloc[ORG26["담당2코드"]], "담당2": r.iloc[ORG26["담당2"]],
            "담당1코드": r.iloc[ORG26["담당1코드"]], "담당1": r.iloc[ORG26["담당1"]],
            "근무부서코드": code, "근무부서명": r.iloc[ORG26["근무부서명"]],
        })
    cols = ["담당3코드", "담당3", "담당2코드", "담당2", "담당1코드", "담당1",
            "근무부서코드", "근무부서명"]
    return pd.DataFrame(rows, columns=cols)


def build_reports(master: pd.DataFrame, roster: pd.DataFrame,
                  org26: pd.DataFrame, months=None) -> Stage4Result:
    m = master.copy()
    m["사번"] = m["사번"].astype(str).str.strip()
    rs = roster.copy()
    rs["사번"] = rs["사번"].astype(str).str.strip()
    rs[ROSTER_AN] = rs[ROSTER_AN].astype(str).str.strip()

    if months is None:
        months = [x for x in m["근무월"].dropna().unique() if str(x).strip() != ""]
        months = sorted(months, key=lambda s: int(str(s).replace("월", "")) if str(s).replace("월", "").isdigit() else 99)

    chg = m[m["근무구분"].astype(str).str.strip() == CHOGEUN].copy()
    chg["조정계"] = pd.to_numeric(chg["조정계"], errors="coerce").fillna(0.0)
    chg["org"] = chg["코드_변환(인사자료)"].astype(str).str.strip()

    teams = _team_table(org26)
    team_codes = set(teams["근무부서코드"])

    # ── 인별 ──
    per_person_month = (chg.groupby(["사번", "근무월"])["조정계"].sum().unstack(fill_value=0.0))
    code_by_sabun = dict(zip(rs["사번"], rs[ROSTER_AN]))
    name_by_code = dict(zip(teams["근무부서코드"], teams["근무부서명"]))
    인별 = rs[["사번", "성명"]].copy()
    인별["근무코드"] = 인별["사번"].map(code_by_sabun)
    인별["근무부서"] = 인별["근무코드"].map(name_by_code)
    for mo in months:
        인별[mo] = 인별["사번"].map(
            per_person_month[mo] if mo in per_person_month.columns else {}).fillna(0.0)

    # ── 팀별 인원 (월 무관 스냅샷 → 각 월에 동일 적용) ──
    head_by_code = rs.groupby(ROSTER_AN).size()
    인원 = teams.copy()
    for mo in months:
        인원[mo] = 인원["근무부서코드"].map(head_by_code).fillna(0).astype(int)

    # ── 팀별 총초근 ──
    team_month = (chg.groupby(["org", "근무월"])["조정계"].sum().unstack(fill_value=0.0))
    총초근 = teams.copy()
    for mo in months:
        총초근[mo] = 총초근["근무부서코드"].map(
            team_month[mo] if mo in team_month.columns else {}).fillna(0.0)

    # ── 팀별 평균 = 총초근/인원 ──
    평균 = teams.copy()
    for mo in months:
        평균[mo] = (총초근[mo] / 인원[mo].replace(0, pd.NA)).fillna(0.0)

    팀별 = {"인원": 인원, "총초근": 총초근, "평균": 평균}

    # ── 담당 롤업 (팀별 결과를 그룹합산) ──
    def rollup(level_code, level_name):
        keys = [level_code, level_name]
        head = 인원.groupby(keys, dropna=False)[list(months)].sum().reset_index()
        tot = 총초근.groupby(keys, dropna=False)[list(months)].sum().reset_index()
        avg = tot.copy()
        for mo in months:
            avg[mo] = (tot[mo] / head[mo].replace(0, pd.NA)).fillna(0.0)
        return {"인원": head, "총초근": tot, "평균": avg}

    담당1 = rollup("담당1코드", "담당1")   # 3)담당1 : 담당Ⅰ 기준
    담당2 = rollup("담당2코드", "담당2")   # 4)담당2 : 담당Ⅱ 기준

    # ── QA 정합성 ──
    team_head_sum = int(인원[months[0]].sum()) if months else 0
    roster_head = len(rs)
    ot_total = float(chg["조정계"].sum())
    ot_mapped = float(chg[chg["org"].isin(team_codes)]["조정계"].sum())
    ot_unmapped = ot_total - ot_mapped
    qa = {
        "처리월": list(months),
        "팀인원합": team_head_sum,
        "재직인원": roster_head,
        "인원_일치": team_head_sum == roster_head,
        "초근총시간": round(ot_total, 4),
        "팀집계시간": round(ot_mapped, 4),
        "미매핑잔차시간": round(ot_unmapped, 4),
        "시간_일치": abs(ot_unmapped) < 1e-6,
    }

    return Stage4Result(인별=인별, 팀별=팀별, 담당1=담당1, 담당2=담당2, qa=qa)
