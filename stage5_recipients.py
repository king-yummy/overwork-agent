# -*- coding: utf-8 -*-
"""Stage 5 — 수신인(RPA 발송 대상) 생성.

매뉴얼 사단계:
  · 2)팀별에서 초근(연합)>0 인 팀만 대상.
  · 각 팀의 팀장을 찾아 수신인 시트(조직코드·조직명·사번·수신인·이메일) 구성.
  · 팀장 없는 팀(TF·신설 등)은 '미지정'으로 표시 → 사람이 채우거나(이메일 정보) 처리.
  · CEO에게는 발송하지 않음.

팀장 식별: 재직추출(Stage2 roster)에서 직책명='팀장' → 근무부서 조직코드(AN)로 매칭.
이메일: SKEMAIL.
"""
from dataclasses import dataclass, field
import pandas as pd
from columns import ROSTER_AN

CEO_CODE = "AA005"


@dataclass
class RecipientResult:
    수신인: pd.DataFrame                       # 조직코드·조직명·사번·수신인·이메일
    미지정: pd.DataFrame = field(default_factory=pd.DataFrame)  # 팀장 못 찾은 팀
    report: dict = field(default_factory=dict)


def build_team_recipients(roster: pd.DataFrame,
                          팀별_총초근: pd.DataFrame,
                          org_table: pd.DataFrame,
                          months) -> RecipientResult:
    return _build_recipients(roster, 팀별_총초근, "근무부서코드", "근무부서명",
                             org_table, dict(zip(org_table["조직코드"].astype(str), org_table["근무부서명"])),
                             {"팀장"}, months)


def build_dept_recipients(roster: pd.DataFrame,
                          담당1_총초근: pd.DataFrame,
                          months) -> RecipientResult:
    """담당Ⅰ 단위 수신인 = 직책 '담당' 또는 '실장' (근무부서코드 = 담당Ⅰ코드)."""
    name_by = dict(zip(담당1_총초근["담당1코드"].astype(str), 담당1_총초근["담당1"]))
    return _build_recipients(roster, 담당1_총초근, "담당1코드", "담당1",
                             None, name_by, {"담당", "실장"}, months)


def _build_recipients(roster, 총초근, code_col, name_col, org_table, 부서명_by,
                      직책set, months) -> RecipientResult:
    rs = roster.copy()
    rs[ROSTER_AN] = rs[ROSTER_AN].astype(str).str.strip()
    jik = rs["직책명"].astype(str).str.strip() if "직책명" in rs.columns else pd.Series([""] * len(rs))
    leaders = rs[jik.isin(직책set)].copy()
    by = {}
    for _, r in leaders.iterrows():
        code = str(r[ROSTER_AN]).strip()
        if code and code not in by:
            by[code] = (str(r["사번"]).strip(), r["성명"], r.get("SKEMAIL", ""))

    tt = 총초근.copy()
    tt["_합"] = tt[list(months)].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    rows, miss = [], []
    for _, tr in tt[tt["_합"] > 0].iterrows():
        code = str(tr[code_col]).strip()
        if code == CEO_CODE or not code:
            continue
        조직명 = 부서명_by.get(code, tr.get(name_col, ""))
        if code in by:
            사번, 수신, email = by[code]
            rows.append({"조직코드": code, "조직명": 조직명, "사번": 사번, "수신인": 수신, "이메일": email})
        else:
            rows.append({"조직코드": code, "조직명": 조직명, "사번": "", "수신인": "", "이메일": ""})
            miss.append({"조직코드": code, "조직명": 조직명})
    수신인 = pd.DataFrame(rows, columns=["조직코드", "조직명", "사번", "수신인", "이메일"])
    미지정 = pd.DataFrame(miss, columns=["조직코드", "조직명"])
    return RecipientResult(수신인=수신인, 미지정=미지정,
                           report={"발송대상": len(수신인), "수신인미지정": len(미지정)})


def _build_team_recipients_OLD(roster: pd.DataFrame,
                          팀별_총초근: pd.DataFrame,
                          org_table: pd.DataFrame,
                          months) -> RecipientResult:
    rs = roster.copy()
    rs[ROSTER_AN] = rs[ROSTER_AN].astype(str).str.strip()

    # 팀장 후보: 직책명=='팀장'
    jikchaek = rs["직책명"].astype(str).str.strip() if "직책명" in rs.columns else pd.Series([""] * len(rs))
    팀장 = rs[jikchaek == "팀장"].copy()
    # 팀(AN)별 팀장 1명 (여러 명이면 첫 번째)
    팀장_by = {}
    for _, r in 팀장.iterrows():
        code = str(r[ROSTER_AN]).strip()
        if code and code not in 팀장_by:
            팀장_by[code] = (str(r["사번"]).strip(), r["성명"], r.get("SKEMAIL", ""))

    부서명_by = dict(zip(org_table["조직코드"].astype(str), org_table["근무부서명"]))

    # 초근(해당 월 합)>0 인 팀
    tt = 팀별_총초근.copy()
    tt["_합"] = tt[list(months)].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    대상 = tt[tt["_합"] > 0]["근무부서코드"].astype(str).str.strip().tolist()

    rows, miss = [], []
    for code in 대상:
        if code == CEO_CODE:                       # CEO 발송 제외
            continue
        조직명 = 부서명_by.get(code, "")
        if code in 팀장_by:
            사번, 수신, email = 팀장_by[code]
            rows.append({"조직코드": code, "조직명": 조직명, "사번": 사번,
                         "수신인": 수신, "이메일": email})
        else:
            rows.append({"조직코드": code, "조직명": 조직명, "사번": "",
                         "수신인": "", "이메일": ""})
            miss.append({"조직코드": code, "조직명": 조직명})

    수신인 = pd.DataFrame(rows, columns=["조직코드", "조직명", "사번", "수신인", "이메일"])
    미지정 = pd.DataFrame(miss, columns=["조직코드", "조직명"])
    report = {"발송대상팀수": len(수신인), "팀장미지정": len(미지정)}
    return RecipientResult(수신인=수신인, 미지정=미지정, report=report)