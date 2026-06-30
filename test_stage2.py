"""Stage 4 검증 — 손계산 정답과 대조.
초근master(보강완료)·재직추출·26년조직을 직접 구성해 집계 숫자를 검증한다.
"""
import pandas as pd
from columns import CHOGEUN_MASTER_HEADERS, ROSTER_HEADERS, ROSTER_AN
from stage4_reports import build_reports

# ── 초근master(초근 4월) : 필요한 열만 채운 빈 66열 틀 ──
def mrow(사번, org, 근무구분, bn):
    d = {h: None for h in CHOGEUN_MASTER_HEADERS}
    d["근무월"] = "4월"; d["사번"] = 사번; d["성명"] = "이름"
    d["코드_변환(인사자료)"] = org      # B (조직코드)
    d["근무구분"] = 근무구분            # AA
    d["조정계"] = bn                    # BN
    return d

master = pd.DataFrame([
    mrow("001", "ORG_A", "초근", 2.0),
    mrow("001", "ORG_A", "초근", 3.0),   # 001 합계 5.0
    mrow("001", "ORG_A", "기타", 99.0),  # 근무구분≠초근 → 제외돼야 함
    mrow("002", "ORG_A", "초근", 4.0),   # ORG_A 총 = 9.0
    mrow("003", "ORG_B", "초근", 6.0),   # ORG_B 총 = 6.0
])

# ── 재직추출 : 사번 + AN(조직코드) ──
def rrow(사번, org):
    d = {h: None for h in ROSTER_HEADERS}
    d["사번"] = 사번; d["성명"] = "이름"; d["현원여부"] = "Y"
    d[ROSTER_AN] = org
    return d
roster = pd.DataFrame([rrow("001", "ORG_A"), rrow("002", "ORG_A"), rrow("003", "ORG_B")])
# 인원: ORG_A=2, ORG_B=1, 합 3

# ── 26년조직 : 팀 2개, 담당계층 ──
# 위치: B=조직코드 C=부서명 D=담당1명 E=담당1코드 F=담당2명 G=담당2코드 H=담당3명 I=담당3코드
org26 = pd.DataFrame([
    ["A팀", "ORG_A", "A팀", "a담당", "DAM_a", "x본부", "BON_x", "씨이오", "CEO"],
    ["B팀", "ORG_B", "B팀", "b담당", "DAM_b", "x본부", "BON_x", "씨이오", "CEO"],
])

res = build_reports(master, roster, org26, months=["4월"])
print("QA:", res.qa)
print("\n인별:\n", res.인별.to_string(index=False))
print("\n팀별 총초근:\n", res.팀별["총초근"][["근무부서코드", "4월"]].to_string(index=False))
print("팀별 인원:\n", res.팀별["인원"][["근무부서코드", "4월"]].to_string(index=False))
print("팀별 평균:\n", res.팀별["평균"][["근무부서코드", "4월"]].to_string(index=False))
print("\n담당1(담당Ⅰ) 총초근/인원/평균:")
print(res.담당1["총초근"][["담당1코드", "4월"]].to_string(index=False))
print(res.담당1["평균"][["담당1코드", "4월"]].to_string(index=False))
print("\n담당2(담당Ⅱ) 총초근/평균:")
print(res.담당2["총초근"][["담당2코드", "4월"]].to_string(index=False))
print(res.담당2["평균"][["담당2코드", "4월"]].to_string(index=False))

# ── 검증 ──
ip = res.인별.set_index("사번")["4월"].to_dict()
assert abs(ip["001"] - 5.0) < 1e-9      # 2+3, '기타' 99 제외
assert abs(ip["002"] - 4.0) < 1e-9
assert abs(ip["003"] - 6.0) < 1e-9
assert res.인별.set_index("사번").loc["001", "근무부서"] == "A팀"

tt = res.팀별["총초근"].set_index("근무부서코드")["4월"].to_dict()
assert abs(tt["ORG_A"] - 9.0) < 1e-9
assert abs(tt["ORG_B"] - 6.0) < 1e-9
th = res.팀별["인원"].set_index("근무부서코드")["4월"].to_dict()
assert th["ORG_A"] == 2 and th["ORG_B"] == 1
ta = res.팀별["평균"].set_index("근무부서코드")["4월"].to_dict()
assert abs(ta["ORG_A"] - 4.5) < 1e-9    # 9/2
assert abs(ta["ORG_B"] - 6.0) < 1e-9

# 담당Ⅰ 롤업
d1t = res.담당1["총초근"].set_index("담당1코드")["4월"].to_dict()
d1a = res.담당1["평균"].set_index("담당1코드")["4월"].to_dict()
assert abs(d1t["DAM_a"] - 9.0) < 1e-9 and abs(d1t["DAM_b"] - 6.0) < 1e-9
assert abs(d1a["DAM_a"] - 4.5) < 1e-9 and abs(d1a["DAM_b"] - 6.0) < 1e-9

# 담당Ⅱ 롤업 (두 팀이 같은 본부 BON_x)
d2t = res.담당2["총초근"].set_index("담당2코드")["4월"].to_dict()
d2h = res.담당2["인원"].set_index("담당2코드")["4월"].to_dict()
d2a = res.담당2["평균"].set_index("담당2코드")["4월"].to_dict()
assert abs(d2t["BON_x"] - 15.0) < 1e-9   # 9+6
assert d2h["BON_x"] == 3                  # 2+1
assert abs(d2a["BON_x"] - 5.0) < 1e-9    # 15/3

# 정합성
assert res.qa["인원_일치"] is True and res.qa["팀인원합"] == 3
assert res.qa["시간_일치"] is True and abs(res.qa["팀집계시간"] - 15.0) < 1e-9

print("\n✅ 모든 검증 통과 — 인별·팀별·담당 집계 + 평균 + 정합성(인원/시간) 정확")
