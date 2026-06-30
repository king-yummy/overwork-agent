"""Stage 3 검증 — Stage1·Stage2를 실제로 연결해 합성 데이터로 대조.
(세 단계가 맞물려 도는지까지 확인)
"""
import pandas as pd
from stage1_chogeun import build_chogeun_master
from stage2_jaejik import build_jaejik
from stage3_enrich import enrich_chogeun_master

# ── 정보검색 (Stage1 입력) : 53열, 위치로 채움 ──
NJ = 53
def jrow(사번, 부서코드, 근무구분="초근", 야간=0):
    r = [""] * NJ
    r[0] = 1            # 구분자
    r[1] = "20260401"   # 근무일자
    r[2] = 사번
    r[3] = "이름"
    r[5] = 부서코드     # F 부서코드 → 초근master R
    r[9] = "작업"       # 업무내용
    r[14] = 근무구분
    r[16] = 야간        # 야간근로시간
    return r

jeongbo = pd.DataFrame([
    jrow("0001", "ORG_FC", 야간=3),   # 정상 · D=True 기대
    jrow("0002", "XX000",  야간=2),   # 근무≠부서 (B=ORG_DC ≠ R=XX000)
    jrow("0003", "ORG_ZZ", 야간=1),   # 담당 매핑누락 (ORG_ZZ 없음)
    jrow("0099", "ORG_FC", 야간=4),   # 사번 미발견 (명부에 없음)
])

# ── 인사자료추출 (Stage2 입력) : 39열 ──
NI = 39
def prow(사번, 근무부서, 소속부서=None):
    r = [""] * NI
    r[0] = 사번
    r[1] = "이름"
    r[2] = "Y"                          # 현원
    r[18] = 소속부서 or 근무부서        # S 소속부서
    r[19] = "부서명"                    # T 근무부서명
    r[20] = 근무부서                    # U 근무부서(코드)
    return r

insa = pd.DataFrame([
    prow("0001", "FC004"),
    prow("0002", "DCCI0"),
    prow("0003", "ZZ001"),
    # 0099 명부에 없음 → 사번 미발견 유발
])
# 연중부서코드매핑 : 근무부서코드 → 26년 조직코드
yeonjung = pd.DataFrame([
    ["수북구축팀", "FC004", "수북구축팀", "ORG_FC", "수북조직"],
    ["DC센터",    "DCCI0", "DC센터",    "ORG_DC", "DC조직"],
    ["미지정",    "ZZ001", "미지정",    "ORG_ZZ", "미지정조직"],
])
# 26년조직 : B=조직코드 C=부서명 D=담당1명 E=담당1코드 F=담당2명 G=담당2코드 H=담당3명 I=담당3코드
# (ORG_ZZ 는 일부러 누락)
org26 = pd.DataFrame([
    ["수북구축팀", "ORG_FC", "수북구축팀", "구축담당", "DAM1", "인프라본부", "DAM2", "기술부문", "DAM3"],
    ["DC센터",     "ORG_DC", "DC센터",    "DC담당",   "DDM1", "DC본부",     "DDM2", "기술부문", "DAM3"],
])

# ── 파이프라인 연결 ──
s1 = build_chogeun_master(jeongbo, geunmu_month="4월")
s2 = build_jaejik(insa, yeonjung)
s3 = enrich_chogeun_master(s1.master, s2.roster, org26)

m = s3.master
print("Stage3 리포트:", s3.report)
cols = ["사번", "부서코드", "코드_변환(인사자료)", "근무부서명", "담당1", "담당2", "담당3", "검증_변환=소속"]
print("\n초근master 보강 결과:")
print(m[cols].to_string(index=False))
print("\n예외① 사번미발견:\n", s3.sabun_missing.to_string(index=False))
print("\n예외② 담당매핑누락:\n", s3.dam_unmapped.to_string(index=False))
print("\n예외③ 근무≠부서:\n", s3.dept_mismatch.to_string(index=False))

# ── 검증 ──
byid = m.set_index("사번")
# 0001 정상
assert byid.loc["0001", "코드_변환(인사자료)"] == "ORG_FC"
assert byid.loc["0001", "근무부서명"] == "수북구축팀"
assert byid.loc["0001", "담당1"] == "구축담당"
assert byid.loc["0001", "담당1코드"] == "DAM1"
assert byid.loc["0001", "담당2"] == "인프라본부"
assert byid.loc["0001", "담당3"] == "기술부문"
assert byid.loc["0001", "검증_변환=소속"] == True       # ORG_FC == ORG_FC
# 0002 근무≠부서
assert byid.loc["0002", "코드_변환(인사자료)"] == "ORG_DC"
assert byid.loc["0002", "검증_변환=소속"] == False      # ORG_DC != XX000
# 0003 담당 매핑누락 → 담당계층 None
assert byid.loc["0003", "코드_변환(인사자료)"] == "ORG_ZZ"
assert pd.isna(byid.loc["0003", "근무부서명"]) or byid.loc["0003", "근무부서명"] is None
# 0099 사번 미발견 → B None
assert pd.isna(byid.loc["0099", "코드_변환(인사자료)"])

r = s3.report
assert r["사번_미발견_명"] == 1 and set(s3.sabun_missing["사번"]) == {"0099"}
assert r["담당매핑누락_조직수"] == 1 and set(s3.dam_unmapped["조직코드"]) == {"ORG_ZZ"}
assert r["근무≠부서_건수"] == 1 and set(s3.dept_mismatch["사번"]) == {"0002"}
assert r["초근master_고유인원"] == 3   # 0001,0002,0003 (0099 제외)

print("\n✅ 모든 검증 통과 — A~M 보강 + 담당계층 + D검증 + 예외 3종 정확 (Stage1·2·3 연동 확인)")
