#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""초과근무 리포트 엔진 — 로컬 실행기 (Stage 1~4)

★ 이 스크립트는 당신의 PC에서만 동작합니다. 어떤 데이터도 외부로 전송하지 않습니다. ★

사용법:
  1) 아래 [설정] 부분의 파일 경로 4개와 '근무월'을 본인 환경에 맞게 수정
  2) 터미널에서:  python run_local.py
  3) 화면의 QA/예외 목록 확인 + 같은 폴더에 생기는 '_확인용_결과.xlsx' 열어보기
"""
import sys
import pandas as pd

from stage1_chogeun import build_chogeun_master
from stage2_jaejik import build_jaejik
from stage3_enrich import enrich_chogeun_master
from stage4_reports import build_reports
from stage_orgtable import build_org_table, org_table_to_26form
from stage5_recipients import build_team_recipients


from stage_orgtable import build_org_table, org_table_to_26form
from stage5_recipients import build_team_recipients, build_dept_recipients
import os
from stage6_distribute import build_team_file, build_dept_file


# ══════════════════════════ [설정] 여기만 고치세요 ══════════════════════════
정보검색_파일   = r"정보검색.xlsx"
정보검색_시트   = "Sheet1"

인사자료추출_파일 = r"인사자료추출.xlsx"
인사자료추출_시트 = "Sheet1"

관리_파일        = r"26년 초과근무 관리.xlsx"   # 연중부서코드매핑 시트가 든 파일
연중부서코드매핑_시트 = "연중부서코드매핑"

근무월 = "5월"     # 이번에 처리할 월 (정보검색 데이터의 해당 월)

결과파일 = r"_확인용_결과.xlsx"

# 배포 파일 템플릿 (가공·그래프 수식이 든 빈 템플릿). 있으면 자동으로 채워 생성.
팀별_템플릿   = r"초과근무_팀별.xlsx"
담당별_템플릿 = r"초과근무_담당별.xlsx"
# ═══════════════════════════════════════════════════════════════════════════


def 읽기(파일, 시트):
    # 사번·코드의 앞자리 0 보존을 위해 전부 문자열로 읽음
    return pd.read_excel(파일, sheet_name=시트, header=0, dtype=str).fillna("")


def 인사자료_열정렬(df):
    """인사자료추출의 열을 '헤더 이름' 기준으로 표준 순서(ROSTER_HEADERS)에 맞춘다.
    파일 버전마다 열 순서가 달라도(차차상위/직상위 위치 등) 정확히 매핑되게 한다.
    집계에 꼭 필요한 핵심 열이 이름으로 안 잡히면 즉시 알려준다."""
    from columns import ROSTER_HEADERS
    avail = {str(c).strip(): c for c in df.columns}
    핵심 = ["사번", "현원여부", "차차상위명", "차차상위", "차상위명", "차상위",
            "직상위명", "직상위", "소속부서명", "소속부서", "근무부서명", "근무부서"]
    빠짐 = [n for n in 핵심 if n not in avail]
    if 빠짐:
        print("‼ 인사자료추출에서 다음 핵심 열을 이름으로 찾지 못했습니다:", 빠짐)
        print("   실제 헤더:", list(df.columns))
        raise SystemExit("헤더 이름 불일치 — 위 목록을 알려주시면 매핑을 맞추겠습니다.")
    cols = []
    for name in ROSTER_HEADERS:
        cols.append(df[avail[name]].rename(name) if name in avail
                    else pd.Series([""] * len(df), name=name))
    return pd.concat(cols, axis=1)


def 구분선(t):
    print("\n" + "=" * 64 + f"\n  {t}\n" + "=" * 64)


def main():
    try:
        정보검색 = 읽기(정보검색_파일, 정보검색_시트)
        인사자료 = 읽기(인사자료추출_파일, 인사자료추출_시트)
        인사자료 = 인사자료_열정렬(인사자료)   # 헤더 이름 기준 표준 순서로 정렬
        연중매핑 = 읽기(관리_파일, 연중부서코드매핑_시트)
    except Exception as e:
        print("‼ 파일을 읽지 못했습니다. 경로/시트명을 확인하세요.\n  ", e)
        sys.exit(1)

    print(f"입력 행수 — 정보검색:{len(정보검색)}  인사자료:{len(인사자료)} "
          f" 연중매핑:{len(연중매핑)}")

    # ── Stage 1 : 정보검색 → 초근master ──
    s1 = build_chogeun_master(정보검색, geunmu_month=근무월)
    구분선("Stage 1 · 초근master 적재")
    print(s1.report)

    # ── Stage 2 : 인사자료추출 → 재직추출 ──
    s2 = build_jaejik(인사자료, 연중매핑)
    구분선("Stage 2 · 재직추출 (필터 + 근무부서 매핑)")
    print(s2.report)
    if len(s2.mismatch_rows):
        print(f"\n[참고] 소속≠근무 인원: {len(s2.mismatch_rows)}명 (근무부서 기준 집계, 확인용 엑셀 참조)")

    # ── 26년조직(팀→담당계층) 자동 생성 : 집계대상 재직추출에서 도출 ──
    org = build_org_table(s2.roster)
    조직매핑 = org_table_to_26form(org.org_table)   # 26년조직 위치형식
    구분선("26년조직 자동생성 (인사자료추출 기반)")
    print(org.report)
    if len(org.pure_work_tf):
        print("\n[근무전용 TF] native(소속=근무) 멤버가 없어 담당Ⅰ=자기자신으로 둔 팀:")
        print(org.pure_work_tf.to_string(index=False))

    # ── Stage 3 : 초근master 보강 ──
    s3 = enrich_chogeun_master(s1.master, s2.roster, 조직매핑)
    구분선("Stage 3 · 초근master 보강 (담당계층 + 검증)")
    print(s3.report)
    if len(s3.sabun_missing):
        print("\n[예외①] 초근 기록 있으나 재직명부에 없는 사번:")
        print(s3.sabun_missing.to_string(index=False))
    if len(s3.dam_unmapped):
        print("\n[예외②] 담당계층 매핑 안 된 조직코드:")
        print(s3.dam_unmapped.to_string(index=False))
    if len(s3.dept_mismatch):
        print(f"\n[예외③] 명부부서≠신청서부서: {len(s3.dept_mismatch)}건 (확인용 엑셀 참조)")

    # ── Stage 4 : 리포트 집계 (이번 달) ──
    s4 = build_reports(s3.master, s2.roster, 조직매핑, months=[근무월])
    구분선("Stage 4 · 리포트 집계 + 정합성 검증")
    for k, v in s4.qa.items():
        mark = ""
        if k == "인원_일치":
            mark = "  ← ✅" if v else "  ← ‼ 불일치!"
        if k == "시간_일치":
            mark = "  ← ✅" if v else "  ← ‼ 초근시간이 어딘가 새고 있음!"
        print(f"  {k}: {v}{mark}")

    # ── Stage 5 : 수신인(팀별) 생성 ──
    rcpt = build_team_recipients(s2.roster, s4.팀별["총초근"], org.org_table, months=[근무월])
    구분선("Stage 5 · 수신인 생성 (팀별)")
    print(rcpt.report)
    if len(rcpt.미지정):
        print("\n[팀장 미지정] 초근 발생했으나 팀장을 못 찾은 팀 (TF/신설 등 — 수동 지정 필요):")
        print(rcpt.미지정.to_string(index=False))

    # ── 확인용 엑셀 저장 ──
    with pd.ExcelWriter(결과파일, engine="openpyxl") as w:
        s3.master.to_excel(w, sheet_name="초근master", index=False)
        s4.인별.to_excel(w, sheet_name="1)인별", index=False)
        s4.팀별["인원"].to_excel(w, sheet_name="2)팀별_인원", index=False)
        s4.팀별["총초근"].to_excel(w, sheet_name="2)팀별_총초근", index=False)
        s4.팀별["평균"].to_excel(w, sheet_name="2)팀별_평균", index=False)
        s4.담당1["총초근"].to_excel(w, sheet_name="3)담당1_총초근", index=False)
        s4.담당2["총초근"].to_excel(w, sheet_name="4)담당2_총초근", index=False)
        # 자동생성 26년조직 + 예외 모음
        org.org_table.to_excel(w, sheet_name="26년조직_자동생성", index=False)
        org.pure_work_tf.to_excel(w, sheet_name="근무전용TF", index=False)
        rcpt.수신인.to_excel(w, sheet_name="수신인", index=False)
        rcpt.미지정.to_excel(w, sheet_name="수신인_팀장미지정", index=False)
        s2.mismatch_rows.to_excel(w, sheet_name="참고_소속근무다름", index=False)
        s3.sabun_missing.to_excel(w, sheet_name="예외_사번미발견", index=False)
        s3.dam_unmapped.to_excel(w, sheet_name="예외_담당매핑누락", index=False)
        s3.dept_mismatch.to_excel(w, sheet_name="예외_근무부서불일치", index=False)

    구분선("완료")
    print(f"확인용 결과가 저장되었습니다 →  {결과파일}")
    print("엑셀을 열어 초근master·리포트·예외 시트를 눈으로 확인하세요.")

    # ── Stage 6 : 배포 파일 생성 (템플릿이 있으면) ──
    구분선("Stage 6 · 배포 파일 생성")
    if os.path.exists(팀별_템플릿):
        out_t = f"초과근무_팀별_{근무월}.xlsx"
        build_team_file(팀별_템플릿, out_t, s4.인별, s4.팀별["평균"],
                        s4.팀별["총초근"], rcpt.수신인, 근무월)
        print(f"  ✅ 팀별 배포 파일 생성 → {out_t}  (1~{근무월} 채움 · 이전 월은 템플릿에서 캐리 · 차트 보존)")
    else:
        print(f"  (팀별 템플릿 '{팀별_템플릿}' 없음 → 배포 파일 건너뜀)")
    if not os.path.exists(담당별_템플릿):
        print(f"  (담당별 템플릿 '{담당별_템플릿}' 없음 → 담당별 건너뜀)")
    else:
        rcpt_d = build_dept_recipients(s2.roster, s4.담당1["총초근"], months=[근무월])
        out_d = f"초과근무_담당별_{근무월}.xlsx"
        build_dept_file(담당별_템플릿, out_d, s4.인별, s4.팀별["평균"], s4.팀별["총초근"],
                        s4.담당1["평균"], s4.담당1["총초근"], rcpt_d.수신인, 근무월, org.org_table)
        print(f"  ✅ 담당별 배포 파일 생성 → {out_d}  (수신인 {rcpt_d.report.get('발송대상',0)}명 · 미지정 {rcpt_d.report.get('수신인미지정',0)})")


if __name__ == "__main__":
    main()