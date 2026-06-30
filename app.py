import streamlit as st
import pandas as pd
import os
import io

# 기존 프로젝트 모듈 임포트
from stage1_chogeun import build_chogeun_master
from stage2_jaejik import build_jaejik
from stage3_enrich import enrich_chogeun_master
from stage4_reports import build_reports
from stage_orgtable import build_org_table, org_table_to_26form

# 페이지 기본 설정
st.set_page_config(page_title="초과근무 AI Agent", layout="wide")
st.title("🤖 초과근무 리포트 자동생성 에이전트")
st.markdown("데이터를 업로드하시면 예외 사항을 분석하고 자동으로 리포트를 생성해 드립니다.")

# 사이드바: 데이터 업로드
with st.sidebar:
    st.header("📂 데이터 업로드")
    file_jeongbo = st.file_uploader("1. 정보검색 파일 (초과근무 원본)", type=["xlsx", "xls"])
    file_insa = st.file_uploader("2. 인사자료추출 파일", type=["xlsx", "xls"])
    file_manage = st.file_uploader("3. 26년 초과근무 관리 (연중부서코드매핑)", type=["xlsx", "xls"])
    
    target_month = st.selectbox("집계 근무월 선택", ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"])
    
    run_btn = st.button("🚀 에이전트 실행", use_container_width=True)

# 메인 화면: 에이전트 실행 로직
if run_btn:
    if not (file_jeongbo and file_insa and file_manage):
        st.error("세 가지 파일을 모두 업로드해 주세요!")
    else:
        with st.spinner("AI 에이전트가 데이터를 분석하고 있습니다..."):
            try:
                # 1. 파일 읽기 (문자열로 읽기 유지)
                정보검색 = pd.read_excel(file_jeongbo, sheet_name=0, dtype=str).fillna("")
                인사자료 = pd.read_excel(file_insa, sheet_name=0, dtype=str).fillna("")
                연중매핑 = pd.read_excel(file_manage, sheet_name="연중부서코드매핑", dtype=str).fillna("")
                
                # 인사자료 열정렬 (기존 run_local.py 로직)
                from columns import ROSTER_HEADERS
                avail = {str(c).strip(): c for c in 인사자료.columns}
                cols = [인사자료[avail[name]].rename(name) if name in avail else pd.Series([""] * len(인사자료), name=name) for name in ROSTER_HEADERS]
                인사자료 = pd.concat(cols, axis=1)

                # 2. Stage 로직 실행
                s1 = build_chogeun_master(정보검색, geunmu_month=target_month)
                s2 = build_jaejik(인사자료, 연중매핑)
                org = build_org_table(s2.roster)
                조직매핑 = org_table_to_26form(org.org_table)
                s3 = enrich_chogeun_master(s1.master, s2.roster, 조직매핑)
                s4 = build_reports(s3.master, s2.roster, 조직매핑, months=[target_month])

                # 3. 챗봇 UI 흉내내기 (결과 브리핑)
                st.success("✨ 데이터 처리가 완료되었습니다!")
                
                with st.chat_message("assistant"):
                    st.write(f"안녕하세요! {target_month} 초과근무 데이터 분석 결과를 브리핑해 드립니다.")
                    
                    # QA 정합성 브리핑
                    is_valid = s4.qa.get('시간_일치', False)
                    if is_valid:
                        st.write(f"✅ **정합성 검증 완료:** 총 초과근무 시간 **{s4.qa.get('초근총시간', 0)}시간**이 누락 없이 각 부서로 정확히 배분되었습니다.")
                    else:
                        st.error(f"⚠️ **주의:** 초과근무 시간이 어딘가 새고 있습니다. (잔차: {s4.qa.get('미매핑잔차시간')}시간)")
                    
                    # 예외 사항 브리핑
                    exceptions = []
                    if len(s3.sabun_missing) > 0:
                        exceptions.append(f"초근 기록은 있으나 명부에 없는 사번이 **{len(s3.sabun_missing)}명** 있습니다.")
                    if len(s3.dept_mismatch) > 0:
                        exceptions.append(f"명부 부서와 신청서 부서가 다른 인원이 **{len(s3.dept_mismatch)}명** 있습니다.")
                        
                    if exceptions:
                        st.warning("🔍 **확인이 필요한 특이사항:**\n" + "\n".join([f"- {e}" for e in exceptions]))
                    else:
                        st.info("👍 이번 달은 확인이 필요한 특이사항(예외 케이스)이 발견되지 않았습니다.")

                # 4. 결과 엑셀 다운로드 (확인용 결과 파일 생성)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as w:
                    s3.master.to_excel(w, sheet_name="초근master", index=False)
                    s4.인별.to_excel(w, sheet_name="1)인별", index=False)
                    s4.팀별["총초근"].to_excel(w, sheet_name="2)팀별_총초근", index=False)
                    if len(s3.sabun_missing) > 0:
                        s3.sabun_missing.to_excel(w, sheet_name="예외_사번미발견", index=False)
                    if len(s3.dept_mismatch) > 0:
                        s3.dept_mismatch.to_excel(w, sheet_name="예외_근무부서불일치", index=False)
                
                output.seek(0)
                st.download_button(
                    label="📊 가공 완료된 리포트 다운로드 (Excel)",
                    data=output,
                    file_name=f"초과근무_리포트_{target_month}_자동화결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")