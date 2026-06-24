import streamlit as st
import pandas as pd
from datetime import datetime

# --- 페이지 설정 ---
st.set_page_config(page_title="올리브영 재고 대시보드", layout="wide")
st.title("📦 올리브영 재고 대시보드")

# --- 1. 파일 업로드 ---
uploaded_file = st.file_uploader("재고 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # --- 2. 데이터 로드 및 ffill 처리 (피벗/병합 셀 빈칸 채우기) ---
        df = pd.read_excel(uploaded_file)
        df = df.ffill()

        # --- 3. 핵심 전처리 (공백 제거 및 대소문자 통일) ---
        # 상품코드가 엑셀 내에 다른 이름(예: 'Item Code')으로 되어있다면 컬럼명을 수정해주세요.
        if '상품코드' in df.columns:
            # 앞뒤 공백 완전 제거 후 대문자로 강제 변환
            df['상품코드'] = df['상품코드'].astype(str).str.strip().str.upper()
        if '상품명' in df.columns:
            df['상품명'] = df['상품명'].astype(str).str.strip()

        # --- 4. 🚨 긴급 디버깅: ME90621GGF 생존 확인 ---
        st.markdown("---")
        st.subheader("🔍 [디버깅] ME90621GGF 추적 모니터")
        if '상품코드' in df.columns:
            debug_df = df[df['상품코드'] == 'ME90621GGF']
            if not debug_df.empty:
                st.success("✅ 원본 데이터 및 전처리 직후에 'ME90621GGF'가 정상적으로 인식되었습니다!")
                st.dataframe(debug_df)
            else:
                st.error("❌ 전처리 직후 데이터에서 'ME90621GGF'를 찾을 수 없습니다. 원본 엑셀 파일에 해당 코드가 누락되었거나 컬럼명이 다를 수 있습니다.")
        
        # --- 5. 유효일자 필터링 로직 (에러 방지 적용) ---
        if '유효일자' in df.columns:
            # 날짜 형식으로 변환 불가능한 값은 에러를 내지 않고 NaT(빈 값)으로 처리
            df['유효일자_DT'] = pd.to_datetime(df['유효일자'], errors='coerce')
            
            # 오늘 날짜 기준
            today = pd.to_datetime('today')
            
            # 기한이 남았거나(오늘 이후), 날짜 파싱이 안 된 데이터도 일단 살려둠 (오류로 인한 누락 방지)
            df_filtered = df[(df['유효일자_DT'] >= today) | (df['유효일자_DT'].isna())]
            
            # 화면 표시를 위해 임시 계산용 컬럼은 제외
            df_display = df_filtered.drop(columns=['유효일자_DT'])
        else:
            df_display = df.copy()

        st.markdown("---")

        # --- 6. 하이브리드 매칭 검색 ---
        st.subheader("🔎 재고 검색 (코드 & 상품명)")
        search_query = st.text_input("찾으시는 상품코드나 상품명 키워드를 입력하세요.")

        if search_query:
            search_query = search_query.strip().upper()
            
            # 코드 일치 여부 확인
            mask_code = df_display['상품코드'].str.contains(search_query, na=False) if '상품코드' in df_display.columns else False
            # 상품명 포함 여부 확인 (대소문자 무시)
            mask_name = df_display['상품명'].str.upper().str.contains(search_query, na=False) if '상품명' in df_display.columns else False
            
            # 둘 중 하나라도 맞으면 화면에 표시
            result_df = df_display[mask_code | mask_name]
            
            st.write(f"검색 결과: **{len(result_df)}** 건")
            st.dataframe(result_df)
        else:
            # 검색어가 없을 때는 전체 데이터 표시
            st.subheader("📊 전체 재고 리스트")
            st.dataframe(df_display)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
