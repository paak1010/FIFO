import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📦 선입선출(FEFO) 자동 할당 시스템 (특수조건 포함)")

# 1. 파일 업로드 
uploaded_file = st.file_uploader("작업할 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 2. 데이터 불러오기 
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 데이터 전처리
        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # 3. 핵심 로직: LOT 및 유효일자 매핑 함수
        def assign_inventory(row):
            mecode = row['MECODE']
            order_qty = row['수량']
            
            # 발주 수량이 0이거나 MECODE가 없으면 건너뜀
            if pd.isna(mecode) or order_qty == 0:
                return pd.Series([row['LOT'], row['유효일자']])

            # 기본 조건: 재고 상품코드 일치 & 환산수량 > 발주수량
            valid_inv = df_inv[(df_inv['상품'] == mecode) & (df_inv['환산'] > order_qty)]
            
            # 💡 [추가된 예외 로직 1] ME00621PMM: 유효일자가 2028년인 재고만
            if mecode == 'ME00621PMM':
                valid_inv = valid_inv[valid_inv['유효일자'].dt.year == 2028]
                
            # 💡 [추가된 예외 로직 2] ME90621OC2: LOT에 "분리배출"이 포함된 재고만
            if mecode == 'ME90621OC2':
                # 결측치(NaN)가 있을 수 있으므로 빈 문자열로 채운 뒤 검색
                valid_inv = valid_inv[valid_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출')]

            # 가용 재고가 있는지 확인
            if not valid_inv.empty:
                # 조건 2: 유효일자가 가장 빠른 순으로 정렬 (FEFO)
                valid_inv = valid_inv.sort_values(by='유효일자', ascending=True)
                
                # 가장 첫 번째(유효일자가 제일 빠른) 데이터 선택
                best_match = valid_inv.iloc[0]
                
                return pd.Series([best_match['화주LOT'], best_match['유효일자'].strftime('%Y-%m-%d')])
            else:
                return pd.Series(["조건불충분/재고부족", "재고부족"])

        # 4. 함수 적용
        with st.spinner('재고 매핑 및 특수 조건 검사 중...'):
            df_order[['LOT', '유효일자']] = df_order.apply(assign_inventory, axis=1)

        # 5. 결과 확인
        st.subheader("✅ 할당 완료 결과 미리보기")
        st.dataframe(df_order[['MECODE', '상품명', '수량', 'LOT', '유효일자']].head(15))

        # 6. 엑셀 파일로 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="작업 완료 엑셀 다운로드 📥",
            data=buffer.getvalue(),
            file_name="수주업로드_LOT할당완료(특수조건).xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
