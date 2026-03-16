
import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📦 선입선출(FEFO) 자동 할당 시스템")

# 1. 파일 업로드 (엑셀 파일 하나에 여러 시트가 있다고 가정)
uploaded_file = st.file_uploader("작업할 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 2. 데이터 불러오기 
        # 주의: 실제 엑셀 파일의 표 헤더(열 이름) 위치에 맞춰 header를 지정해야 합니다.
        # 예시 데이터 기준: 수주업로드 시트는 2번째 줄, 재고 시트는 3번째 줄이 헤더
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 데이터 전처리: 계산을 위해 숫자 및 날짜 형식으로 변환
        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # 3. 핵심 로직: LOT 및 유효일자 매핑 함수
        def assign_inventory(row):
            barcode = row['바코드']
            order_qty = row['수량']
            
            # 발주 수량이 0이거나 바코드가 없으면 건너뜀
            if pd.isna(barcode) or order_qty == 0:
                return pd.Series([row['LOT'], row['유효일자']])

            # 조건 1: 같은 바코드이면서, 재고 '환산' 수량이 수주 '수량'보다 큰 재고만 필터링
            # (만약 크거나 같다(>=) 조건이 필요하면 '>=' 로 수정하세요)
            valid_inv = df_inv[(df_inv['상품바코드'] == barcode) & (df_inv['환산'] > order_qty)]
            
            if not valid_inv.empty:
                # 조건 2: 유효일자가 가장 빠른 순으로 정렬 (FEFO)
                valid_inv = valid_inv.sort_values(by='유효일자', ascending=True)
                
                # 가장 첫 번째(유효일자가 제일 빠른) 데이터 선택
                best_match = valid_inv.iloc[0]
                
                # 추출한 화주LOT와 유효일자를 반환
                # 유효일자는 엑셀 형식에 맞게 문자열(YYYY-MM-DD)로 변환
                return pd.Series([best_match['화주LOT'], best_match['유효일자'].strftime('%Y-%m-%d')])
            else:
                # 조건을 만족하는 재고가 없으면 "재고부족" 표시
                return pd.Series(["재고부족", "재고부족"])

        # 4. 함수 적용하여 수주업로드 데이터프레임 업데이트
        with st.spinner('재고 매핑 중...'):
            df_order[['LOT', '유효일자']] = df_order.apply(assign_inventory, axis=1)

        # 5. 결과 확인
        st.subheader("✅ 할당 완료 결과 미리보기")
        st.dataframe(df_order[['바코드', '상품명', '수량', 'LOT', '유효일자']].head(15))

        # 6. 엑셀 파일로 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="작업 완료 엑셀 다운로드 📥",
            data=buffer.getvalue(),
            file_name="수주업로드_LOT할당완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
