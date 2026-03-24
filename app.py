import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📦 선입선출(FEFO) 자동 할당 시스템 (BOX 단위 부분할당 지원)")

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

        # '재고' 시트에서 BOX 입수량 컬럼 찾기 (예: '합계 : 입수량(BOX)')
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        # 3. 핵심 로직: BOX 단위 부분 할당
        def assign_inventory(row):
            mecode = row['MECODE']
            order_qty = row['수량']
            
            if pd.isna(mecode) or order_qty == 0:
                return pd.Series([row['LOT'], row['유효일자'], order_qty, "제외"])

            # 조건 1: MECODE 일치 및 특수 조건
            valid_inv = df_inv[df_inv['상품'] == mecode]
            
            if mecode == 'ME00621PMM':
                valid_inv = valid_inv[valid_inv['유효일자'].dt.year == 2028]
            if mecode == 'ME90621OC2':
                valid_inv = valid_inv[valid_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출')]
                
            valid_inv = valid_inv[valid_inv['환산'] > 0]
            if valid_inv.empty:
                return pd.Series(["재고없음", "재고없음", order_qty, "재고없음"])

            # 조건 2: 발주 수량을 100% 충족하는 재고 확인
            full_match_inv = valid_inv[valid_inv['환산'] >= order_qty]
            
            if not full_match_inv.empty:
                best_match = full_match_inv.sort_values(by='유효일자', ascending=True).iloc[0]
                return pd.Series([best_match['화주LOT'], best_match['유효일자'].strftime('%Y-%m-%d'), order_qty, "정상할당"])
            
            else:
                # 💡 [BOX 단위 부분 할당 로직] 
                best_match = valid_inv.sort_values(by='유효일자', ascending=True).iloc[0]
                max_qty = best_match['환산']
                
                # BOX 입수량 가져오기 (숫자가 아니거나 없으면 기본값 1 처리)
                try:
                    box_unit = int(best_match[box_col_name])
                    if box_unit <= 0: box_unit = 1
                except:
                    box_unit = 1
                
                # 📦 출고 가능한 최대 박스 수 계산 (예: 950 // 200 = 4박스)
                possible_boxes = int(max_qty // box_unit)
                
                # 📦 최종 출고 수량 (예: 4박스 * 200 = 800개)
                allocated_qty = possible_boxes * box_unit
                
                if allocated_qty == 0:
                    # 1박스조차 안 되는 자투리 재고만 남은 경우
                    return pd.Series(["박스단위부족", "박스단위부족", order_qty, "박스수량 미달"])
                else:
                    return pd.Series([
                        best_match['화주LOT'], 
                        best_match['유효일자'].strftime('%Y-%m-%d'), 
                        allocated_qty,  # 낱개를 버린 BOX 배수 수량으로 업데이트!
                        f"부분할당({possible_boxes}BOX)"
                    ])

        # 4. 함수 적용
        with st.spinner('재고 매핑 및 BOX 단위 최적화 중...'):
            df_order[['LOT', '유효일자', '수량', '할당상태']] = df_order.apply(assign_inventory, axis=1)

            # 수량 변경에 따른 발주금액 재계산
            if '발주원가' in df_order.columns:
                df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

        # 5. 결과 확인
        st.subheader("✅ 할당 완료 결과 미리보기")
        st.dataframe(df_order[['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']].head(15))

        # 6. 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="작업 완료 엑셀 다운로드 📥",
            data=buffer.getvalue(),
            file_name="수주업로드_BOX단위_부분할당완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
