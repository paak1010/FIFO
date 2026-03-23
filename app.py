import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="올리브영 자동 입력 시스템",
    page_icon="🌿",
    layout="wide"
)
st.title("올리브영 자동 입력 시스템")

# 1. 파일 업로드 
uploaded_file = st.file_uploader("서식 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 2. 데이터 불러오기 
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 데이터 전처리
        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # '재고' 시트에서 BOX 입수량 컬럼 찾기
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        with st.spinner('재고 매핑 및 BOX 단위 최적화 중...'):
            # 병합 전 특수 조건(PMM, OC2) 및 일반 상품 필터링
            cond_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year == 2028)
            cond_oc2 = (df_inv['상품'] == 'ME90621OC2') & (df_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출'))
            cond_normal = (~df_inv['상품'].isin(['ME00621PMM', 'ME90621OC2']))
            
            df_inv_filtered = df_inv[cond_pmm | cond_oc2 | cond_normal].copy()

            # 💡 [요청하신 디테일 반영] 환산 수량 기준 내림차순 정렬 
            # -> 이렇게 하면 합칠 때 '화주LOT'의 첫 번째 값(first)이 가장 수량이 큰 로트가 됩니다.
            df_inv_filtered = df_inv_filtered.sort_values(by=['상품', '유효일자', '환산'], ascending=[True, True, False])

            # Groupby로 환산 수량 합치기
            agg_dict = {'환산': 'sum', '화주LOT': 'first'}
            if box_col_name:
                agg_dict[box_col_name] = 'first'
                
            df_inv_agg = df_inv_filtered.groupby(['상품', '유효일자'], as_index=False).agg(agg_dict)
            df_inv_agg = df_inv_agg[df_inv_agg['환산'] > 0] # 남은 재고가 있는 것만 유지

            # 실시간 재고 차감을 위한 리스트 초기화
            allocated_lots, allocated_dates, allocated_qtys, allocated_statuses = [], [], [], []

            # iterrows로 한 줄씩 처리하며 실시간 차감
            for idx, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']

                # 예외 처리
                if pd.isna(mecode) or order_qty == 0:
                    allocated_lots.append(row['LOT'])
                    allocated_dates.append(row['유효일자'])
                    allocated_qtys.append(order_qty)
                    allocated_statuses.append("제외")
                    continue

                # 현재 남은 재고 확인
                valid_inv = df_inv_agg[(df_inv_agg['상품'] == mecode) & (df_inv_agg['환산'] > 0)].sort_values(by='유효일자')

                if valid_inv.empty:
                    allocated_lots.append("재고없음")
                    allocated_dates.append("재고없음")
                    allocated_qtys.append(order_qty)
                    allocated_statuses.append("재고없음")
                    continue

                # 발주 수량을 100% 충족하는 재고 확인
                full_match_inv = valid_inv[valid_inv['환산'] >= order_qty]

                if not full_match_inv.empty:
                    # 정상 할당
                    best_idx = full_match_inv.index[0]
                    best_match = df_inv_agg.loc[best_idx]
                    
                    allocated_lots.append(best_match['화주LOT'])
                    allocated_dates.append(best_match['유효일자'].strftime('%Y-%m-%d'))
                    allocated_qtys.append(order_qty)
                    allocated_statuses.append("정상할당")
                    
                    # 📉 할당된 수량만큼 재고 실시간 차감
                    df_inv_agg.loc[best_idx, '환산'] -= order_qty
                
                else:
                    # BOX 단위 부분 할당 로직
                    best_idx = valid_inv.index[0]
                    best_match = df_inv_agg.loc[best_idx]
                    max_qty = best_match['환산']
                    
                    try:
                        box_unit = int(best_match[box_col_name]) if box_col_name else 1
                        if box_unit <= 0: box_unit = 1
                    except:
                        box_unit = 1
                    
                    possible_boxes = int(max_qty // box_unit)
                    allocated_qty = possible_boxes * box_unit
                    
                    if allocated_qty == 0:
                        allocated_lots.append("박스단위부족")
                        allocated_dates.append("박스단위부족")
                        allocated_qtys.append(order_qty)
                        allocated_statuses.append("박스수량 미달")
                    else:
                        allocated_lots.append(best_match['화주LOT'])
                        allocated_dates.append(best_match['유효일자'].strftime('%Y-%m-%d'))
                        allocated_qtys.append(allocated_qty)
                        allocated_statuses.append(f"부분할당({possible_boxes}BOX)")
                        
                        # 📉 할당된 수량만큼 재고 실시간 차감
                        df_inv_agg.loc[best_idx, '환산'] -= allocated_qty

            # 결과를 DataFrame에 반영
            df_order['LOT'] = allocated_lots
            df_order['유효일자'] = allocated_dates
            df_order['수량'] = allocated_qtys
            df_order['할당상태'] = allocated_statuses

            # 발주금액 재계산
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
