import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📦 선입선출(FEFO) 자동 할당 시스템 (실시간 차감 및 재고 병합)")

# 1. 파일 업로드 
uploaded_file = st.file_uploader("작업할 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 2. 데이터 불러오기 
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        # 3. 💡 [사전 작업 1] 특수 조건(2028년, 분리배출)에 맞지 않는 불량 재고 미리 걸러내기
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 4. 💡 [사전 작업 2] 동일 유효일자 재고 환산 병합 및 대표 로트(최대수량) 선정
        if not df_inv_valid.empty:
            # 동일 상품 & 유효일자 중 환산수량이 가장 큰 행의 인덱스 찾기
            idx_max_lot = df_inv_valid.groupby(['상품', '유효일자'])['환산'].idxmax()
            max_lots = df_inv_valid.loc[idx_max_lot, ['상품', '유효일자', '화주LOT']]
            
            # 수량 합치기 및 BOX 입수량 가져오기
            inv_grouped = df_inv_valid.groupby(['상품', '유효일자']).agg({
                '환산': 'sum',
                box_col_name: 'first' if box_col_name else lambda x: 1
            }).reset_index()
            
            # 병합된 수량 옆에 대표 화주LOT 번호 붙이기
            inv_grouped = pd.merge(inv_grouped, max_lots, on=['상품', '유효일자'], how='left')
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자', '환산', '화주LOT', box_col_name])

        # 5. 수주업로드 시트에 결과용 빈 컬럼 만들기
        df_order['할당상태'] = ''
        df_order['부족시_최대가능수량'] = None
        df_order['부족시_LOT'] = ''
        df_order['부족시_유효일자'] = ''

        # 6. 💡 [핵심 로직] 한 줄씩 발주를 읽으며 실시간 재고 차감 처리
        with st.spinner('실시간 재고 차감 및 박스 단위 최적화 중...'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                if pd.isna(mecode) or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                    
                # 현재 남아있는 가용 재고 중 해당 상품 필터링 후 유효일자 빠른 순 정렬
                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)].sort_values(by='유효일자')
                
                if available_inv.empty:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    continue

                # 가장 유통기한 빠른 1순위 재고 픽업
                best_match = available_inv.iloc[0]
                best_idx = best_match.name # 차감할 때 사용할 인덱스
                
                max_qty = best_match['환산']
                lot_str = best_match['화주LOT']
                date_str = best_match['유효일자'].strftime('%Y-%m-%d')
                
                try:
                    box_unit = int(best_match[box_col_name])
                    if box_unit <= 0: box_unit = 1
                except:
                    box_unit = 1
                    
                # 출고 가능 박스 및 수량 계산
                if max_qty >= order_qty:
                    allocated_boxes = int(order_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                else:
                    allocated_boxes = int(max_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = f"부분할당({allocated_boxes}BOX)" if allocated_qty > 0 else "박스단위부족"

                # 계산된 수량 엑셀에 기록
                if allocated_qty > 0:
                    df_order.at[i, '수량'] = allocated_qty
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = state
                    
                    # 💥 [가장 중요한 부분] 남은 재고에서 이번에 나간 수량만큼 실시간 차감!
                    inv_grouped.at[best_idx, '환산'] -= allocated_qty
                    
                    # 100% 충족이 안 되어 쇼트가 난 경우, 부족 알림 열에 기록
                    if allocated_qty < order_qty:
                        df_order.at[i, '부족시_최대가능수량'] = max_qty
                        df_order.at[i, '부족시_LOT'] = lot_str
                        df_order.at[i, '부족시_유효일자'] = date_str
                        
                else:
                    # 박스 수량이 안 나와서 1박스도 못 나가는 경우 (차감 안 함)
                    df_order.at[i, 'LOT'] = '박스단위부족'
                    df_order.at[i, '유효일자'] = '박스단위부족'
                    df_order.at[i, '할당상태'] = '박스단위부족'
                    
                    # 못 나갔을 때 원래 얼마가 있었는지 참고할 수 있게 기록
                    df_order.at[i, '부족시_최대가능수량'] = max_qty
                    df_order.at[i, '부족시_LOT'] = lot_str
                    df_order.at[i, '부족시_유효일자'] = date_str

            # 수량 변경에 따른 발주금액 재계산
            if '발주원가' in df_order.columns:
                df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

        # 7. 결과 화면 및 다운로드
        st.subheader("✅ 할당 완료 결과 미리보기")
        # 새로 추가된 열들이 보이게 컬럼 지정
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(20))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="작업 완료 엑셀 다운로드 📥",
            data=buffer.getvalue(),
            file_name="수주업로드_재고차감_할당완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
