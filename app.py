import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📦 선입선출(FEFO) 자동 할당 시스템 (재고 병합 완벽 적용)")

uploaded_file = st.file_uploader("작업할 엑셀 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 1. 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 2. 데이터 정제 (공백 및 쉼표 제거)
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip()

        df_order['수량'] = pd.to_numeric(df_order['수량'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # 💡 [핵심 방어 1] 엑셀 셀에 숨은 시간(Time) 단위 무시하고 순수 날짜(Date)로만 통일!
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()

        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        # 3. 불량 재고 걸러내기 (특수 조건)
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 4. 💡 [핵심 방어 2] 동일 유효일자 환산 SUM & 수량 최대 LOT 뽑기
        if not df_inv_valid.empty:
            # 환산(수량)을 기준으로 내림차순 정렬 (가장 수량이 많은 로트가 맨 위로 오게 됨)
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자', '환산'], ascending=[True, True, False])
            
            # 병합 룰(Rule) 설정
            agg_dict = {
                '환산': 'sum',          # 환산 수량은 무조건 전부 합치기 (SUM)
                '화주LOT': 'first'      # 정렬해뒀으므로 첫 번째(수량 제일 큰) 로트 번호를 대표로 가져오기
            }
            if box_col_name:
                agg_dict[box_col_name] = 'first'
                
            # 병합 실행! (빈칸이 있어도 튕기지 않게 dropna=False 지정)
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자'], dropna=False).agg(agg_dict).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자', '환산', '화주LOT'] + ([box_col_name] if box_col_name else []))

        # 5. 수주업로드 시트 초기화
        df_order['할당상태'] = ''
        df_order['부족시_최대가능수량'] = None
        df_order['부족시_LOT'] = ''
        df_order['부족시_유효일자'] = ''

        # 6. 할당 로직 시작
        with st.spinner('실시간 재고 차감 및 박스 단위 최적화 중...'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                # 'nan' 문자열 방어
                if pd.isna(mecode) or str(mecode).lower() == 'nan' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                    
                # 합산된 가상 재고(inv_grouped)에서 해당 상품 찾기
                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)].sort_values(by='유효일자')
                
                if available_inv.empty:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    continue

                best_match = available_inv.iloc[0]
                best_idx = best_match.name
                
                max_qty = best_match['환산']
                lot_str = best_match['화주LOT']
                date_str = best_match['유효일자'].strftime('%Y-%m-%d') if pd.notna(best_match['유효일자']) else '일자없음'
                
                try:
                    box_unit = int(best_match[box_col_name])
                    if box_unit <= 0: box_unit = 1
                except:
                    box_unit = 1
                    
                # 출고 로직 (BOX 단위)
                if max_qty >= order_qty:
                    allocated_boxes = int(order_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                else:
                    allocated_boxes = int(max_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = f"부분할당({allocated_boxes}BOX)" if allocated_qty > 0 else "박스단위부족"

                # 결과 입력 및 실시간 차감
                if allocated_qty > 0:
                    df_order.at[i, '수량'] = allocated_qty
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = state
                    
                    # 💥 [중요] 나간 만큼 합산된 재고에서 차감
                    inv_grouped.at[best_idx, '환산'] -= allocated_qty
                    
                    if allocated_qty < order_qty:
                        df_order.at[i, '부족시_최대가능수량'] = max_qty
                        df_order.at[i, '부족시_LOT'] = lot_str
                        df_order.at[i, '부족시_유효일자'] = date_str
                else:
                    df_order.at[i, 'LOT'] = '박스단위부족'
                    df_order.at[i, '유효일자'] = '박스단위부족'
                    df_order.at[i, '할당상태'] = '박스단위부족'
                    df_order.at[i, '부족시_최대가능수량'] = max_qty
                    df_order.at[i, '부족시_LOT'] = lot_str
                    df_order.at[i, '부족시_유효일자'] = date_str

            if '발주원가' in df_order.columns:
                df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

        # 7. 결과 확인
        st.subheader("✅ 할당 완료 결과 미리보기")
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(20))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="작업 완료 엑셀 다운로드 📥",
            data=buffer.getvalue(),
            file_name="수주업로드_완벽합산완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
