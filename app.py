import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="올리브영 자동 입력 시스템", page_icon="🌿", layout="wide")
st.title("올리브영 자동 입력 시스템")

uploaded_file = st.file_uploader("서식 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 1. 데이터 로드 (기존 헤더 위치 유지)
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 2. 데이터 청소 (이게 안 되면 수량 인식을 못 함)
        # MECODE 문자열 정리
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        
        # 수량 데이터 숫자 강제 변환 (콤마 제거 등)
        for col in ['수량', '발주원가']:
            if col in df_order.columns:
                df_order[col] = pd.to_numeric(df_order[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        df_inv['환산'] = pd.to_numeric(df_inv['환산'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # BOX 입수량 컬럼 찾기 및 숫자 변환
        box_col = next((c for c in df_inv.columns if 'BOX' in c.upper() or '입수량' in c), None)
        if box_col:
            df_inv[box_col] = pd.to_numeric(df_inv[box_col].astype(str).str.replace(',', ''), errors='coerce').fillna(1)

        def process_inventory(df_ord, df_i):
            updated_data = []
            inv_working = df_i.copy()

            for idx, row in df_ord.iterrows():
                mecode = row['MECODE']
                needed_qty = row['수량']
                
                # 기본값 (초기화)
                lot, expiry, final_qty, status = row['LOT'], row['유효일자'], needed_qty, "제외"

                # 처리 대상이 아닌 경우 패스
                if not mecode or mecode == 'NAN' or needed_qty <= 0:
                    updated_data.append([lot, expiry, final_qty, status])
                    continue

                # 🔍 해당 상품의 모든 로트/재고 가져오기
                mask = (inv_working['상품'] == mecode) & (inv_working['환산'] > 0)
                
                # 특수 조건 (기존 유지)
                if mecode == 'ME00621PMM':
                    mask &= (inv_working['유효일자'].dt.year == 2028)
                elif mecode == 'ME90621OC2':
                    mask &= (inv_working['화주LOT'].fillna('').astype(str).str.contains('분리배출'))

                valid_inv = inv_working[mask].sort_values(by=['유효일자', '화주LOT'])

                if valid_inv.empty:
                    updated_data.append(["재고없음", "재고없음", needed_qty, "재고없음"])
                    continue

                # ⭐ [동일 유효일자 묶음 처리]
                first_expiry = valid_inv.iloc[0]['유효일자']
                same_date_group = valid_inv[valid_inv['유효일자'] == first_expiry]
                
                # 해당 유효일자의 총 가용량
                total_available = same_date_group['환산'].sum()
                rep_idx = same_date_group.index[0]
                rep_lot = inv_working.at[rep_idx, '화주LOT']
                rep_exp_str = first_expiry.strftime('%Y-%m-%d')
                
                # 입수량 (없으면 1로 강제 설정하여 에러 방지)
                b_unit = int(inv_working.at[rep_idx, box_col]) if box_col else 1
                if b_unit <= 0: b_unit = 1

                # 🚛 할당 판단
                if total_available >= needed_qty:
                    # 충분함 -> 정상 할당
                    lot, expiry, final_qty, status = rep_lot, rep_exp_str, needed_qty, "정상할당"
                    deduct_qty = needed_qty
                else:
                    # 부족함 -> 해당 날짜의 총량을 박스 단위로 끊어서 할당
                    possible_boxes = int(total_available // b_unit)
                    allocated_qty = possible_boxes * b_unit
                    
                    if allocated_qty <= 0:
                        lot, expiry, final_qty, status = "박스단위부족", "박스단위부족", needed_qty, "박스수량 미달"
                        deduct_qty = 0
                    else:
                        lot, expiry, final_qty, status = rep_lot, rep_exp_str, allocated_qty, f"부분할당({possible_boxes}BOX)"
                        deduct_qty = allocated_qty

                # 🔥 재고 차감 실행
                if deduct_qty > 0:
                    temp_deduct = deduct_qty
                    for s_idx in same_date_group.index:
                        if temp_deduct <= 0: break
                        current_row_qty = inv_working.at[s_idx, '환산']
                        if current_row_qty >= temp_deduct:
                            inv_working.at[s_idx, '환산'] -= temp_deduct
                            temp_deduct = 0
                        else:
                            temp_deduct -= current_row_qty
                            inv_working.at[s_idx, '환산'] = 0
                
                updated_data.append([lot, expiry, final_qty, status])
            
            return updated_data

        if st.button("할당 시작 🚀"):
            with st.spinner('재고 수량을 정밀 검사 중...'):
                results = process_inventory(df_order, df_inv)
                res_df = pd.DataFrame(results, columns=['LOT', '유효일자', '수량', '할당상태'])
                
                df_order['LOT'] = res_df['LOT']
                df_order['유효일자'] = res_df['유효일자']
                df_order['수량'] = res_df['수량']
                df_order['할당상태'] = res_df['할당상태']

                if '발주원가' in df_order.columns:
                    df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

            st.success("할당 프로세스 완료!")
            st.dataframe(df_order[['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']].head(20))

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            st.download_button(label="결과 다운로드 📥", data=buffer.getvalue(), file_name="최종할당결과.xlsx")

    except Exception as e:
        st.error(f"실행 중 오류 발생: {e}")
