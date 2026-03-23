import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="올리브영 자동 입력 시스템", page_icon="🌿", layout="wide")
st.title("올리브영 자동 입력 시스템")

uploaded_file = st.file_uploader("서식 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    try:
        # 1. 데이터 로드
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 2. 전처리 (공백 제거 및 숫자 변환 필수)
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip()
        
        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # BOX 입수량 컬럼 찾기
        box_col = next((c for c in df_inv.columns if 'BOX' in c.upper() or '입수량' in c), None)

        def process_inventory(df_ord, df_i):
            updated_data = []
            # 실시간 차감용 데이터 (복사본)
            inv_working = df_i.copy()

            for _, row in df_ord.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                # 기본값
                lot, expiry, final_qty, status = row['LOT'], row['유효일자'], order_qty, "제외"

                if pd.isna(mecode) or mecode == 'nan' or order_qty <= 0:
                    updated_data.append([lot, expiry, final_qty, status])
                    continue

                # 🔍 가용 재고 필터링
                mask = (inv_working['상품'] == mecode) & (inv_working['환산'] > 0)
                
                # 특수 조건 적용
                if mecode == 'ME00621PMM':
                    mask &= (inv_working['유효일자'].dt.year == 2028)
                elif mecode == 'ME90621OC2':
                    mask &= (inv_working['화주LOT'].fillna('').astype(str).str.contains('분리배출'))

                # 선입선출 정렬 (유효일자가 같으면 로트번호 순)
                valid_inv = inv_working[mask].sort_values(by=['유효일자', '화주LOT'])

                if valid_inv.empty:
                    updated_data.append(["재고없음", "재고없음", order_qty, "재고없음"])
                    continue

                # [핵심] 유효일자가 같은 로트들은 하나로 간주하여 체크
                first_expiry = valid_inv.iloc[0]['유효일자']
                # 같은 날짜를 가진 모든 행의 인덱스와 합계 수량 추출
                same_date_inv = valid_inv[valid_inv['유효일자'] == first_expiry]
                total_qty_on_date = same_date_inv['환산'].sum()
                
                # 대표 로트 정보
                rep_idx = same_date_inv.index[0]
                rep_lot = inv_working.at[rep_idx, '화주LOT']
                rep_expiry_str = first_expiry.strftime('%Y-%m-%d')
                
                # 박스 입수량 확인
                try:
                    b_unit = int(inv_working.at[rep_idx, box_col]) if box_col else 1
                    if b_unit <= 0: b_unit = 1
                except:
                    b_unit = 1

                # 🚛 할당 로직
                if total_qty_on_date >= order_qty:
                    # 해당 날짜 로트들을 합치면 충분할 때 -> 대표 로트로 전량 할당
                    lot, expiry, final_qty, status = rep_lot, rep_expiry_str, order_qty, "정상할당"
                    
                    # 🔥 재고 차감 (여러 행에 걸쳐 있을 수 있으므로 순차 차감)
                    remaining_to_deduct = order_qty
                    for s_idx in same_date_inv.index:
                        if remaining_to_deduct <= 0: break
                        row_qty = inv_working.at[s_idx, '환산']
                        if row_qty >= remaining_to_deduct:
                            inv_working.at[s_idx, '환산'] -= remaining_to_deduct
                            remaining_to_deduct = 0
                        else:
                            remaining_to_deduct -= row_qty
                            inv_working.at[s_idx, '환산'] = 0
                else:
                    # 해당 날짜 다 긁어도 부족할 때 -> BOX 단위 부분 할당
                    possible_boxes = int(total_qty_on_date // b_unit)
                    allocated_qty = possible_boxes * b_unit
                    
                    if allocated_qty <= 0:
                        lot, expiry, final_qty, status = "박스단위부족", "박스단위부족", order_qty, "박스수량 미달"
                    else:
                        lot, expiry, final_qty, status = rep_lot, rep_expiry_str, allocated_qty, f"부분할당({possible_boxes}BOX)"
                        # 해당 날짜 재고 소진 (박스 단위만큼만 차감)
                        remaining_to_deduct = allocated_qty
                        for s_idx in same_date_inv.index:
                            if remaining_to_deduct <= 0: break
                            row_qty = inv_working.at[s_idx, '환산']
                            if row_qty >= remaining_to_deduct:
                                inv_working.at[s_idx, '환산'] -= remaining_to_deduct
                                remaining_to_deduct = 0
                            else:
                                remaining_to_deduct -= row_qty
                                inv_working.at[s_idx, '환산'] = 0
                
                updated_data.append([lot, expiry, final_qty, status])
            return updated_data

        if st.button("할당 시작 🚀"):
            with st.spinner('재고 확인 중...'):
                results = process_inventory(df_order, df_inv)
                res_df = pd.DataFrame(results, columns=['LOT', '유효일자', '수량', '할당상태'])
                
                df_order['LOT'] = res_df['LOT']
                df_order['유효일자'] = res_df['유효일자']
                df_order['수량'] = res_df['수량']
                df_order['할당상태'] = res_df['할당상태']

                if '발주원가' in df_order.columns:
                    df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                    df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

            st.success("완료!")
            st.dataframe(df_order[['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']].head(20))

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            st.download_button(label="엑셀 다운로드 📥", data=buffer.getvalue(), file_name="결과.xlsx")

    except Exception as e:
        st.error(f"오류: {e}")
