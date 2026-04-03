import streamlit as st
import pandas as pd
import io

# 1. 페이지 기본 설정
st.set_page_config(page_title="올리브영 수주업로드 자동 입력 시스템", page_icon="🌿", layout="wide")

st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("Mentholatum : Moving The Heart")

def to_safe_float(series):
    """숫자 변환 시 에러 방지"""
    return pd.to_numeric(series.astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)

if uploaded_file := st.sidebar.file_uploader("올리브영 발주 엑셀 업로드", type=['xlsx']):
    try:
        # 데이터 로드
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1).copy()
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2).copy()

        # 결과 컬럼 초기화 (범용 타입)
        new_cols = ['LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        for col in new_cols:
            df_order[col] = ""
            df_order[col] = df_order[col].astype(object)

        # 데이터 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        
        df_order['수량'] = to_safe_float(df_order['수량']).astype(float)
        df_inv['환산'] = to_safe_float(df_inv['환산']).astype(float)
        
        # 날짜 처리
        df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
        df_inv['유효일자_보존'] = df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31'))
        df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d').fillna('')

        # [복구됨] 박스 입수량 계산 로직
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None
        product_box_unit = {}
        if box_col_name:
            for mecode, group in df_inv.groupby('상품'):
                box_vals = to_safe_float(group[box_col_name])
                box_vals = box_vals[box_vals > 0]
                if not box_vals.empty:
                    product_box_unit[mecode] = int(box_vals.min())

        # 불량 재고 필터링 (최초 코드 로직 복원)
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자_DT'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (df_inv['화주LOT'].astype(str).str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 재고 그룹핑
        df_inv_valid['화주LOT'] = df_inv_valid['화주LOT'].astype(str)
        if not df_inv_valid.empty:
            inv_grouped = df_inv_valid.groupby(['상품', '유효일자_보존']).agg({
                '환산': 'sum', 
                '화주LOT': 'first', 
                '유효일자_STR': 'first'
            }).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '환산', '화주LOT', '유효일자_STR'])

        # 🚀 할당 로직 [완벽 복구 구간]
        for i, row in df_order.iterrows():
            mecode = str(row['MECODE'])
            order_qty = float(row['수량'])
            
            if mecode in ['NAN', '', 'NONE'] or order_qty <= 0:
                df_order.at[i, '할당상태'] = "제외"
                continue
                
            available = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
            if available.empty:
                df_order.at[i, 'LOT'], df_order.at[i, '유효일자'], df_order.at[i, '할당상태'] = '재고없음', '재고없음', '재고없음'
                continue

            # 1. 주문 수량을 온전히 감당할 수 있는 LOT가 있는지 먼저 확인
            full_match_inv = available[available['환산'] >= order_qty]
            if not full_match_inv.empty:
                best_match = full_match_inv.sort_values('유효일자_보존').iloc[0]
            else:
                # 2. 없다면, 남은 것 중 유효일자가 가장 빠른 것 선택
                best_match = available.sort_values('유효일자_보존').iloc[0]

            best_idx = best_match.name
            max_qty = float(best_match['환산'])
            lot_str = str(best_match['화주LOT'])
            date_str = str(best_match['유효일자_STR'])
            
            box_unit = product_box_unit.get(mecode, 1)
            
            # [핵심 복구] 주문 수량과 현재 해당 LOT의 최대 가능 수량 중 작은 값 선택
            potential_qty = min(order_qty, max_qty)
            
            # 박스 단위로 계산
            allocated_boxes = int(potential_qty // box_unit)
            allocated_qty = float(allocated_boxes * box_unit)

            if allocated_qty > 0:
                df_order.at[i, '수량'] = allocated_qty
                df_order.at[i, 'LOT'] = lot_str
                df_order.at[i, '유효일자'] = date_str
                
                if allocated_qty == order_qty:
                    df_order.at[i, '할당상태'] = "정상할당"
                else:
                    df_order.at[i, '할당상태'] = f"부분할당({allocated_boxes}BOX)"
                
                # 할당된 수량만큼만 정확히 차감
                inv_grouped.at[best_idx, '환산'] -= allocated_qty
            else:
                df_order.at[i, '할당상태'] = '박스단위부족'
                df_order.at[i, '부족시_최대가능수량'] = max_qty
                df_order.at[i, '부족시_LOT'] = lot_str
                df_order.at[i, '부족시_유효일자'] = date_str

        # ==========================================
        # 🔥 에러 방지 구역: 화면 출력 (변경 없음)
        # ==========================================
        st.success("✅ 처리가 완료되었습니다!")
        
        view_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']
        existing_cols = [c for c in view_cols if c in df_order.columns]
        
        df_display = df_order[existing_cols].head(100).copy()
        
        # Pandas 메타데이터 파괴 및 순수 문자열 배열로 재구성
        df_safe_display = pd.DataFrame(
            df_display.to_numpy().astype(str), 
            columns=df_display.columns
        )
        
        st.subheader("📊 작업 결과 미리보기")
        st.dataframe(df_safe_display, use_container_width=True, hide_index=True)

        # 💾 다운로드 처리
        buffer = io.BytesIO()
        df_order.to_excel(buffer, index=False)
        st.download_button("💾 엑셀 다운로드", buffer.getvalue(), "결과.xlsx", "application/vnd.ms-excel")

    except Exception as e:
        st.error(f"오류 발생: {e}")
