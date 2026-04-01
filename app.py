import streamlit as st
import pandas as pd
import io
import re

# 1. 페이지 기본 설정
st.set_page_config(page_title="올리브영 수주업로드 자동 입력 시스템", page_icon="🌿", layout="wide")

# ==========================================
# 🎨 사이드바 디자인
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    st.markdown("---")
    st.header("⚙️ 작업 설정")
    uploaded_file = st.file_uploader("올리브영 발주 엑셀 업로드", type=['xlsx'])
    st.markdown("---")
    st.caption("💡 자동 부분 할당 및 재고 차감 적용")
    st.caption("Developed by Jay")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("Mentholatum : Moving The Heart")

def to_safe_float(series):
    cleaned = series.astype(str).str.replace(r'[^0-9.]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

if uploaded_file:
    try:
        # [핵심 수정!] 파일을 읽을 때 문제되는 컬럼들을 미리 문자열(str)로 지정
        # converters를 사용하면 읽어올 때부터 데이터 타입 충돌을 막을 수 있습니다.
        df_order = pd.read_excel(
            uploaded_file, 
            sheet_name='서식(수주업로드)', 
            header=1,
            converters={'MECODE': str, 'LOT': str, '바코드': str}
        )
        
        df_inv = pd.read_excel(
            uploaded_file, 
            sheet_name='재고', 
            header=2,
            converters={'상품': str, '화주LOT': str, '상품바코드': str}
        )

        # 불필요한 열 제거
        if '잔여일수' in df_order.columns:
            start_idx = list(df_order.columns).index('잔여일수')
            cols_to_drop = df_order.columns[start_idx:]
            df_order = df_order.drop(columns=cols_to_drop)

        # 데이터 정제 (데이터 타입 안정성 확보)
        df_order['MECODE'] = df_order['MECODE'].fillna('').astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        df_inv['상품'] = df_inv['상품'].fillna('').astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        
        # LOT 번호가 숫자로 변하지 않게 처리
        df_inv['화주LOT'] = df_inv['화주LOT'].fillna('').astype(str).str.strip()
        
        # 수량 데이터 처리
        df_order['수량'] = to_safe_float(df_order['수량'])
        df_inv['환산'] = to_safe_float(df_inv['환산'])
        
        # 유효일자 처리
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        df_inv['유효일자_보존'] = df_inv['유효일자'].fillna(pd.Timestamp('2099-12-31'))

        # 입수량 단위 찾기
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None
        product_box_unit = {}
        if box_col_name:
            for mecode, group in df_inv.groupby('상품'):
                box_vals = to_safe_float(group[box_col_name])
                box_vals = box_vals[box_vals > 0]
                if not box_vals.empty:
                    product_box_unit[mecode] = int(box_vals.min())

        # 불량 재고 필터링 (기존 로직 유지)
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 재고 그룹핑
        if not df_inv_valid.empty:
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            agg_dict = {'환산': 'sum', '화주LOT': 'first', '유효일자': 'first'}
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존']).agg(agg_dict).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '환산', '화주LOT', '유효일자'])

        # 결과 컬럼 생성
        df_order['할당상태'] = ''
        df_order['부족시_최대가능수량'] = None
        df_order['부족시_LOT'] = ''
        df_order['부족시_유효일자'] = ''
        if 'LOT' not in df_order.columns: df_order['LOT'] = ''
        if '유효일자' not in df_order.columns: df_order['유효일자'] = ''

        # 할당 로직
        with st.spinner('실시간 재고 차감 중...'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                if not mecode or mecode == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                    
                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
                
                if available_inv.empty:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    continue

                full_match_inv = available_inv[available_inv['환산'] >= order_qty]
                best_match = full_match_inv.sort_values(by='유효일자_보존').iloc[0] if not full_match_inv.empty else available_inv.sort_values(by='유효일자_보존').iloc[0]

                best_idx = best_match.name
                max_qty = best_match['환산']
                lot_str = best_match['화주LOT']
                date_str = best_match['유효일자'].strftime('%Y-%m-%d') if pd.notna(best_match['유효일자']) else '일자없음'
                
                box_unit = product_box_unit.get(mecode, 1)
                allocated_boxes = int(min(order_qty, max_qty) // box_unit)
                allocated_qty = allocated_boxes * box_unit

                if allocated_qty > 0:
                    df_order.at[i, '수량'] = allocated_qty
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                    inv_grouped.at[best_idx, '환산'] -= allocated_qty
                else:
                    df_order.at[i, '할당상태'] = '박스단위부족'
                    df_order.at[i, '부족시_최대가능수량'] = max_qty

            if '발주원가' in df_order.columns:
                df_order['발주원가'] = to_safe_float(df_order['발주원가'])
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

        st.success("✅ 처리가 완료되었습니다!")
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(15), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(label="💾 최종 엑셀 다운로드", data=buffer.getvalue(), file_name="올리브영_자동할당완료.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    except Exception as e:
        st.error(f"오류 발생: {e}")
