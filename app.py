import streamlit as st
import pandas as pd
import math
from datetime import datetime

st.set_page_config(page_title="재고 할당 대시보드", layout="wide")
st.title("📦 재고 할당 및 대시보드 (LOT별 관리)")

uploaded_file = st.file_uploader("재고 엑셀 파일을 업로드하세요 (.xlsx, .xls)", type=['xlsx', 'xls'])

# --- 엑셀 열 순서 변경 방어 (자동 키워드 감지) ---
def find_column(df, keywords):
    for col in df.columns:
        for keyword in keywords:
            if keyword in str(col).replace(" ", ""):
                return col
    return None

if uploaded_file:
    # 1. 데이터 로드 및 피벗/병합 빈 셀 채우기
    df = pd.read_excel(uploaded_file)
    df = df.ffill()

    # 열 자동 매칭
    col_code = find_column(df, ['상품코드', '제품코드', 'ItemCode'])
    col_name = find_column(df, ['상품명', '제품명'])
    col_date = find_column(df, ['유효일자', '유통기한'])
    col_lot = find_column(df, ['LOT', '로트'])
    col_qty = find_column(df, ['수량', '재고'])
    col_status = find_column(df, ['상태', '구분']) 
    
    if col_code:
        # ✨ 핵심 수정: ME90621GGF 공백 찌꺼기 및 대소문자 인식 문제 완벽 해결
        df[col_code] = df[col_code].astype(str).str.strip().str.upper()
        
    if col_qty:
        # 오류 방지를 위한 데이터 타입 정제
        df[col_qty] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0)

    # 2. 회송예정, 불량 재고 등 가용 불가능한 항목 필터링
    if col_status:
        df = df[~df[col_status].str.contains('회송|불량', na=False, regex=True)]

    # 3. 유효기간 548일 이하 재고 제외 로직
    if col_date:
        df['유효일자_DT'] = pd.to_datetime(df[col_date], errors='coerce')
        today = pd.to_datetime('today')
        df['남은일수'] = (df['유효일자_DT'] - today).dt.days
        
        # 548일 초과거나, 날짜 데이터가 비어있는 경우만 가용 재고로 인정
        df = df[(df['남은일수'] > 548) | (df['유효일자_DT'].isna())]

    # 4. LOT 및 유효일자별 재고 합산 처리
    if col_code and col_lot and col_qty:
        groupby_cols = [col_code]
        if col_name: groupby_cols.append(col_name)
        groupby_cols.append(col_lot)
        if col_date: groupby_cols.append(col_date)
        
        # 제품 및 로트별 합산
        df_summary = df.groupby(groupby_cols, as_index=False)[col_qty].sum()
        
        st.subheader("📊 LOT별 가용 재고 현황 (548일 초과)")
        st.dataframe(df_summary)

        st.markdown("---")
        
        # 5. 실시간 재고 차감 및 부분 할당 (박스 단위 계산 복구)
        st.subheader("📦 실시간 할당 및 박스 시뮬레이션")
        if 'inventory' not in st.session_state:
            st.session_state.inventory = df_summary.copy()

        col1, col2, col3 = st.columns(3)
        with col1:
            box_capacity = st.number_input("1박스당 입수량", min_value=1, value=10, step=1)
        with col2:
            search_code = st.text_input("할당할 상품코드 입력", placeholder="예: ME90621GGF")
        with col3:
            order_qty = st.number_input("할당 요청 수량", min_value=1, step=1)
        
        if st.button("할당 실행"):
            search_code = search_code.strip().upper()
            target_items = st.session_state.inventory[st.session_state.inventory[col_code] == search_code]
            
            if target_items.empty:
                st.error(f"[{search_code}] 가용 재고가 부족하거나 유효기간(548일 이하) 기준 미달입니다.")
            else:
                remaining_order = order_qty
                allocated_log = []
                
                for idx, row in target_items.iterrows():
                    if remaining_order <= 0:
                        break
                        
                    available_qty = row[col_qty]
                    if available_qty > 0:
                        # 가용 재고 내에서만 부분 할당 실행
                        allocate_qty = min(remaining_order, available_qty)
                        st.session_state.inventory.at[idx, col_qty] -= allocate_qty
                        remaining_order -= allocate_qty
                        
                        # 박스 단위 계산
                        boxes = math.floor(allocate_qty / box_capacity)
                        ea = allocate_qty % box_capacity
                        
                        allocated_log.append({
                            '상품코드': row[col_code],
                            'LOT': row[col_lot],
                            '할당수량': allocate_qty,
                            '포장 단위': f"{boxes} Box + {ea} EA"
                        })
                
                if allocated_log:
                    st.success(f"✅ 할당 완료! (미할당 잔여 주문량: {remaining_order})")
                    st.table(pd.DataFrame(allocated_log))
                    st.write("실시간 반영 후 남은 재고:")
                    st.dataframe(st.session_state.inventory[st.session_state.inventory[col_code] == search_code])
    else:
        st.warning("⚠️ 필수 열(상품코드, LOT, 수량)을 인식할 수 없습니다. 엑셀 컬럼명을 다시 확인해주세요.")
