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

        # 결과 컬럼 초기화
        new_cols = ['LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        for col in new_cols:
            df_order[col] = ""
            df_order[col] = df_order[col].astype(object) # 범용 타입으로 지정

        # 데이터 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        
        # 계산을 위해 수량은 확실한 숫자로 변환
        df_order['수량'] = to_safe_float(df_order['수량']).astype(float)
        df_inv['환산'] = to_safe_float(df_inv['환산']).astype(float)
        
        # 날짜 처리
        df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
        df_inv['유효일자_보존'] = df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31'))
        df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d').fillna('')

        # 재고 그룹핑
        df_inv['화주LOT'] = df_inv['화주LOT'].astype(str)
        inv_grouped = df_inv.groupby(['상품', '유효일자_보존']).agg({
            '환산': 'sum', 
            '화주LOT': 'first', 
            '유효일자_STR': 'first'
        }).reset_index()

        # 🚀 할당 로직
        for i, row in df_order.iterrows():
            mecode = str(row['MECODE'])
            order_qty = float(row['수량'])
            
            if mecode in ['NAN', '', 'NONE'] or order_qty <= 0:
                df_order.at[i, '할당상태'] = "제외"
                continue
                
            available = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
            if available.empty:
                df_order.at[i, '할당상태'] = "재고없음"
                continue

            best = available.sort_values('유효일자_보존').iloc[0]
            
            # 문자열로 명시적 저장
            df_order.at[i, 'LOT'] = str(best['화주LOT'])
            df_order.at[i, '유효일자'] = str(best['유효일자_STR'])
            df_order.at[i, '할당상태'] = "정상할당" 
            inv_grouped.at[best.name, '환산'] -= order_qty

        # ==========================================
        # 🔥 에러 해결 핵심 구역: 강제 무결성 변환
        # ==========================================
        st.success("✅ 처리가 완료되었습니다!")
        
        view_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']
        existing_cols = [c for c in view_cols if c in df_order.columns]
        
        # 1. 화면에 보여줄 데이터만 자르기
        df_display = df_order[existing_cols].head(100).copy()
        
        # 2. [핵심] Pandas 메타데이터를 파괴하고 순수 문자열 배열로 재구성
        # 이 과정을 거치면 PyArrow가 16.0을 실수로 오해할 가능성이 0%가 됩니다.
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
