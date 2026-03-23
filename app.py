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
        # 2. 데이터 불러오기 (시트 이름 및 헤더 위치 유지)
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 데이터 전처리 (숫자 변환 및 결측치 처리)
        df_order['수량'] = pd.to_numeric(df_order['수량'], errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')

        # '재고' 시트에서 BOX 입수량 컬럼 자동 찾기
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        # 3. 핵심 로직: 순차적 재고 차감 (단일 LOT 할당)
        def process_inventory(df_ord, df_i, box_col):
            # 결과값을 임시 저장할 리스트
            updated_data = []
            
            # 재고 데이터프레임의 복사본을 만들어 실시간으로 수정(차감)함
            inv_working = df_i.copy()

            for _, row in df_ord.iterrows():
                mecode = row['MECODE']
                needed_qty = row['수량']
                
                # 기본값 설정
                lot, expiry, final_qty, status = row['LOT'], row['유효일자'], needed_qty, "제외"

                if pd.isna(mecode) or needed_qty == 0:
                    updated_data.append([lot, expiry, final_qty, status])
                    continue

                # 조건에 맞는 가용 재고 필터링 (환산 > 0)
                mask = (inv_working['상품'] == mecode) & (inv_working['환산'] > 0)
                
                # 특수 조건 적용
                if mecode == 'ME00621PMM':
                    mask &= (inv_working['유효일자'].dt.year == 2028)
                if mecode == 'ME90621OC2':
                    mask &= (inv_working['화주LOT'].fillna('').astype(str).str.contains('분리배출'))

                valid_inv = inv_working[mask].sort_values(by='유효일자', ascending=True)

                if valid_inv.empty:
                    updated_data.append(["재고없음", "재고없음", needed_qty, "재고없음"])
                    continue

                # 선입선출 대상 LOT 선택
                target_idx = valid_inv.index[0]
                available_qty = inv_working.at[target_idx, '환산']
                
                # BOX 입수량 확인
                try:
                    box_unit = int(inv_working.at[target_idx, box_col])
                    if box_unit <= 0: box_unit = 1
                except:
                    box_unit = 1

                # 할당 로직
                if available_qty >= needed_qty:
                    # [전량 할당]
                    lot = inv_working.at[target_idx, '화주LOT']
                    expiry = inv_working.at[target_idx, '유효일자'].strftime('%Y-%m-%d')
                    final_qty = needed_qty
                    status = "정상할당"
                    # 재고 차감
                    inv_working.at[target_idx, '환산'] -= needed_qty
                else:
                    # [부분 할당 - 단일 LOT의 BOX 배수만큼만]
                    possible_boxes = int(available_qty // box_unit)
                    allocated_qty = possible_boxes * box_unit
                    
                    if allocated_qty == 0:
                        lot, expiry, final_qty, status = "박스단위부족", "박스단위부족", needed_qty, "박스수량 미달"
                    else:
                        lot = inv_working.at[target_idx, '화주LOT']
                        expiry = inv_working.at[target_idx, '유효일자'].strftime('%Y-%m-%d')
                        final_qty = allocated_qty
                        status = f"부분할당({possible_boxes}BOX)"
                        # 재고 차감 (가진 만큼 다 씀)
                        inv_working.at[target_idx, '환산'] -= allocated_qty
                
                updated_data.append([lot, expiry, final_qty, status])
            
            return updated_data

        # 4. 버튼 클릭 시 실행
        if st.button("할당 시작 🚀"):
            with st.spinner('재고를 실시간으로 차감하며 매핑 중입니다...'):
                results = process_inventory(df_order, df_inv, box_col_name)
                
                # 결과를 원본 데이터프레임에 업데이트
                res_df = pd.DataFrame(results, columns=['LOT', '유효일자', '수량', '할당상태'])
                df_order['LOT'] = res_df['LOT']
                df_order['유효일자'] = res_df['유효일자']
                df_order['수량'] = res_df['수량']
                df_order['할당상태'] = res_df['할당상태']

                # 발주금액 재계산
                if '발주원가' in df_order.columns:
                    df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                    df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

            # 5. 결과 확인
            st.success("처리가 완료되었습니다!")
            st.subheader("✅ 할당 완료 결과 미리보기")
            st.dataframe(df_order[['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']].head(20))

            # 6. 다운로드 버튼
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
                
            st.download_button(
                label="작업 완료 엑셀 다운로드 📥",
                data=buffer.getvalue(),
                file_name="수주업로드_차감형_할당완료.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"오류 발생: {e}")
        st.info("엑셀 파일의 시트 이름이나 컬럼 위치가 정확한지 확인해 주세요.")
