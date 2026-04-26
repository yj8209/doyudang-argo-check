import streamlit as st
import pandas as pd
from PIL import Image
import io
import os
from datetime import datetime
from openpyxl.utils import get_column_letter

# --- 페이지 설정 및 로고 ---
st.set_page_config(page_title="두유당 ARGO 정산 검증 대시보드", layout="wide")

# 로고 이미지 로드
logo_path = "static/KakaoTalk_20260405_223421313.png" 
try:
    image = Image.open(logo_path)
    st.image(image, width=200)
except FileNotFoundError:
    st.warning(f"로고 이미지를 찾을 수 없습니다. 경로를 확인해 주세요: {logo_path}")

st.title("두유당 ARGO 월별 정산 정밀 검증 시스템")

# --- [영구 저장소 설정] CSV 파일 경로 ---
DB_FILE = "compensation_data.csv"

# 데이터 불러오기 함수
def load_comp_data():
    if os.path.exists(DB_FILE):
        return pd.read_csv(DB_FILE, dtype={'주문번호': str})
    else:
        return pd.DataFrame(columns=[
            "주문번호", "접수일", "처리일", "스토어", "수량", 
            "판매가", "박스수", "합포장", "상품배상금", "택배배상비", "총 배상청구액"
        ])

# 데이터 저장 함수
def save_comp_data(df):
    df.to_csv(DB_FILE, index=False, encoding='utf-8-sig')

# 세션 상태에 데이터 로드
if 'compensation_df' not in st.session_state:
    st.session_state.compensation_df = load_comp_data()

# --- 데이터 전처리 헬퍼 함수 (엑셀 로드용) ---
def load_excel_sheet(excel_file, sheet_name, skip_rows):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=skip_rows, dtype=str)
        if '고객사명' not in df.columns and len(df.columns) > 0:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        return df
    except Exception:
        return pd.DataFrame()

# --- 탭 구성 ---
tab1, tab2, tab3, tab4 = st.tabs(["📊 시스템 개요", "📦 입고비 검증", "🚚 출고 배송비 검증", "💰 배상금 정산 관리"])

# (Tab 1, 2, 3 로직 유지...)

with tab4:
    st.header("배상 청구 내역 기록 및 영구 보관")
    
    with st.form("compensation_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            order_no = st.text_input("주문번호 (연관번호)")
            store_type = st.selectbox("스토어 구분", ["네이버스마트스토어", "카페24"])
            inquiry_qty = st.number_input("문의요청 수량", min_value=1, step=1)
        with col2:
            reg_date = st.date_input("접수일자", datetime.now())
            proc_date = st.date_input("처리일자", datetime.now())
            unit_price = st.number_input("개별 판매가격 (원)", min_value=0, step=1)
        with col3:
            box_qty = st.number_input("배송 박스 수량", min_value=1, step=1)
            bundle_type = st.selectbox("합포장 여부", ["없음", "동종", "이종"])
            is_island = st.checkbox("도서산간 여부 (+3,000원)")
        
        submit_btn = st.form_submit_button("배상 내역 영구 저장")
    
    if submit_btn:
        if not order_no:
            st.error("주문번호를 입력해 주세요.")
        else:
            # 계산 로직
            product_comp = unit_price * inquiry_qty
            shipping_unit = 4800 if store_type == "네이버스마트스토어" else 4400
            shipping_comp = (shipping_unit * box_qty) + (3000 if is_island else 0)
            total_comp = product_comp + shipping_comp
            
            new_data = {
                "주문번호": order_no,
                "접수일": reg_date.strftime("%Y-%m-%d"),
                "처리일": proc_date.strftime("%Y-%m-%d"),
                "스토어": store_type,
                "수량": inquiry_qty,
                "판매가": unit_price,
                "박스수": box_qty,
                "합포장": bundle_type,
                "상품배상금": product_comp,
                "택배배상비": shipping_comp,
                "총 배상청구액": total_comp
            }
            
            # 데이터프레임 업데이트 및 파일 저장
            st.session_state.compensation_df = pd.concat([st.session_state.compensation_df, pd.DataFrame([new_data])], ignore_index=True)
            save_comp_data(st.session_state.compensation_df)
            st.success(f"주문번호 {order_no} 내역이 안전하게 파일에 저장되었습니다.")

    # 기록된 내역 표시 및 관리
    if not st.session_state.compensation_df.empty:
        df_display = st.session_state.compensation_df.copy()
        
        # 합계 계산
        total_sum = df_display["총 배상청구액"].sum()
        st.metric("🚨 이번 달 누적 배상 청구 합계", f"{total_sum:,.0f} 원")
        
        # 금액 콤마 포맷팅
        styled_comp = df_display.style.format({
            "판매가": "{:,.0f}", "상품배상금": "{:,.0f}", 
            "택배배상비": "{:,.0f}", "총 배상청구액": "{:,.0f}"
        })
        st.dataframe(styled_comp, use_container_width=True)
        
        # 데이터 삭제 기능 (선택 항목 삭제 등은 복잡하므로 여기서는 전체 초기화 버튼 제공)
        if st.button("전체 기록 초기화 (주의: 파일이 삭제됩니다)"):
            if os.path.exists(DB_FILE):
                os.remove(DB_FILE)
                st.session_state.compensation_df = load_comp_data()
                st.rerun()
        
        # 엑셀 다운로드
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_display.to_excel(writer, index=False, sheet_name='배상금기록')
            worksheet = writer.sheets['배상금기록']
            for i, col in enumerate(df_display.columns):
                max_len = max(df_display[col].astype(str).map(len).max(), len(str(col)))
                worksheet.column_dimensions[get_column_letter(i + 1)].width = (max_len * 1.5) + 5
                
        st.download_button(
            label="📥 전체 배상 내역 엑셀 다운로드",
            data=excel_buffer.getvalue(),
            file_name=f"두유당_배상금관리_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        st.info("현재 저장된 배상 내역이 없습니다.")
