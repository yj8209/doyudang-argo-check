import streamlit as st
import pandas as pd
from PIL import Image
import io
import os
from datetime import datetime
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

# --- 페이지 설정 및 로고 ---
st.set_page_config(page_title="두유당 ARGO 정산 검증 대시보드", layout="wide")

# 로고 이미지 로드
logo_path = "static/KakaoTalk_20260405_223421313.png" 
try:
    image = Image.open(logo_path)
    st.image(image, width=200)
except FileNotFoundError:
    st.warning("로고 이미지를 찾을 수 없습니다. (static 폴더 확인)")

st.title("두유당 ARGO 월별 정산 정밀 검증 시스템")

# --- UI 스타일링 ---
st.markdown("""
<style>
    .reportview-container .main .block-container{
        max-width: 1200px;
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1 { color: #1E3A8A; }
    h2 { color: #2563EB; }
    .stFileUploader { padding-bottom: 2rem; border-bottom: 2px solid #E5E7EB; margin-bottom: 1rem; }
    .stMetric { background-color: #F0F9FF; padding: 15px; border-radius: 10px; border: 1px solid #BAE6FD; }
</style>
""", unsafe_allow_html=True)

# --- 구글 스프레드시트 연결 설정 ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

def get_compensation_data():
    if conn is None:
        return pd.DataFrame(columns=["주문번호", "접수일", "처리일", "스토어", "수량", "판매가", "박스수", "합포장", "상품배상금", "택배배상비", "총 배상청구액"])
    try:
        df = conn.read(ttl="1s")
        if not df.empty:
            num_cols = ["수량", "판매가", "박스수", "상품배상금", "택배배상비", "총 배상청구액"]
            for c in num_cols:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame(columns=["주문번호", "접수일", "처리일", "스토어", "수량", "판매가", "박스수", "합포장", "상품배상금", "택배배상비", "총 배상청구액"])

# --- 데이터 전처리 헬퍼 함수 ---
def load_excel_sheet(excel_file, sheet_name, skip_rows):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=skip_rows, dtype=str)
        if '고객사명' not in df.columns and len(df.columns) > 0:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        return df
    except ValueError:
        st.error(f"엑셀 파일 내에 '{sheet_name}' 시트가 존재하지 않습니다.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame()

# --- 메인 엑셀 파일 업로드 영역 (공통) ---
st.subheader("📁 아르고 정산 엑셀 파일 업로드")
uploaded_excel = st.file_uploader("당월 아르고 정산 원본 엑셀 파일(.xlsx)을 업로드해 주세요.", type=['xlsx', 'xls'])

# --- 탭 구성 ---
tab1, tab2, tab3, tab4 = st.tabs(["📊 시스템 개요", "📦 입고비 검증", "🚚 출고 배송비 검증", "💰 배상금 정산 관리"])

# ==========================================
# TAB 1: 시스템 개요
# ==========================================
with tab1:
    st.header("시스템 개요")
    st.write("""
    두유당의 자체적인 정산 기준을 바탕으로 아르고 풀필먼트 데이터를 정밀하게 검증합니다.
    
    * **📦 입고비 검증:** 당일/일반 입고 유형별 단가 검증
    * **🚚 출고비 검증:** SKU 1~7개 구간별 등급 및 합포장 요율 정밀 대조
    * **💰 배상금 관리:** 구글 시트 연동을 통한 내역 기록 및 월별 정산 합계 확인 (수정/삭제 가능)
    """)

# ==========================================
# TAB 2: 입고비 검증
# ==========================================
with tab2:
    st.header("입고비 정밀 검증")
    if uploaded_excel is None:
        st.info("👆 상단에 엑셀 파일을 먼저 업로드해 주세요.")
    else:
        sku_list = ['하루두유 BLACK', '하루두유 BLACK SWEET', '기타 (직접 입력)']
        selected_sku = st.selectbox("검증할 SKU 선택:", sku_list, key="inbound_sku")
        if selected_sku == '기타 (직접 입력)':
            selected_sku = st.text_input("SKU 이름 입력:", key="inbound_custom_sku")
        actual_qty = st.number_input("실제 입고 수량:", min_value=0, step=1, key="inbound_qty")
        
        if st.button("▶ 입고비 검증 실행", type="primary"):
            df_in = load_excel_sheet(uploaded_excel, sheet_name='입고비', skip_rows=6)
            if not df_in.empty:
                col_idx_qty, col_idx_amt = 10, 9
                for i, col in enumerate(df_in.columns):
                    if '입고 검수비(기본)' in str(col):
                        sub = str(df_in.iloc[0, i]).strip()
                        if sub == '개수': col_idx_qty = i
                        elif sub == '금액': col_idx_amt = i
                
                df_f = df_in[df_in['SKU 이름'] == selected_sku]
                if df_f.empty:
                    st.warning(f"'{selected_sku}' 내역이 없습니다.")
                else:
                    total_q, total_calc, total_billed = 0, 0, 0
                    for _, row in df_f.iterrows():
                        if str(row['고객사명']).strip() == '' or str(row.iloc[col_idx_qty]).strip() == '개수': continue
                        try:
                            q = float(str(row.iloc[col_idx_qty]).strip())
                            b = float(str(row.iloc[col_idx_amt]).strip())
                            total_q +=
