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
    st.warning("로고 이미지를 찾을 수 없습니다. (static 폴더를 확인해주세요)")

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
</style>
""", unsafe_allow_html=True)

# --- 구글 스프레드시트 연결 설정 (배상금 관리용) ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    conn = None

def get_compensation_data():
    if conn is None:
        return pd.DataFrame(columns=["주문번호", "접수일", "처리일", "스토어", "수량", "판매가", "박스수", "합포장", "상품배상금", "택배배상비", "총 배상청구액"])
    try:
        return conn.read(ttl="1s") 
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

# --- 메인 엑셀 파일 업로드 영역 (모든 탭에서 공통 사용) ---
st.subheader("📁 아르고 정산 엑셀 파일 업로드")
uploaded_excel = st.file_uploader("당월 아르고 정산 원본 엑셀 파일(.xlsx)을 이곳에 첨부해 주세요.", type=['xlsx', 'xls'])

# --- 탭 구성 ---
tab1, tab2, tab3, tab4 = st.tabs(["📊 시스템 개요", "📦 입고비 검증", "🚚 출고 배송비 검증", "💰 배상금 정산 관리"])

# ==========================================
# TAB 1: 시스템 개요
# ==========================================
with tab1:
    st.header("시스템 개요")
    st.write("""
    독립 기업인 두유당의 자체적인 기준에 맞추어, 아르고(ARGO) 풀필먼트에서 매월 청구하는 정산 엑셀 데이터의 오류를 정밀하게 검증하는 시스템입니다.
    
    * **이용 방법:** 상단에 원본 엑셀 파일을 업로드하면, 아래 각 탭에서 해당 월의 내역을 분석합니다.
    * **출력 기능:** 식별된 오류 내역은 스타일이 적용된 엑셀 파일로 다운로드하여 증빙 자료로 활용할 수 있습니다.
    * **배상금 관리:** 4번 탭에서 파손/오염 건에 대한 내역을 구글 스프레드시트에 영구적으로 기록할 수 있습니다.
    """)

# ==========================================
# TAB 2: 입고비 검증
# ==========================================
with tab2:
    st.header("입고비 정밀 검증")
    
    if uploaded_excel is None:
        st.info("👆 상단에 엑셀 파일을 먼저 업로드해 주세요.")
    else:
        st.write("아래 검증 조건을 확인하고 실행 버튼을 눌러주세요.")
        
        sku_list = ['하루두유 BLACK', '하루두유 BLACK SWEET', '기타 (직접 입력)']
        selected_sku = st.selectbox("검증할 SKU를 선택하세요:", sku_list, key="inbound_sku")
        if selected_sku == '기타 (직접 입력)':
            selected_sku = st.text_input("SKU 이름을 정확히 입력해 주세요:", key="inbound_custom_sku")
            
        actual_inbound_qty = st.number_input("해당 월의 실제 입고 수량을 기입해 주세요:", min_value=0, step=1, key="inbound_qty")
        
        if st.button("▶ 입고비 검증 실행", type="primary"):
            with st.spinner("입고비 데이터를 분석 중입니다..."):
                df_in = load_excel_sheet(uploaded_excel, sheet_name='입고비', skip_rows=6)
                
                if not df_in.empty:
                    col_idx_in_qty, col_idx_in_amt = 10, 9
                    for i, col_name in enumerate(df_in.columns):
                        if '입고 검수비(기본)' in str(col_name):
                            sub_val = str(df_in.iloc[0, i]).strip()
                            if sub_val == '개수': col_idx_in_qty = i
                            elif sub_val == '금액': col_idx_in_amt = i
                    
                    df_filtered = df_in[df_in['SKU 이름'] == selected_sku]
                    
                    if df_filtered.empty:
                        st.warning(f"업로드된 데이터에 '{selected_sku}' 입고 내역이 존재하지 않습니다.")
                    else:
                        df_filtered['입고유형'] = df_filtered['입고유형'].fillna('')
                        
                        total_billed_qty = 0
                        calculated_total_amt = 0
                        billed_total_amt = 0
                        
                        for index, row in df_filtered.iterrows():
                            if str(row['고객사명']).strip() == '' or str(row.iloc[col_idx_in_qty]).strip() == '개수':
                                continue 
                                
                            try:
                                row_qty_str = str(row.iloc[col_idx_in_qty]).strip()
                                row_billed_amt_str = str(row.iloc[col_idx_in_amt]).strip()
                                
                                row_qty = float(row_qty_str) if row_qty_str.lower() != 'nan' and row_qty_str != '' else 0
                                row_billed_amt = float(row_billed_amt_str) if row_billed_amt_str.lower() != 'nan' and row_billed_amt_str != '' else 0
                                
                                if pd.notna(row_qty):
                                    total_billed_qty += row_qty
                                    if '당일' in str(row['입고유형']):
                                        calculated_total_amt += (row_qty * 200)
                                    else:
                                        calculated_total_amt += (row_qty * 100)
                                        
                                if pd.notna(row_billed_amt):
                                    billed_total_amt += row_billed_amt
                            except:
                                continue
                        
                        st.subheader(f"[{selected_sku}] 입고 검증 결과")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("입력된 실제 입고 수량", f"{actual_inbound_qty:,} 개")
                            st.metric("아르고 청구 입고 수량", f"{int(total_billed_qty):,} 개")
                        with col2:
                            st.metric("아르고 청구 금액", f"{int(billed_total_amt):,} 원")
                            st.metric("자체 산정 금액", f"{int(calculated_total_amt):,} 원")
                        
                        if actual_inbound_qty == total_billed_qty and billed_total_amt == calculated_total_amt:
                            st.success("✅ 수량 및 청구 금액이 모두 정확히 일치합니다.")
                        else:
                            st.error("❌ 수량 또는 금액에 불일치가 발생했습니다. 확인이 필요합니다.")

# ==========================================
# TAB 3: 출고 배송비 검증
# ==========================================
with tab3:
    st.header("출고 배송비 정밀 검증")
    
    if uploaded_excel is None:
        st.info("👆 상단에 엑셀 파일을 먼저 업로드해 주세요.")
    else:
        st.write("데이터가 준비되었습니다. 버튼을 눌러 아르고 청구 내역과 자체 산정 기준을 비교 검증하세요.")
        
        if st.button("▶ 출고 배송비 검증 실행", type="primary"):
            with st.spinner("수천 건의 출고 배송비 데이터를 분석 중입니다. 잠시만 기다려주세요..."):
                df_out = load_excel_sheet(uploaded_excel, sheet_name='출고 배송비', skip_rows=4)
                
                if not df_out.empty:
                    errors_list = []
                    warnings_list = []
                    
                    col_idx_total, col_idx_island, col_idx_same, col_idx_diff = 14, 19, 15, 16
                    for i, val in enumerate(df_out.iloc[0]):
                        val_str = str(val).strip()
                        if val_str == '총 금액': col_idx_total = i
                        elif val_str == '도서 산간 추가 택배비': col_idx_island = i
                        elif val_str == '합포장(동종)': col_idx_same = i
                        elif val_str == '합포장(이종)': col_idx_diff = i
                    
                    for index, row in df_out.iterrows():
                        if index == 0: continue
                        
                        try:
                            sku_count_str = str(row['SKU 개수']).strip()
                            if sku_count_str.lower() == 'nan
