import streamlit as st
import pandas as pd
from PIL import Image
import io
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
    st.warning(f"로고 이미지를 찾을 수 없습니다. 경로를 확인해 주세요: {logo_path}")

st.title("두유당 ARGO 월별 정산 정밀 검증 시스템")
st.markdown("""
<style>
    .reportview-container .main .block-container{
        max-width: 1200px;
        padding-top: 2rem;
        padding-right: 2rem;
        padding-left: 2rem;
        padding-bottom: 2rem;
    }
    h1 { color: #1E3A8A; }
    h2 { color: #2563EB; }
    .stFileUploader { padding-bottom: 2rem; border-bottom: 2px solid #E5E7EB; }
</style>
""", unsafe_allow_html=True)

# --- 구글 스프레드시트 연결 설정 (배상금 관리용) ---
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.warning("구글 스프레드시트 연동이 아직 설정되지 않았거나 연결에 실패했습니다. (설정 전이라면 배상금 기록 기능만 제한됩니다.)")
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
    
    * **이용 방법:** 상단에 아르고에서 전달받은 원본 엑셀 파일 하나만 업로드하면, 아래 각 탭에서 해당 월의 내역을 자동으로 분석합니다.
    * **배상금 관리:** 파손/오염 건에 대한 내역을 구글 스프레드시트에 영구적으로 기록하고 관리할 수 있습니다.
    * **출력 기능:** 검증 완료 후 식별된 오류 내역은 스타일이 적용된 엑셀 파일로 다운로드하여 증빙 자료로 활용할 수 있습니다.
    """)

# ==========================================
# TAB 2: 입고비 검증
# ==========================================
with tab2:
    st.header("입고비 정밀 검증")
    
    if uploaded_excel is None:
        st
