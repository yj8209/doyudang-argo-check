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
                            total_q += q
                            total_billed += b
                            total_calc += (q * (200 if '당일' in str(row['입고유형']) else 100))
                        except: continue
                    st.metric("청구 수량", f"{int(total_q):,} 개")
                    st.metric("청구 금액", f"{int(total_billed):,} 원")
                    st.metric("자체 산정", f"{int(total_calc):,} 원")

# ==========================================
# TAB 3: 출고 배송비 검증
# ==========================================
with tab3:
    st.header("출고 배송비 정밀 검증")
    if uploaded_excel is None:
        st.info("👆 상단에 엑셀 파일을 먼저 업로드해 주세요.")
    else:
        if st.button("▶ 출고 배송비 검증 실행", type="primary"):
            df_out = load_excel_sheet(uploaded_excel, sheet_name='출고 배송비', skip_rows=4)
            if not df_out.empty:
                errs, warns = [], []
                c_tot, c_isl, c_same, c_diff = 14, 19, 15, 16
                for i, v in enumerate(df_out.iloc[0]):
                    s = str(v).strip()
                    if s == '총 금액': c_tot = i
                    elif s == '도서 산간 추가 택배비': c_isl = i
                    elif s == '합포장(동종)': c_same = i
                    elif s == '합포장(이종)': c_diff = i
                
                for idx, r in df_out.iterrows():
                    if idx == 0: continue
                    try:
                        sku_c = int(float(str(r['SKU 개수']).strip()))
                    except: continue 
                    ono = str(r['주문번호']).replace('.0', '').strip()
                    stn = str(r['스토어명']).strip()
                    agrd = str(r['등급']).strip()
                    btot = float(str(r.iloc[c_tot]).replace(',', '').strip())
                    isl_c = 3000 if str(r.iloc[c_isl]).strip().lower() not in ['nan', ''] else 0
                    hs = str(r.iloc[c_same]).strip().lower() not in ['nan', '']
                    hd = str(r.iloc[c_diff]).strip().lower() not in ['nan', '']
                    is_n = '네이버스마트스토어' in stn
                    
                    egrd, base, box, pack = "", 0, 0, 0
                    if sku_c >= 8:
                        warns.append({'주문번호': ono, '스토어': stn, 'SKU': sku_c, '금액': btot})
                        continue
                    elif sku_c == 1: egrd, base, box = "극소", (3050 if is_n else 2750), 220
                    elif sku_c == 2:
                        egrd, base, box = "소", (3600 if is_n else 3300), 220
                        pack = 100 if hd else (50 if hs else 0)
                    elif sku_c in [3, 4]:
                        egrd, base, box = "중", (4100 if is_n else 3800), 450
                        pack = 250 if hd else (150 if hs else 0)
                    elif sku_c == 5:
                        egrd, base, box = "대", (5500 if is_n else 5000), 800
                        pack = 250 if hd else (200 if hs else 0)
                    elif sku_c in [6, 7]:
                        egrd, base, box = "특대", (6300 if is_n else 5800), 800
                        if hd: pack = 400
                        elif hs: pack = 250 if sku_c == 6 else 300
                    
                    etot = base + box + pack + isl_c
                    if btot > etot or agrd != egrd:
                        errs.append({'행': idx+6, '주문번호': ono, 'SKU': sku_c, '청구': btot, '산정': etot, '초과': max(0, btot-etot)})
                
                if errs:
                    df_e = pd.DataFrame(errs)
                    st.error(f"{len(errs)}건의 오류 발견")
                    st.metric("🚨 총 초과 청구액", f"{int(df_e['초과'].sum()):,}")
                    st.dataframe(df_e.style.format({'청구': '{:,.0f}', '산정': '{:,.0f}', '초과': '{:,.0f}'}))
                else: st.success("모든 내역이 정상입니다.")

# ==========================================
# TAB 4: 배상금 정산 관리
# ==========================================
with tab4:
    st.header("💰 배상금 관리 및 데이터 편집")
    
    # 데이터 로드
    df_comp = get_compensation_data()
    display_cols = ["주문번호", "접수일", "처리일", "스토어", "수량", "판매가", "박스수", "합포장", "상품배상금", "택배배상비", "총 배상청구액"]

    if not df_comp.empty:
        df_comp['접수일_DT'] = pd.to_datetime(df_comp['접수일'], errors='coerce')
        df_comp['정산월'] = df_comp['접수일_DT'].dt.strftime('%Y-%m')
        
        m_list = sorted(df_comp['정산월'].dropna().unique(), reverse=True)
        sel_m = st.selectbox("📅 정산 확인 및 편집 월 선택", ["전체 보기"] + m_list)
        
        # 필터링
        if sel_m == "전체 보기": f_df = df_comp[display_cols].copy()
        else: f_df = df_comp[df_comp['정산월'] == sel_m][display_cols].copy()
        
        st.subheader(f"🔍 {sel_m} 상세 내역 (표 안의 값을 더블 클릭하여 수정 가능)")
        st.info("💡 행을 삭제하려면 왼쪽 체크박스 선택 후 Delete 키를 누르세요. 수정 후 하단 버튼을 꼭 눌러야 시트에 반영됩니다.")
        
        edited_df = st.data_editor(
            f_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "총 배상청구액": st.column_config.NumberColumn(format="%d 원"),
                "상품배상금": st.column_config.NumberColumn(format="%d 원"),
                "판매가": st.column_config.NumberColumn(format="%d 원")
            }
        )
        
        col_btn1, col_btn2 = st.columns([1, 4])
        with col_btn1:
            if st.button("💾 변경사항 구글 시트에 최종 적용", type="primary"):
                try:
                    if sel_m == "전체 보기":
                        final_to_save = edited_df
                    else:
                        other_months = df_comp[df_comp['정산월'] != sel_m][display_cols]
                        final_to_save = pd.concat([other_months, edited_df], ignore_index=True)
                    
                    conn.update(data=final_to_save)
                    st.success("✅ 구글 시트 데이터가 성공적으로 수정되었습니다!")
                    st.rerun()
                except Exception as e:
                    st.error(f"저장 중 오류가 발생했습니다: {e}")

        st.metric(f"🚨 {sel_m} 정산 합계", f"{edited_df['총 배상청구액'].sum():,.0f} 원")

    st.markdown("---")
    st.subheader("📝 신규 배상 내역 입력")
    with st.form("new_comp_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            on = st.text_input("주문번호")
            stt = st.selectbox("스토어", ["네이버스마트스토어", "카페24"])
            iq = st.number_input("수량", min_value=1)
        with c2:
            # 변수명을 pd에서 p_date로 변경하여 pandas(pd) 덮어쓰기 방지!
            r_date = st.date_input("접수일", datetime.now())
            p_date = st.date_input("처리일", datetime.now()) 
            up = st.number_input("단가", min_value=0)
        with c3:
            bq = st.number_input("박스 수", min_value=1)
            bt = st.selectbox("합포장", ["없음", "동종", "이종"])
            ii = st.checkbox("도서산간 (+3,000)")
        
        if st.form_submit_button("저장하기"):
            if on:
                pc = up * iq
                sc = ((4800 if stt == "네이버스마트스토어" else 4400) * bq) + (3000 if ii else 0)
                new_r = pd.DataFrame([{
                    "주문번호": on, "접수일": r_date.strftime("%Y-%m-%d"), "처리일": p_date.strftime("%Y-%m-%d"),
                    "스토어": stt, "수량": iq, "판매가": up, "박스수": bq, "합포장": bt,
                    "상품배상금": pc, "택배배상비": sc, "총 배상청구액": pc + sc
                }])
                conn.update(data=pd.concat([df_comp[display_cols], new_r], ignore_index=True))
                st.success("저장 완료!")
                st.rerun()
