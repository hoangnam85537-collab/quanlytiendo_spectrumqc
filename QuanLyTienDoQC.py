import os
import io
import json
import sqlite3
import time
import re
from contextlib import contextmanager
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# ==========================================
# 1. CẤU HÌNH ĐƯỜNG DẪN & CƠ SỞ DỮ LIỆU
# ==========================================
BASE_DIR = r"D:\DATA_PHANMEMQUANLY"
FOLDER_TIEN_DO = os.path.join(BASE_DIR, "TIEN DO DU AN")
FOLDER_REPORT = os.path.join(BASE_DIR, "report bao cao tu dien")
FOLDER_DB = os.path.join(BASE_DIR, "DATABASE")
DB_PATH = os.path.join(FOLDER_DB, "qc_database.db")
FOLDER_TB_TEST = r"D:\QUẢN L. THIẾT BỊ TEST"

FOLDER_DANH_SACH_TU = os.path.join(BASE_DIR, "DANH SÁCH TỦ DỰ ÁN")
FOLDER_HANG_THIEU = os.path.join(BASE_DIR, "danh sách hàng thiếu")
FILE_CHECKLIST_DEFAULT = os.path.join(FOLDER_HANG_THIEU, "Checklist_HN-2026.xlsx")

for folder in [BASE_DIR, FOLDER_TIEN_DO, FOLDER_REPORT, FOLDER_DB, FOLDER_TB_TEST, FOLDER_DANH_SACH_TU, FOLDER_HANG_THIEU]:
    os.makedirs(folder, exist_ok=True)

@contextmanager
def get_db():
    """Context manager giúp quản lý kết nối SQLite an toàn và tự động đóng"""
    conn = sqlite3.connect(DB_PATH)
    try:
        yield conn
    finally:
        conn.close()

def init_db():
    """Khởi tạo CSDL SQLite"""
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS bien_ban_qc (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_tu TEXT UNIQUE,
            ten_du_an TEXT,
            loai_tu TEXT DEFAULT 'DB',
            ngay_kiem_tra TEXT,
            trang_thai_qc TEXT DEFAULT 'Đạt QC',
            trang_thai_test TEXT DEFAULT 'Chưa Test',
            ket_qua TEXT,
            duong_dan_file TEXT,
            danh_sach_loi TEXT,
            ghi_chu TEXT
        )
        ''')
        for col_def in [
            ("trang_thai_qc", "TEXT DEFAULT 'Đạt QC'"),
            ("trang_thai_test", "TEXT DEFAULT 'Chưa Test'"),
            ("loai_tu", "TEXT DEFAULT 'DB'")
        ]:
            try:
                cursor.execute(f"ALTER TABLE bien_ban_qc ADD COLUMN {col_def[0]} {col_def[1]}")
            except Exception:
                pass

        cursor.execute('''
        CREATE TABLE IF NOT EXISTS thiet_bi_test (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten_thiet_bi TEXT,
            model TEXT,
            so_sn TEXT,
            ngay_het_han TEXT,
            ghi_chu TEXT
        )
        ''')
        
        # Bảng lưu trữ Note / Tình trạng dự án chỉnh sửa thủ công
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS ghi_chu_du_an (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ten_du_an TEXT UNIQUE,
            tinh_trang_note TEXT,
            ngay_cap_nhat TEXT
        )
        ''')
        conn.commit()

init_db()

def clean_str_key(s):
    """Hàm chuẩn hóa chuỗi để so sánh mã tủ chính xác hơn"""
    if not s or pd.isna(s):
        return ""
    return re.sub(r'[\s_\-]+', '', str(s).strip().upper())

# ==========================================
# CÁC HÀM CÔNG CỤ (TOOL CALLING CHO AI)
# ==========================================
def cap_nhat_thieu_thanh_du(ten_du_an: str = "", ma_tu: str = ""):
    """Cập nhật tủ thiếu thiết bị hoặc toàn bộ dự án sang trạng thái ĐỦ & PASS."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            if ma_tu and ma_tu != "ALL":
                cur.execute("""
                    UPDATE bien_ban_qc 
                    SET danh_sach_loi = '[]', ket_qua = 'PASS', trang_thai_qc = 'Đạt QC', trang_thai_test = 'Test Pass', ghi_chu = ghi_chu || ' | Đã bổ sung đủ thiết bị qua AI'
                    WHERE ma_tu LIKE ? OR ten_du_an LIKE ?
                """, (f"%{ma_tu}%", f"%{ten_du_an}%"))
            elif ten_du_an:
                cur.execute("""
                    UPDATE bien_ban_qc 
                    SET danh_sach_loi = '[]', ket_qua = 'PASS', trang_thai_qc = 'Đạt QC', trang_thai_test = 'Test Pass', ghi_chu = ghi_chu || ' | Đã bổ sung đủ thiết bị qua AI'
                    WHERE ten_du_an LIKE ?
                """, (f"%{ten_du_an}%",))
            else:
                cur.execute("""
                    UPDATE bien_ban_qc 
                    SET danh_sach_loi = '[]', ket_qua = 'PASS', trang_thai_qc = 'Đạt QC', trang_thai_test = 'Test Pass', ghi_chu = ghi_chu || ' | Đã bổ sung đủ thiết bị qua AI'
                    WHERE ket_qua = 'Thiếu vật tư' OR danh_sach_loi != '[]'
                """)
            conn.commit()
            rowCount = cur.rowcount
        return f"ĐÃ CẬP NHẬT: Đã chuyển {rowCount} bản ghi sang trạng thái ĐỦ & PASS!"
    except Exception as e:
        return f"LỖI CSDL: {str(e)}"

def cap_nhat_du_an_da_giao(ten_du_an: str, ma_tu: str = "TU-GIAO", loai_tu: str = "DB", so_luong: int = 1, ghi_chu: str = "Giao hàng qua AI"):
    """Ghi nhận số lượng tủ đã giao và đạt chuẩn."""
    try:
        with get_db() as conn:
            cur = conn.cursor()
            so_luong_num = int(so_luong) if str(so_luong).isdigit() and int(so_luong) > 0 else 1
            for i in range(so_luong_num):
                ma_single = f"{ma_tu}_{i+1}" if so_luong_num > 1 else ma_tu
                cur.execute("""
                    INSERT OR REPLACE INTO bien_ban_qc (ma_tu, ten_du_an, loai_tu, ngay_kiem_tra, trang_thai_qc, trang_thai_test, ket_qua, duong_dan_file, danh_sach_loi, ghi_chu)
                    VALUES (?, ?, ?, date('now'), 'Đạt QC', 'Test Pass', 'PASS', 'AI_UPDATED', '[]', ?)
                """, (ma_single, ten_du_an.strip().upper(), loai_tu, ghi_chu))
            conn.commit()
        return f"ĐÃ GHI NHẬN {so_luong_num} tủ {loai_tu} cho dự án '{ten_du_an}' (Đạt QC & Test Pass)!"
    except Exception as e:
        return f"LỖI CSDL: {str(e)}"

def get_latest_excel_tiendo():
    """Đọc file Excel tiến độ mới nhất"""
    if not os.path.exists(FOLDER_TIEN_DO):
        return None, "Thư mục không tồn tại", {}, pd.DataFrame()
    excel_files = [os.path.join(FOLDER_TIEN_DO, f) for f in os.listdir(FOLDER_TIEN_DO) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    if not excel_files:
        return None, "Không tìm thấy file Excel", {}, pd.DataFrame()
    latest_file = max(excel_files, key=os.path.getmtime)
    try:
        df = pd.read_excel(latest_file)
        stats = {'tong_du_an': 0, 'tong_so_tu': 0, 'du_an_den_lich': 0}
        df_reminders = pd.DataFrame()
        col_du_an = next((c for c in df.columns if 'DỰ ÁN' in str(c).upper()), None)
        col_so_tu = next((c for c in df.columns if 'SỐ TỦ' in str(c).upper()), None)
        col_ngay_giao = next((c for c in df.columns if 'NGÀY GIAO' in str(c).upper()), None)
        col_ngay_fat = next((c for c in df.columns if 'NGÀY FAT' in str(c).upper()), None)
        
        if col_du_an:
            df_clean = df.dropna(subset=[col_du_an]).copy()
            df_clean = df_clean[~df_clean[col_du_an].astype(str).str.contains('TỦ ĐIỆN|Dự án ưu tiên', case=False, na=False)]
            stats['tong_du_an'] = df_clean[col_du_an].nunique()
            if col_so_tu:
                stats['tong_so_tu'] = int(pd.to_numeric(df_clean[col_so_tu], errors='coerce').fillna(0).sum())
            target_date_col = col_ngay_giao if col_ngay_giao else col_ngay_fat
            if target_date_col:
                now = pd.Timestamp.now().normalize()
                parsed_dates = pd.to_datetime(df_clean[target_date_col], errors='coerce')
                df_clean['Ngay_Giao_Parsed'] = parsed_dates
                df_clean['So_Ngay_Con_Lai'] = (df_clean['Ngay_Giao_Parsed'] - now).dt.days
                df_reminders = df_clean[df_clean['Ngay_Giao_Parsed'].notna() & (df_clean['So_Ngay_Con_Lai'] <= 7)].copy()
                stats['du_an_den_lich'] = df_reminders[col_du_an].nunique()
        return df, os.path.basename(latest_file), stats, df_reminders
    except Exception as e:
        return None, f"Lỗi đọc file Excel: {str(e)}", {}, pd.DataFrame()

# ==========================================
# 2. SCHEMA AI TRÍCH XUẤT ẢNH QC
# ==========================================
class MissingItem(BaseModel):
    ten_vat_tu: str = Field(description="Tên thiết bị hoặc vật tư thiếu ghi trên biên bản")
    so_luong: int = Field(description="Số lượng thiếu")

class QCRecordSchema(BaseModel):
    du_an: str = Field(description="Tên dự án")
    ma_tu: str = Field(description="Mã tủ điện")
    loai_tu: str = Field(default="DB", description="MSB, DB, ATS, MCC hoặc Khác")
    danh_sach_thieu: list[MissingItem] = Field(description="Danh sách vật tư thiếu")
    ghi_chu: str = Field(default="", description="Ghi chú thêm trên biên bản")

# ==========================================
# 3. GIAO DIỆN STREAMLIT MAIN
# ==========================================
st.set_page_config(page_title="Hệ Thống Quản Lý QC & Tiến Độ Tủ Điện", layout="wide", page_icon="⚡")

st.markdown("""
<style>
.stMetric { background-color: var(--secondary-background-color) !important; color: var(--text-color) !important; padding: 15px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.2); border: 1px solid var(--border-color); }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: #f1f5f9; border-radius: 8px 8px 0px 0px; padding: 10px 16px; font-weight: 600; color: #334155; }
    .stTabs [aria-selected="true"] { background-color: #0284c7 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

if "success_msg" in st.session_state:
    st.toast(st.session_state["success_msg"], icon="🎉")
    del st.session_state["success_msg"]

st.sidebar.title("⚙️ Cấu Hình Hệ Thống")
api_key_input = st.sidebar.text_input("🔑 Nhập Gemini API Key:", type="password")
model_choice = st.sidebar.selectbox("🤖 Mô hình Gemini AI:", ["gemini-3.6-flash", "gemini-3.6-pro"], index=0)

if api_key_input:
    st.sidebar.success("✅ Đã nhận API Key")
else:
    st.sidebar.warning("⚠️ Nhập API Key để dùng tính năng AI.")

st.sidebar.markdown("---")
st.sidebar.info(
    f"📁 **Thư mục Báo cáo QC:**\n`{FOLDER_REPORT}`\n\n"
    f"🔬 **Thư mục Thiết Bị:**\n`{FOLDER_TB_TEST}`\n\n"
    f"📑 **Danh sách tủ dự án:**\n`{FOLDER_DANH_SACH_TU}`\n\n"
    f"📦 **Checklist Hàng Thiếu:**\n`{FOLDER_HANG_THIEU}`"
)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📊 Tổng Quan & Tiến Độ Dự Án", 
    "📋 Nhập Liệu QC (Thủ Công/Ảnh Mẫu)", 
    "🔍 Tra Cứu & Báo Cáo Chi Tiết",
    "💬 Trợ Lý AI Hỏi Đáp & Cập Nhật",
    "🔬 Quản Lý Thiết Bị Test",
    "📸 Quét Thư Mục Ảnh QC Tự Động",
    "📋 Đối Chiếu Tủ Dự Án & Hàng Thiếu"
])

with get_db() as conn:
    df_bien_ban = pd.read_sql_query("SELECT * FROM bien_ban_qc ORDER BY id DESC", conn)

if not df_bien_ban.empty and 'ten_du_an' in df_bien_ban.columns:
    df_bien_ban['ten_du_an_clean'] = df_bien_ban['ten_du_an'].astype(str).str.strip().str.upper()
else:
    df_bien_ban['ten_du_an_clean'] = ""

df_excel, excel_info, excel_stats, df_reminders = get_latest_excel_tiendo()

if not df_bien_ban.empty and not df_reminders.empty:
    delivered_projects = df_bien_ban[df_bien_ban['ket_qua'].isin(['PASS', 'Đã Giao', 'Đã giao'])]['ten_du_an_clean'].dropna().tolist()
    col_du_an_rem = next((c for c in df_reminders.columns if 'DỰ ÁN' in str(c).upper()), 'DỰ ÁN')
    def check_not_delivered(row):
        p_name = str(row[col_du_an_rem]).strip().upper()
        for d_p in delivered_projects:
            if d_p in p_name or p_name in d_p: return False
        return True
    df_reminders = df_reminders[df_reminders.apply(check_not_delivered, axis=1)]
    excel_stats['du_an_den_lich'] = df_reminders[col_du_an_rem].nunique() if not df_reminders.empty else 0

# ==========================================
# TAB 1: TỔNG QUAN & TIẾN ĐỘ DỰ ÁN
# ==========================================
with tab1:
    st.title("📊 Tổng Quan Dự Án & Tiến Độ Sản Xuất Tủ Điện")
    st.caption("💡 *Giao diện tích hợp toàn diện các chỉ số KPI, 3 biểu đồ phân tích chuyên sâu và bảng ghi chú tình trạng dự án theo thời gian thực.*")
    
    san_bay_dh_df = df_bien_ban[df_bien_ban['ten_du_an_clean'].str.contains('SÂN BAY ĐỒNG HỚI|DỒNG HỚI|DONG HOI', na=False)]
    san_bay_dh_count = len(san_bay_dh_df)
    
    tong_so_tu_excel = excel_stats.get('tong_so_tu', 0)
    total_tu_qc = len(df_bien_ban)
    
    if san_bay_dh_count == 1:
        total_tu_qc_display = max(0, total_tu_qc - 1)
        tru_text = " (Đã trừ 1 tủ Sân bay Đồng Hới)"
    else:
        total_tu_qc_display = total_tu_qc
        tru_text = ""

    ncr_count = len(df_bien_ban[df_bien_ban['ket_qua'] == 'Thiếu vật tư']) if total_tu_qc > 0 else 0
    msb_count = len(df_bien_ban[df_bien_ban['loai_tu'].astype(str).str.upper().str.contains('MSB', na=False)]) if not df_bien_ban.empty else 0
    db_count = len(df_bien_ban[df_bien_ban['loai_tu'].astype(str).str.upper().str.contains('DB', na=False)]) if not df_bien_ban.empty else 0

    st.markdown("##### ⚡ Chỉ Số Tổng Quan Hệ Thống")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("📦 Kế Hoạch (Excel)", f"{tong_so_tu_excel} Tủ")
    kpi2.metric("✅ Đã QC Check", f"{total_tu_qc_display} Tủ", delta=f"MSB: {msb_count} | DB: {db_count}{tru_text}")
    kpi3.metric("⚡ Đã Test Pass", f"{len(df_bien_ban[df_bien_ban['trang_thai_test'] == 'Test Pass'])} Tủ")
    kpi4.metric("⚠️ Tủ Thiếu Vật Tư (NCR)", f"{ncr_count} Tủ", delta="Cần mua bù" if ncr_count > 0 else "Tốt", delta_color="inverse")
    
    st.markdown("---")
    
    if isinstance(df_excel, pd.DataFrame) and not df_excel.empty:
        col_du_an = next((c for c in df_excel.columns if 'DỰ ÁN' in str(c).upper()), None)
        col_so_tu = next((c for c in df_excel.columns if 'SỐ TỦ' in str(c).upper()), None)
        col_ngay_giao = next((c for c in df_excel.columns if 'NGÀY GIAO' in str(c).upper()), None)
        col_ngay_fat = next((c for c in df_excel.columns if 'NGÀY FAT' in str(c).upper()), None)
        
        if col_du_an:
            df_project_plan = df_excel.dropna(subset=[col_du_an]).copy()
            df_project_plan = df_project_plan[~df_project_plan[col_du_an].astype(str).str.contains('TỦ ĐIỆN|Dự án ưu tiên', case=False, na=False)]
            df_project_plan['Tong_Tu_Ke_Hoach'] = pd.to_numeric(df_project_plan[col_so_tu], errors='coerce').fillna(1) if col_so_tu else 1
            
            def clean_project_name(text):
                if not isinstance(text, str): return ""
                text_no_code = re.sub(r'^\d+[\s\-]*', '', text.strip())
                return " ".join(text_no_code.upper().split())

            df_project_plan['key_excel'] = df_project_plan[col_du_an].apply(clean_project_name)
            df_summary = df_project_plan.groupby([col_du_an, 'key_excel'])['Tong_Tu_Ke_Hoach'].sum().reset_index()
            df_summary.columns = ['ten_du_an_goc', 'key_excel', 'tong_so_tu']

            so_tu_qc_list = []
            so_tu_test_list = []

            if not df_bien_ban.empty:
                df_bien_ban['key_qc'] = df_bien_ban['ten_du_an'].apply(clean_project_name)

                for _, row in df_summary.iterrows():
                    k_excel = row['key_excel']
                    if not k_excel:
                        so_tu_qc_list.append(0)
                        so_tu_test_list.append(0)
                        continue

                    matched_qc = df_bien_ban[df_bien_ban['key_qc'].apply(lambda x: bool(x and (x in k_excel or k_excel in x)))]
                    count_qc = len(matched_qc)
                    
                    if "ĐỒNG HỚI" in k_excel or "DONG HOI" in k_excel:
                        if count_qc == 1:
                            count_qc = 0
                            
                    count_test = len(matched_qc[matched_qc['trang_thai_test'] == 'Test Pass'])

                    so_tu_qc_list.append(count_qc)
                    so_tu_test_list.append(count_test)
            else:
                so_tu_qc_list = [0] * len(df_summary)
                so_tu_test_list = [0] * len(df_summary)

            df_summary['so_tu_dat_qc'] = so_tu_qc_list
            df_summary['so_tu_test_pass'] = so_tu_test_list

            # ==========================================
            # 3 BIỂU ĐỒ TRỰC QUAN
            # ==========================================
            st.markdown("### 📊 HỆ THỐNG 3 BIỂU ĐỒ PHÂN TÍCH TIẾN ĐỘ")
            
            c_chart1, c_chart2 = st.columns(2)
            
            with c_chart1:
                # 1. Biểu Đồ Số Lượng Tủ Điện Tổng Các Dự Án
                st.markdown("##### 1️⃣ Biểu Đồ Số Lượng Tủ Điện Tổng Các Dự Án")
                fig_total_bar = px.bar(
                    df_summary,
                    x='ten_du_an_goc',
                    y='tong_so_tu',
                    text='tong_so_tu',
                    color='tong_so_tu',
                    color_continuous_scale='Blues',
                    labels={'ten_du_an_goc': 'Tên Dự Án', 'tong_so_tu': 'Tổng Số Tủ'}
                )
                fig_total_bar.update_traces(textposition='outside')
                fig_total_bar.update_layout(
                    height=380,
                    xaxis={'tickangle': -35},
                    margin=dict(t=30, b=80, l=20, r=20),
                    showlegend=False
                )
                st.plotly_chart(fig_total_bar, use_container_width=True, key="chart_total_tu_dien")

            # 2. BIỂU ĐỒ CỘT NGANG ĐẾM NGƯỢC THỜI GIAN GIAO TỦ THEO LỊCH
            st.markdown("##### 2️⃣ Biểu Đồ Cột Ngang Đếm Ngược Ngày Giao Tủ")
            target_col_date = col_ngay_giao if col_ngay_giao else col_ngay_fat

            if target_col_date:
                df_countdown = df_project_plan.dropna(subset=[target_col_date]).copy()
                now_ts = pd.Timestamp.now().normalize()
                
                def extract_valid_date(text):
                    if pd.isna(text) or str(text).strip() == "":
                        return pd.NaT
                    text_str = str(text)
                    match_full = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{4})', text_str)
                    if match_full:
                        return pd.to_datetime(match_full.group(1), dayfirst=True, errors='coerce')
                    match_iso = re.search(r'(\d{4}-\d{1,2}-\d{1,2})', text_str)
                    if match_iso:
                        return pd.to_datetime(match_iso.group(1), errors='coerce')
                    match_short = re.search(r'(\d{1,2}[/-]\d{1,2})', text_str)
                    if match_short:
                        return pd.to_datetime(f"{match_short.group(1)}/{pd.Timestamp.now().year}", dayfirst=True, errors='coerce')
                    return pd.NaT

                df_countdown['Ngay_Giao_Parsed'] = df_countdown[target_col_date].apply(extract_valid_date)
                df_countdown = df_countdown.dropna(subset=['Ngay_Giao_Parsed'])
                df_countdown['So_Ngay_Con_Lai'] = (df_countdown['Ngay_Giao_Parsed'] - now_ts).dt.days
                df_countdown = df_countdown.sort_values(by='So_Ngay_Con_Lai', ascending=True)

                if not df_countdown.empty:
                    fig_countdown = px.bar(
                        df_countdown,
                        x='So_Ngay_Con_Lai',
                        y=col_du_an,
                        orientation='h',
                        text='So_Ngay_Con_Lai',
                        color='So_Ngay_Con_Lai',
                        color_continuous_scale='Tealrose',
                        labels={'So_Ngay_Con_Lai': 'Số Ngày Còn Lại (Ngày)', col_du_an: 'Dự Án'}
                    )
                    fig_countdown.update_traces(texttemplate='%{text} ngày', textposition='outside')
                    fig_countdown.update_layout(
                        height=380,
                        margin=dict(t=30, b=20, l=20, r=40),
                        showlegend=False
                    )
                    st.plotly_chart(fig_countdown, use_container_width=True, key="chart_countdown_ngay_giao")
                else:
                    st.info("⚠️ Không có dữ liệu ngày giao hợp lệ để vẽ biểu đồ đếm ngược.")
            else:
                st.info("⚠️ Không tìm thấy cột ngày giao trong file Excel tiến độ.")  

            # 3. BIỂU ĐỒ CỘT NGANG TỔNG HỢP TIẾN ĐỘ (Kế hoạch vs QC vs Test)
            st.markdown("##### 3️⃣ Biểu Đồ Cột Ngang So Sánh Tổng Hợp (Kế Hoạch — QC — Test)")
            df_summary_sorted = df_summary.sort_values(by='tong_so_tu', ascending=True)

            fig_summary = go.Figure()
            fig_summary.add_trace(go.Bar(
                y=df_summary_sorted['ten_du_an_goc'], 
                x=df_summary_sorted['tong_so_tu'], 
                name='📦 Tổng Tủ Kế Hoạch', 
                orientation='h',
                marker=dict(color='#64748b', opacity=0.85)
            ))
            fig_summary.add_trace(go.Bar(
                y=df_summary_sorted['ten_du_an_goc'], 
                x=df_summary_sorted['so_tu_dat_qc'], 
                name='🟢 Đã QC Check', 
                orientation='h',
                marker=dict(color='#059669'),
                text=df_summary_sorted['so_tu_dat_qc'], 
                textposition='auto'
            ))
            fig_summary.add_trace(go.Bar(
                y=df_summary_sorted['ten_du_an_goc'], 
                x=df_summary_sorted['so_tu_test_pass'], 
                name='⚡ Đã Test Pass', 
                orientation='h',
                marker=dict(color='#2563eb'),
                text=df_summary_sorted['so_tu_test_pass'], 
                textposition='auto'
            ))

            dynamic_height = max(420, len(df_summary_sorted) * 35)
            fig_summary.update_layout(
                barmode='group',
                xaxis_title="Số Lượng Tủ Điện",
                yaxis_title="",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                height=dynamic_height,
                margin=dict(l=10, r=20, t=40, b=20),
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_summary, use_container_width=True, key="plotly_project_summary_fixed_v7")

            st.markdown("---")
            
            # TÍNH NĂNG NOTE & TÌNH TRẠNG DỰ ÁN TRÊN BẢNG THÔNG TIN
            st.markdown("### 📝 Bảng Thông Tin Dự Án & Ghi Chú Tình Trạng (Note)")
            st.caption("💡 *Anh có thể cập nhật trực tiếp ghi chú/tình trạng sản xuất của từng dự án vào bảng dưới đây. Dữ liệu sẽ được lưu tự động vào Cơ sở dữ liệu.*")

            with get_db() as conn:
                df_notes = pd.read_sql_query("SELECT * FROM ghi_chu_du_an", conn)
            note_dict = {row['ten_du_an']: row['tinh_trang_note'] for _, row in df_notes.iterrows()} if not df_notes.empty else {}

            df_display_table = df_project_plan.copy()
            df_display_table['Tình Trạng & Ghi Chú (Note)'] = df_display_table[col_du_an].map(note_dict).fillna("")

            edited_notes_df = st.data_editor(
                df_display_table[[col_du_an, col_so_tu if col_so_tu else col_du_an, col_ngay_giao if col_ngay_giao else col_du_an, 'Tình Trạng & Ghi Chú (Note)']],
                use_container_width=True,
                num_rows="fixed",
                key="editor_tinh_trang_du_an",
                column_config={
                    "Tình Trạng & Ghi Chú (Note)": st.column_config.TextColumn(
                        "📝 Ghi Chú Tình Trạng Dự Án (Click đúp để sửa)",
                        help="Nhập tình trạng tiến độ, vướng mắc vật tư, lịch giao hàng thực tế...",
                        width="large"
                    )
                }
            )

            if st.button("💾 Lưu Ghi Chú & Tình Trạng Dự Án", type="primary", key="btn_save_project_notes"):
                try:
                    with get_db() as conn:
                        cur = conn.cursor()
                        for _, r in edited_notes_df.iterrows():
                            p_name = r[col_du_an]
                            p_note = r['Tình Trạng & Ghi Chú (Note)']
                            if pd.notna(p_name) and str(p_name).strip():
                                cur.execute("""
                                    INSERT OR REPLACE INTO ghi_chu_du_an (ten_du_an, tinh_trang_note, ngay_cap_nhat)
                                    VALUES (?, ?, date('now'))
                                """, (str(p_name).strip(), str(p_note).strip()))
                        conn.commit()
                    st.success("✅ Đã lưu thành công ghi chú tình trạng dự án vào cơ sở dữ liệu!")
                except Exception as e:
                    st.error(f"❌ Lỗi khi lưu ghi chú: {str(e)}")

# ==========================================
# TAB 2: NHẬP LIỆU QC
# ==========================================
with tab2:
    st.title("📋 Nhập Liệu & Kiểm Duyệt Biên Bản QC")
    uploaded_file = st.file_uploader("📸 Tải lên ảnh chụp biên bản QC (Đã check xong)", type=["jpg", "png", "jpeg"], key="upload_tab3")
    if uploaded_file:
        col_img, col_form = st.columns([1, 1])
        image = Image.open(uploaded_file)
        with col_img:
            st.image(image, use_container_width=True)
        with col_form:
            if not api_key_input:
                st.error("❌ Chưa nhập Gemini API Key!")
            else:
                if "parsed_qc_data" not in st.session_state or st.button("🔄 Quét lại bằng AI"):
                    with st.spinner("AI đang đọc biên bản QC..."):
                        try:
                            client = genai.Client(api_key=api_key_input)
                            response = client.models.generate_content(
                                model=model_choice,
                                contents=[image, "Đọc và trích xuất thông tin biên bản QC này từ ảnh."],
                                config=types.GenerateContentConfig(response_mime_type="application/json", response_schema=QCRecordSchema)
                            )
                            st.session_state.parsed_qc_data = json.loads(response.text)
                            st.success("✅ Đã quét xong!")
                        except Exception as e:
                            st.error(f"Lỗi AI: {str(e)}")
                
                if "parsed_qc_data" in st.session_state:
                    data = st.session_state.parsed_qc_data
                    du_an = st.text_input("Tên Dự Án:", value=data.get("du_an", ""))
                    ma_tu = st.text_input("Mã Tủ Điện:", value=data.get("ma_tu", ""))
                    loai_tu_selected = st.selectbox("🏷️ Loại Tủ:", ["MSB", "DB", "ATS", "MCC", "Khác"], index=1)
                    
                    st.markdown("##### Trạng thái kiểm tra:")
                    trang_thai_qc = st.selectbox("QC Check Status:", ["Đạt QC", "Đã Check (Có lỗi/Thiếu)"], index=0)
                    trang_thai_test = st.selectbox("Test Status:", ["Test Pass", "Chưa Test", "Test Fail"], index=0)
                    
                    ghi_chu = st.text_area("Ghi chú thêm:", value=data.get("ghi_chu", ""))
                    edited_items = st.data_editor(data.get("danh_sach_thieu", []), num_rows="dynamic", key="qc_items_ed", use_container_width=True)
                    
                    if st.button("💾 Lưu Vào CSDL (Ghi đè nếu trùng mã tủ)", type="primary"):
                        if ma_tu:
                            file_save_path = os.path.join(FOLDER_REPORT, f"{ma_tu}_QC.jpg")
                            image.save(file_save_path)
                            ket_qua_status = "Thiếu vật tư" if len(edited_items) > 0 else "PASS"
                            
                            with get_db() as conn:
                                cur = conn.cursor()
                                cur.execute("DELETE FROM bien_ban_qc WHERE ma_tu = ?", (ma_tu,))
                                cur.execute("""
                                    INSERT INTO bien_ban_qc 
                                    (ma_tu, ten_du_an, loai_tu, ngay_kiem_tra, trang_thai_qc, trang_thai_test, ket_qua, duong_dan_file, danh_sach_loi, ghi_chu) 
                                    VALUES (?, ?, ?, date('now'), ?, ?, ?, ?, ?, ?)
                                """, (ma_tu, du_an.strip().upper(), loai_tu_selected, trang_thai_qc, trang_thai_test, ket_qua_status, file_save_path, json.dumps(edited_items, ensure_ascii=False), ghi_chu))
                                conn.commit()
                            
                            del st.session_state["parsed_qc_data"]
                            st.session_state["success_msg"] = f"Đã ghi đè & lưu thành công tủ {ma_tu}!"
                            st.rerun()

# ==========================================
# TAB 3: TRA CỨU & BÁO CÁO CHI TIẾT
# ==========================================
with tab3:
    st.title("🔍 Tra Cứu & Báo Cáo Chi Tiết Trạng Thái Test Tủ")
    
    if not df_bien_ban.empty:
        st.subheader("📊 Biểu Đồ Cột Trạng Thái Test Theo Từng Dự Án")
        
        unique_projects = [p for p in df_bien_ban['ten_du_an'].unique() if p and str(p).strip()]
        
        if unique_projects:
            st.info(f"💡 Hệ thống ghi nhận **{len(unique_projects)} dự án** trong dữ liệu. Dưới đây là {len(unique_projects)} biểu đồ cột tương ứng:")
            
            cols_t4 = st.columns(2 if len(unique_projects) > 1 else 1)
            
            for p_idx, p_name in enumerate(unique_projects):
                df_p = df_bien_ban[df_bien_ban['ten_du_an'] == p_name]
                
                test_counts = df_p['trang_thai_test'].value_counts().reset_index()
                test_counts.columns = ['trang_thai', 'so_luong']
                
                fig_bar_p = px.bar(
                    test_counts,
                    x='trang_thai',
                    y='so_luong',
                    text='so_luong',
                    color='trang_thai',
                    title=f"🏢 Dự án: {p_name}",
                    color_discrete_map={
                        'Test Pass': '#059669',
                        'Chưa Test': '#f59e0b',
                        'Test Fail': '#dc2626'
                    }
                )
                fig_bar_p.update_traces(textposition='outside')
                fig_bar_p.update_layout(
                    height=300,
                    xaxis_title="",
                    yaxis_title="Số tủ",
                    showlegend=False,
                    margin=dict(l=20, r=20, t=40, b=20)
                )
                
                with cols_t4[p_idx % len(cols_t4)]:
                    st.plotly_chart(fig_bar_p, use_container_width=True, key=f"bar_chart_proj_tab4_{p_idx}")
        else:
            st.info("ℹ️ Chưa có dự án nào có dữ liệu trong CSDL.")

    st.markdown("---")
    st.subheader("🔍 Tra Cứu & Chỉnh Sửa Bản Ghi QC")
    search_keyword = st.text_input("🔎 Tìm kiếm theo mã tủ hoặc dự án:", "")
    if not df_bien_ban.empty:
        df_display = df_bien_ban[df_bien_ban['ma_tu'].str.contains(search_keyword, case=False, na=False) | df_bien_ban['ten_du_an'].str.contains(search_keyword, case=False, na=False)] if search_keyword else df_bien_ban
        st.dataframe(df_display[['id', 'ma_tu', 'ten_du_an', 'loai_tu', 'ngay_kiem_tra', 'trang_thai_qc', 'trang_thai_test', 'ket_qua', 'ghi_chu']], use_container_width=True)
        
        selected_id = st.selectbox("Chọn ID tủ để sửa/xóa:", df_display['id'].tolist() if not df_display.empty else [])
        if selected_id:
            row = df_display[df_display['id'] == selected_id].iloc[0]
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                with st.form(key=f"edit_form_{selected_id}"):
                    e_ma = st.text_input("Mã Tủ:", value=str(row['ma_tu']))
                    e_da = st.text_input("Tên Dự Án:", value=str(row['ten_du_an']))
                    e_loai = st.selectbox("Loại Tủ:", ["MSB", "DB", "ATS", "MCC", "Khác"], index=1)
                    e_tt_qc = st.selectbox("Trạng thái QC:", ["Đạt QC", "Đã Check (Có lỗi/Thiếu)"], index=0)
                    e_tt_test = st.selectbox("Trạng thái Test:", ["Test Pass", "Chưa Test", "Test Fail"], index=0 if row.get('trang_thai_test') == 'Test Pass' else 1)
                    e_kq = st.selectbox("Kết Quả:", ["PASS", "Thiếu vật tư"], index=0 if row['ket_qua'] in ['PASS', 'Đã Giao'] else 1)
                    e_gc = st.text_area("Ghi chú:", value=str(row['ghi_chu'] or ''))
                    if st.form_submit_button("💾 Cập Nhật"):
                        with get_db() as conn:
                            cur = conn.cursor()
                            cur.execute("UPDATE bien_ban_qc SET ma_tu=?, ten_du_an=?, loai_tu=?, trang_thai_qc=?, trang_thai_test=?, ket_qua=?, ghi_chu=? WHERE id=?", 
                                        (e_ma, e_da.strip().upper(), e_loai, e_tt_qc, e_tt_test, e_kq, e_gc, int(selected_id)))
                            conn.commit()
                        st.session_state["success_msg"] = f"Đã cập nhật tủ ID {selected_id}!"
                        st.rerun()
            with col_e2:
                if st.button("🗑️ Xóa Bản Ghi Này"):
                    with get_db() as conn:
                        cur = conn.cursor()
                        cur.execute("DELETE FROM bien_ban_qc WHERE id=?", (int(selected_id),))
                        conn.commit()
                    st.session_state["success_msg"] = f"Đã xóa tủ ID {selected_id}!"
                    st.rerun()

# ==========================================
# TAB 4: TRỢ LÝ AI HỎI ĐÁP & FUNCTION CALLING
# ==========================================
with tab4:
    st.title("💬 Trợ Lý AI Hỏi Đáp & Cập Nhật Giao Hàng / Vật Tư")
    if not api_key_input:
        st.info("💡 Vui lòng nhập Gemini API Key ở Sidebar.")
    else:
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = [{"role": "assistant", "content": "Xin chào! Bạn cần tra cứu dự án nào hoặc muốn cập nhật tủ thiếu thiết bị thành đủ, cứ nhắn tôi nhé!"}]
        for msg in st.session_state.chat_messages:
            st.chat_message(msg["role"]).write(msg["content"])
            
        user_prompt = st.chat_input("Gõ yêu cầu...")
        if user_prompt:
            st.session_state.chat_messages.append({"role": "user", "content": user_prompt})
            st.chat_message("user").write(user_prompt)
            excel_context = df_excel.to_string() if isinstance(df_excel, pd.DataFrame) else "Không có dữ liệu Excel"
            qc_context = df_bien_ban.to_string() if not df_bien_ban.empty else "Không có dữ liệu QC"
            
            system_instruction = f"""
            Bạn là trợ lý AI quản lý sản xuất và QC tủ điện Hải Nam.
            Dữ liệu tiến độ Excel: \n{excel_context}\n
            Dữ liệu QC SQLite: \n{qc_context}\n
            Nếu người dùng yêu cầu cập nhật tủ thiếu thành đủ hoặc xác nhận giao hàng, hãy gọi hàm tương ứng.
            """
            try:
                client = genai.Client(api_key=api_key_input)
                response = client.models.generate_content(
                    model=model_choice,
                    contents=user_prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        tools=[cap_nhat_thieu_thanh_du, cap_nhat_du_an_da_giao]
                    )
                )
                
                if response.function_calls:
                    for function_call in response.function_calls:
                        name = function_call.name
                        args = function_call.args
                        result_msg = ""
                        if name == "cap_nhat_thieu_thanh_du":
                            result_msg = cap_nhat_thieu_thanh_du(**args)
                        elif name == "cap_nhat_du_an_da_giao":
                            result_msg = cap_nhat_du_an_da_giao(**args)
                        
                        st.session_state.chat_messages.append({"role": "assistant", "content": result_msg})
                        st.chat_message("assistant").write(result_msg)
                        st.rerun()
                else:
                    reply_text = response.text
                    st.session_state.chat_messages.append({"role": "assistant", "content": reply_text})
                    st.chat_message("assistant").write(reply_text)
            except Exception as e:
                err_msg = f"Lỗi AI: {str(e)}"
                st.error(err_msg)

# ==========================================
# TAB 5: QUẢN LÝ THIẾT BỊ TEST
# ==========================================
with tab5:
    st.title("🔬 Quản Lý Thiết Bị Test & Hiệu Chuẩn")
    st.info(f"📂 Thư mục quản lý thiết bị test: `{FOLDER_TB_TEST}`")
    
    with st.form("form_them_thiet_bi"):
        st.subheader("➕ Thêm Thiết Bị Test Mới")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            tb_ten = st.text_input("Tên thiết bị (VD: Đồng hồ vạn năng, Hioki...):")
            tb_model = st.text_input("Model:")
        with col_t2:
            tb_sn = st.text_input("Số Serial (S/N):")
            tb_ngay_het_han = st.date_input("Ngày hết hạn hiệu chuẩn:")
        
        tb_ghi_chu = st.text_area("Ghi chú thiết bị:")
        btn_submit_tb = st.form_submit_button("💾 Thêm Thiết Bị")
        
        if btn_submit_tb:
            if tb_ten:
                with get_db() as conn:
                    cur = conn.cursor()
                    cur.execute("""
                        INSERT INTO thiet_bi_test (ten_thiet_bi, model, so_sn, ngay_het_han, ghi_chu)
                        VALUES (?, ?, ?, ?, ?)
                    """, (tb_ten, tb_model, tb_sn, str(tb_ngay_het_han), tb_ghi_chu))
                    conn.commit()
                st.success(f"✅ Đã thêm thiết bị {tb_ten}!")
                st.rerun()
            else:
                st.error("⚠️ Vui lòng nhập tên thiết bị.")

    st.markdown("---")
    st.subheader("📊 Biểu Đồ Quạt Hạn Hiệu Chuẩn Từng Thiết Bị")
    
    with get_db() as conn:
        df_tb = pd.read_sql_query("SELECT * FROM thiet_bi_test ORDER BY id DESC", conn)
    
    if not df_tb.empty:
        gauge_cols = st.columns(min(3, len(df_tb)))
        now_dt = pd.Timestamp.now().normalize()
        
        for idx, row in df_tb.iterrows():
            ngay_hh = pd.to_datetime(row['ngay_het_han'], errors='coerce')
            days_left = (ngay_hh - now_dt).days if pd.notna(ngay_hh) else 0
            days_left_display = max(0, days_left)
            
            with gauge_cols[idx % len(gauge_cols)]:
                st.markdown(f"**🔬 {row['ten_thiet_bi']}**")
                st.caption(f"Model: {row['model'] or 'N/A'} | S/N: {row['so_sn'] or 'N/A'}")
                
                fig_gauge = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=days_left_display,
                    domain={'x': [0, 1], 'y': [0, 1]},
                    number={'suffix': " ngày"},
                    gauge={
                        'axis': {'range': [0, 365], 'tickwidth': 1, 'tickcolor': "darkblue"},
                        'bar': {'color': "#1e40af"},
                        'bgcolor': "white",
                        'borderwidth': 1,
                        'bordercolor': "gray",
                        'steps': [
                            {'range': [0, 30], 'color': '#f87171'},
                            {'range': [30, 90], 'color': '#fbbf24'},
                            {'range': [90, 365], 'color': '#4ade80'}
                        ]
                    }
                ))
                fig_gauge.update_layout(height=170, margin=dict(l=15, r=15, t=10, b=10))
                st.plotly_chart(fig_gauge, use_container_width=True, key=f"gauge_tb_{row['id']}")

        st.markdown("---")
        st.subheader("📋 Danh Sách Thiết Bị Test Hiện Có")
        st.dataframe(df_tb, use_container_width=True)
    else:
        st.info("Chưa có dữ liệu thiết bị test.")

# ==========================================
# TAB 6: QUÉT THƯ MỤC ẢNH QC TỰ ĐỘNG
# ==========================================
with tab6:
    st.title("📸 Quét Thư Mục Ảnh QC Tự Động")
    st.write(f"Hệ thống sẽ quét các tệp ảnh trong thư mục: `{FOLDER_REPORT}`")
    
    if not api_key_input:
        st.warning("⚠️ Vui lòng nhập Gemini API Key ở Sidebar để sử dụng tính năng quét tự động.")
    else:
        if st.button("🚀 Bắt đầu quét thư mục", type="primary"):
            image_extensions = ('.jpg', '.jpeg', '.png')
            all_image_files = [f for f in os.listdir(FOLDER_REPORT) if f.lower().endswith(image_extensions)]
            
            with get_db() as conn:
                scanned_paths = pd.read_sql_query("SELECT duong_dan_file FROM bien_ban_qc", conn)['duong_dan_file'].dropna().tolist()
            
            scanned_filenames = [os.path.basename(p) for p in scanned_paths]
            image_files = [f for f in all_image_files if f not in scanned_filenames]
            
            if not image_files:
                st.info("ℹ️ Tất cả tệp ảnh trong thư mục đã được quét trước đó (Không có ảnh mới).")
            else:
                st.info(f"🔎 Bắt đầu quét {len(image_files)} ảnh mới...")
                progress_bar = st.progress(0)
                status_text = st.empty()
                client = genai.Client(api_key=api_key_input)
                
                success_count = 0
                for idx, file_name in enumerate(image_files):
                    file_path = os.path.join(FOLDER_REPORT, file_name)
                    status_text.text(f"Đang xử lý AI ({idx+1}/{len(image_files)}): {file_name}")
                    
                    try:
                        img = Image.open(file_path)
                        response = client.models.generate_content(
                            model=model_choice,
                            contents=[img, "Đọc và trích xuất chính xác tên mã tủ và thông tin biên bản QC này từ ảnh."],
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                response_schema=QCRecordSchema
                            )
                        )
                        parsed_data = json.loads(response.text)
                        
                        ma_tu_ai = parsed_data.get("ma_tu") or os.path.splitext(file_name)[0]
                        du_an = parsed_data.get("du_an", "KHÔNG RÕ")
                        loai_tu = parsed_data.get("loai_tu", "DB")
                        danh_sach_thieu = parsed_data.get("danh_sach_thieu", [])
                        ghi_chu = parsed_data.get("ghi_chu", "")
                        
                        ket_qua_status = "Thiếu vật tư" if len(danh_sach_thieu) > 0 else "PASS"
                        
                        with get_db() as conn:
                            cur = conn.cursor()
                            cur.execute("DELETE FROM bien_ban_qc WHERE ma_tu = ?", (ma_tu_ai,))
                            
                            cur.execute("""
                                INSERT INTO bien_ban_qc 
                                (ma_tu, ten_du_an, loai_tu, ngay_kiem_tra, trang_thai_qc, trang_thai_test, ket_qua, duong_dan_file, danh_sach_loi, ghi_chu) 
                                VALUES (?, ?, ?, date('now'), 'Đạt QC', 'Test Pass', ?, ?, ?, ?)
                            """, (
                                ma_tu_ai,
                                du_an.strip().upper(),
                                loai_tu,
                                ket_qua_status,
                                file_path,
                                json.dumps(danh_sach_thieu, ensure_ascii=False),
                                ghi_chu
                            ))
                            conn.commit()
                        success_count += 1
                    except Exception as e:
                        st.error(f"Lỗi khi xử lý file {file_name}: {str(e)}")
                    
                    progress_bar.progress((idx + 1) / len(image_files))
                
                status_text.text("✅ Hoàn tất quá trình quét!")
                st.success(f"🎉 Đã tự động xử lý và cập nhật thành công {success_count}/{len(image_files)} tệp ảnh mới!")

import unicodedata
from difflib import SequenceMatcher

# ==========================================
# THUẬT TOÁN SO KHỚP THÔNG MINH (OFFLINE)
# ==========================================
def remove_vietnamese_accents(text):
    if not text or pd.isna(text):
        return ""
    text = unicodedata.normalize('NFKD', str(text))
    text = "".join([c for c in text if not unicodedata.combining(c)])
    return text.replace('đ', 'd').replace('Đ', 'D')

def clean_single_tu(s):
    if not s or pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r'\(.*?\)', '', s)
    s = remove_vietnamese_accents(s)
    s = re.sub(r'[^A-Za-z0-9]', '', s)
    return s.upper()

def extract_tu_tokens(raw_name):
    if not raw_name or pd.isna(raw_name):
        return []
    s = str(raw_name).strip()
    s = re.sub(r'\(.*?\)', '', s)
    parts = re.split(r'[,/;&]|(?:\s+và\s+)', s, flags=re.IGNORECASE)
    tokens = []
    for p in parts:
        cleaned = clean_single_tu(p)
        if cleaned:
            tokens.append(cleaned)
    return tokens

def is_smart_match(name1, name2, threshold=0.80):
    tokens1 = extract_tu_tokens(name1)
    tokens2 = extract_tu_tokens(name2)
    
    if not tokens1 or not tokens2:
        return False
    
    for t1 in tokens1:
        for t2 in tokens2:
            if t1 == t2 or (len(t1) > 3 and t1 in t2) or (len(t2) > 3 and t2 in t1):
                return True
            if SequenceMatcher(None, t1, t2).ratio() >= threshold:
                return True
    return False

# ==========================================
# TAB 7: ĐỐI CHIẾU TỦ DỰ ÁN & CHECKLIST HÀNG THIẾU
# ==========================================
with tab7:
    st.title("📋 Đối Chiếu Danh Sách Tủ Dự Án & Checklist Hàng Thiếu")
    
    col_f1, col_f2 = st.columns(2)
    excel_tu_files = [f for f in os.listdir(FOLDER_DANH_SACH_TU) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$') and not f.startswith('KetQua_')]
    excel_checklist_files = [f for f in os.listdir(FOLDER_HANG_THIEU) if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    
    file_tu_path = None
    saved_file_path = None
    saved_file_name = None
    
    with col_f1:
        if excel_tu_files:
            default_tu_index = 0
            for i, f_name in enumerate(excel_tu_files):
                if 'đồng hới' in f_name.lower() or 'dong hoi' in f_name.lower():
                    default_tu_index = i
                    break
            selected_tu_file = st.selectbox("📂 **Chọn File Danh Sách Tủ Dự Án (Tham Chiếu Chuẩn):**", excel_tu_files, index=default_tu_index, key="select_tu_project_file")
            
            file_tu_path = os.path.join(FOLDER_DANH_SACH_TU, selected_tu_file)
            saved_file_name = f"KetQua_DoiChieu_{selected_tu_file}"
            saved_file_path = os.path.join(FOLDER_DANH_SACH_TU, saved_file_name)
        else:
            st.info(f"ℹ️ Thư mục `{FOLDER_DANH_SACH_TU}` chưa có file Excel nào.")
            
    file_hang_thieu_path = None
    with col_f2:
        if excel_checklist_files:
            default_chk_index = 0
            for i, f_name in enumerate(excel_checklist_files):
                if 'checklist' in f_name.lower():
                    default_chk_index = i
                    break
            selected_chk_file = st.selectbox("📦 **Chọn File Checklist Hàng Thiếu:**", excel_checklist_files, index=default_chk_index, key="select_checklist_file")
            file_hang_thieu_path = os.path.join(FOLDER_HANG_THIEU, selected_chk_file)
        else:
            file_hang_thieu_path = FILE_CHECKLIST_DEFAULT
            st.info(f"ℹ️ Chưa tìm thấy file Excel trong thư mục `{FOLDER_HANG_THIEU}`. Đã gán đường dẫn mặc định: `{FILE_CHECKLIST_DEFAULT}`")

    st.markdown("---")
    
    def load_checklist_hang_thieu(file_path):
        hang_thieu_dict = {}
        debug_info = []

        if not file_path or not os.path.exists(file_path):
            return hang_thieu_dict, debug_info

        try:
            xls = pd.ExcelFile(file_path)
            kw_tu = ['TÊN TỦ', 'PANEL NAME', 'MÃ TỦ', 'KÝ HIỆU', 'STT TỦ', 'TỦ']
            kw_tb = ['TÊN THIẾT BỊ', 'MÃ THIẾT BỊ', 'TÊN HÀNG', 'VẬT TƯ', 'TÊN VẬT TƯ', 'THIẾT BỊ', 'DIỄN GIẢI']
            kw_sl = ['SỐ LƯỢNG', 'SL', 'QTY', 'QUANTITY']

            for sheet in xls.sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
                if df_raw.empty:
                    continue

                header_row_idx = 0
                for idx, row_vals in df_raw.iterrows():
                    if idx > 30: 
                        break
                    row_str = " ".join([str(v).upper() for v in row_vals.values if pd.notna(v)])
                    if any(k in row_str for k in kw_tu) or any(k in row_str for k in kw_tb):
                        header_row_idx = idx
                        break

                df_sheet = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx).dropna(how='all')

                col_tu = next((c for c in df_sheet.columns if any(k in str(c).upper() for k in kw_tu) and 'LOẠI' not in str(c).upper() and 'TRẠNG THÁI' not in str(c).upper()), None)
                col_tb = next((c for c in df_sheet.columns if any(k in str(c).upper() for k in kw_tb)), None)
                col_sl = next((c for c in df_sheet.columns if any(k in str(c).upper() for k in kw_sl)), None)

                if col_tu:
                    df_sheet[col_tu] = df_sheet[col_tu].ffill()

                debug_info.append({
                    "Sheet": sheet,
                    "Dòng Header": header_row_idx,
                    "Cột Tên Tủ": str(col_tu),
                    "Cột Thiết Bị": str(col_tb),
                    "Cột Số Lượng": str(col_sl),
                    "Tổng số dòng đọc": len(df_sheet)
                })

                for _, r in df_sheet.iterrows():
                    raw_tu = str(r[col_tu]).strip() if col_tu and pd.notna(r[col_tu]) else ""
                    if not raw_tu or raw_tu.upper() in ['NAN', 'NONE', '']:
                        continue

                    row_text_full = " ".join([str(v).strip().upper() for v in r.values if pd.notna(v)])

                    if "THIẾT BỊ" in row_text_full or "THIẾU" in row_text_full:
                        tb_str = str(r[col_tb]).strip() if col_tb and pd.notna(r[col_tb]) else "Thiết bị chưa rõ tên"
                        sl_val = str(r[col_sl]).strip() if col_sl and pd.notna(r[col_sl]) else ""

                        if sl_val and sl_val.upper() not in ['NAN', 'NONE', '0']:
                            item_detail = f"{tb_str} (SL thiếu: {sl_val})"
                        else:
                            item_detail = tb_str

                        if raw_tu not in hang_thieu_dict:
                            hang_thieu_dict[raw_tu] = []
                        if item_detail not in hang_thieu_dict[raw_tu]:
                            hang_thieu_dict[raw_tu].append(item_detail)

        except Exception as e:
            st.error(f"Lỗi khi đọc file checklist: {str(e)}")

        return hang_thieu_dict, debug_info

    dict_hang_thieu_file, debug_checklist_info = load_checklist_hang_thieu(file_hang_thieu_path)
    
    with st.expander("🔍 **Mở rộng để kiểm tra cấu trúc File Checklist (Debug)**"):
        if debug_checklist_info:
            st.markdown("##### 1. Phân tích nhận diện từ File Checklist:")
            st.json(debug_checklist_info)
            st.markdown("##### 2. Dữ liệu Hàng Thiếu trích xuất được:")
            st.json(dict_hang_thieu_file if dict_hang_thieu_file else "⚠️ Không tìm thấy vị trí bị thiếu thiết bị nào!")
        else:
            st.warning("⚠️ Không tìm thấy dữ liệu từ file checklist.")

    if file_tu_path and os.path.exists(file_tu_path):
        try:
            key_state = f"df_tu_du_an_{selected_tu_file}"
            
            if key_state not in st.session_state:
                if os.path.exists(saved_file_path):
                    df_tu_du_an = pd.read_excel(saved_file_path)
                    st.toast(f"📥 Đã tự động tải lại dữ liệu đã lưu trước đó của dự án: {selected_tu_file}", icon="📂")
                else:
                    df_tu_du_an = pd.read_excel(file_tu_path)
                    
                    col_ma_tu_proj = next((c for c in df_tu_du_an.columns if any(k in str(c).upper() for k in ['TÊN TỦ', 'MÃ TỦ', 'KÝ HIỆU', 'PANEL NAME']) and 'LOẠI' not in str(c).upper()), df_tu_du_an.columns[0])
                    col_so_luong = next((c for c in df_tu_du_an.columns if any(k in str(c).upper() for k in ['SỐ LƯỢNG', 'SL', 'QUANTITY', 'SỐ LƯỢNG TỦ'])), None)
                    
                    if col_so_luong:
                        df_tu_du_an['SL_Thuc_Te'] = pd.to_numeric(df_tu_du_an[col_so_luong], errors='coerce').fillna(1).astype(int)
                    else:
                        df_tu_du_an['SL_Thuc_Te'] = 1

                    trang_thai_list = []
                    hang_thieu_list = []
                    
                    qc_db_map = {str(r['ma_tu']).strip().upper(): r for _, r in df_bien_ban.iterrows()} if not df_bien_ban.empty else {}

                    for idx, row in df_tu_du_an.iterrows():
                        val_tu_ref = str(row[col_ma_tu_proj]).strip()
                        
                        matched_record = next((rec for k_db, rec in qc_db_map.items() if is_smart_match(val_tu_ref, k_db)), None)
                        if matched_record is not None:
                            kq_pass = str(matched_record.get('ket_qua', 'PASS'))
                            status_str = f"🟢 Đã kiểm tra (Đủ thiết bị | Đã Check ({kq_pass}) | Đã Test)"
                        else:
                            status_str = "⏳ Chờ kiểm tra"
                            
                        trang_thai_list.append(status_str)
                        
                        missing_items_display = []
                        for k_chk_tu, list_items in dict_hang_thieu_file.items():
                            if is_smart_match(val_tu_ref, k_chk_tu):
                                missing_items_display.extend(list_items)

                        if missing_items_display:
                            unique_missing = list(dict.fromkeys(missing_items_display))
                            hang_thieu_list.append("; ".join(unique_missing))
                        else:
                            hang_thieu_list.append("NA")

                    df_tu_du_an['Trạng Thái'] = trang_thai_list
                    df_tu_du_an['Hàng Thiếu'] = hang_thieu_list

                col_so_luong_chk = next((c for c in df_tu_du_an.columns if any(k in str(c).upper() for k in ['SỐ LƯỢNG', 'SL', 'QUANTITY', 'SỐ LƯỢNG TỦ'])), None)
                if col_so_luong_chk and 'SL_Thuc_Te' not in df_tu_du_an.columns:
                    df_tu_du_an['SL_Thuc_Te'] = pd.to_numeric(df_tu_du_an[col_so_luong_chk], errors='coerce').fillna(1).astype(int)
                elif 'SL_Thuc_Te' not in df_tu_du_an.columns:
                    df_tu_du_an['SL_Thuc_Te'] = 1

                df_tu_du_an['Hàng Thiếu'] = df_tu_du_an['Hàng Thiếu'].fillna("NA").astype(str)
                df_tu_du_an['Hàng Thiếu'] = df_tu_du_an['Hàng Thiếu'].apply(lambda x: "NA" if x.strip().lower() in ['', 'nan', 'none', 'null', 'nat'] else x)

                st.session_state[key_state] = df_tu_du_an

            df_curr = st.session_state[key_state]
            df_curr['Hàng Thiếu'] = df_curr['Hàng Thiếu'].fillna("NA").astype(str)

            project_title_label = selected_tu_file.replace('.xlsx', '').replace('.xls', '')
            st.subheader(f"📊 Bảng Đối Chiếu Tiến Độ & Hàng Thiếu — [{project_title_label.upper()}]")
            st.caption(f"💡 *Hiển thị toàn bộ danh sách tủ dự án. Mọi chỉnh sửa được lưu độc lập vào file riêng (`{saved_file_name}`), không ảnh hưởng file gốc.*")
            
            tong_tu_du_an = int(df_curr['SL_Thuc_Te'].sum())
            df_da_quet = df_curr[df_curr['Trạng Thái'].str.contains("🟢|Đã kiểm tra|Đã test|Đã Quét|đã kiểm tra|đã test", na=False)]
            tu_da_quet = int(df_da_quet['SL_Thuc_Te'].sum()) if not df_da_quet.empty else 0
            
            def la_co_hang_thieu(val):
                if pd.isna(val):
                    return False
                text = str(val).strip().lower()
                if text in ['', 'none', 'na', 'nan', 'null']:
                    return False
                return True

            df_thieu_thuc_te = df_curr[df_curr['Hàng Thiếu'].apply(la_co_hang_thieu)]
            tu_co_hang_thieu = int(df_thieu_thuc_te['SL_Thuc_Te'].sum()) if not df_thieu_thuc_te.empty else 0

            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("📦 Tổng Số Lượng Tủ Dự Án", f"{tong_tu_du_an} Tủ", delta=f"{len(df_curr)} dòng danh mục")
            col_m2.metric("📸 Đã kiểm tra", f"{tu_da_quet} Tủ", delta=f"{round(tu_da_quet/tong_tu_du_an*100, 1)}%" if tong_tu_du_an > 0 else "0%")
            col_m3.metric("⚠️ Tủ Có Hàng Thiếu / Sai mã (Checklist)", f"{tu_co_hang_thieu} Tủ", delta="Cần bổ sung" if tu_co_hang_thieu > 0 else "Đủ hàng", delta_color="inverse")
            
            st.markdown("##### 🍩 Biểu Đồ Tròn: Phân Bố Trạng Thái Tủ Dự Án")
            df_chart_status = df_curr.groupby('Trạng Thái')['SL_Thuc_Te'].sum().reset_index()
            df_chart_status.columns = ['Trạng Thái', 'Số Lượng']
            
            fig_status = px.pie(
                df_chart_status, 
                names="Trạng Thái", 
                values="Số Lượng",
                hole=0.4
            )
            fig_status.update_traces(textposition='inside', textinfo='percent+label+value')
            fig_status.update_layout(margin=dict(t=20, b=20, l=10, r=10), showlegend=True, height=350)
            st.plotly_chart(fig_status, use_container_width=True)

            existing_statuses = df_curr['Trạng Thái'].astype(str).unique().tolist()
            default_options = [
                "⏳ Chờ kiểm tra",
                "🟢 Đã kiểm tra (Đủ thiết bị | Đã Check | Đã Test)",
                "🟢 Đã kiểm tra (Thiếu vật tư)",
                "✅ Đã test xong"
            ]
            all_status_options = list(dict.fromkeys(existing_statuses + default_options))

            edited_df = st.data_editor(
                df_curr,
                use_container_width=True,
                num_rows="fixed",
                column_config={
                    "Trạng Thái": st.column_config.SelectboxColumn(
                        "Trạng Thái",
                        options=all_status_options,
                        required=True
                    ),
                    "Hàng Thiếu": st.column_config.TextColumn(
                        "Hàng Thiếu (Viết tay / Tự động Excel)",
                        help="Dữ liệu quét tự động từ File Checklist."
                    )
                },
                key=f"data_editor_{selected_tu_file}"
            )

            st.session_state[key_state] = edited_df

            col_save1, col_save2 = st.columns([1, 4])
            with col_save1:
                if st.button("💾 Lưu vào File Riêng", type="primary", key=f"btn_save_rieng_{selected_tu_file}"):
                    try:
                        df_to_save = edited_df.drop(columns=['SL_Thuc_Te'], errors='ignore')
                        df_to_save.to_excel(saved_file_path, index=False)
                        st.success(f"✅ Đã lưu thành công vào file riêng: `{saved_file_name}`!")
                    except Exception as e:
                        st.error(f"❌ Lỗi khi lưu file: {str(e)}")
            
            with col_save2:
                st.caption(f"ℹ️ *Hệ thống sẽ tự động lưu vào file riêng `{saved_file_name}`. File gốc `{selected_tu_file}` sẽ được giữ nguyên an toàn 100%.*")

        except Exception as e:
            st.error(f"Lỗi khi đọc hoặc xử lý file tủ dự án: {str(e)}")

    st.markdown("---")
    st.subheader("📈 Biểu Đồ Tổng Hợp Tiến Độ Tất Cả Các Dự Án")
    st.caption("💡 *Tổng quan nhanh tổng số tủ và số lượng tủ đã kiểm tra của toàn bộ các dự án trong thư mục.*")

    if excel_tu_files:
        summary_data = []
        for f_proj in excel_tu_files:
            f_proj_path = os.path.join(FOLDER_DANH_SACH_TU, f_proj)
            f_saved_path = os.path.join(FOLDER_DANH_SACH_TU, f"KetQua_DoiChieu_{f_proj}")
            
            try:
                if os.path.exists(f_saved_path):
                    df_p = pd.read_excel(f_saved_path)
                else:
                    df_p = pd.read_excel(f_proj_path)
                
                col_sl_p = next((c for c in df_p.columns if any(k in str(c).upper() for k in ['SỐ LƯỢNG', 'SL', 'QUANTITY', 'SỐ LƯỢNG TỦ'])), None)
                if col_sl_p:
                    df_p['SL_Thuc_Te'] = pd.to_numeric(df_p[col_sl_p], errors='coerce').fillna(1).astype(int)
                else:
                    df_p['SL_Thuc_Te'] = 1
                
                tong_tuan_proj = int(df_p['SL_Thuc_Te'].sum())
                
                if 'Trạng Thái' in df_p.columns:
                    df_checked_proj = df_p[df_p['Trạng Thái'].astype(str).str.contains("🟢|Đã kiểm tra|Đã test|Đã Quét|đã kiểm tra|đã test", na=False)]
                    da_check_proj = int(df_checked_proj['SL_Thuc_Te'].sum()) if not df_checked_proj.empty else 0
                else:
                    da_check_proj = 0
                
                proj_name_clean = f_proj.replace('.xlsx', '').replace('.xls', '')
                summary_data.append({
                    "Dự Án": proj_name_clean,
                    "Tổng Số Tủ": tong_tuan_proj,
                    "Đã Kiểm Tra": da_check_proj,
                    "Chưa Kiểm Tra": max(0, tong_tuan_proj - da_check_proj)
                })
            except Exception:
                continue
        
        if summary_data:
            df_summary_all = pd.DataFrame(summary_data)
            
            fig_all_projects = px.bar(
                df_summary_all,
                x="Dự Án",
                y=["Tổng Số Tủ", "Đã Kiểm Tra"],
                barmode="group",
                title="So Sánh Tổng Số Tủ & Số Tủ Đã Kiểm Tra Giữa Các Dự Án",
                labels={"value": "Số Lượng (Tủ)", "variable": "Chỉ Số", "Dự Án": "Tên Dự Án"}
            )
            fig_all_projects.update_layout(margin=dict(t=40, b=40, l=20, r=20), height=420, legend_title="Chú thích")
            st.plotly_chart(fig_all_projects, use_container_width=True)
            
            with st.expander("📋 Xem Bảng Số Liệu Tổng Hợp Chi Tiết Các Dự Án"):
                df_summary_all['Tỷ Lệ Hoàn Thành (%)'] = (df_summary_all['Đã Kiểm Tra'] / df_summary_all['Tổng Số Tủ'] * 100).round(1).astype(str) + '%'
                st.dataframe(df_summary_all, use_container_width=True)
        else:
            st.warning("⚠️ Không thể tổng hợp dữ liệu từ các file dự án.")
    else:
        st.info("ℹ️ Chưa có file dự án nào để tổng hợp.")