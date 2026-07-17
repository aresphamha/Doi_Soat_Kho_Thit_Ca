import streamlit as st
import pandas as pd
import numpy as np
import io
import requests

def get_excel_bytes(df):
    output = io.BytesIO()
    df_to_export = df.copy()
    if isinstance(df_to_export.columns, pd.MultiIndex):
        df_to_export.columns = [' - '.join(str(c) for c in col if c).strip() for col in df_to_export.columns.values]
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_to_export.to_excel(writer, index=False)
    return output.getvalue()

def display_df_with_download(styled_df, filename, height=None):
    if height:
        st.dataframe(styled_df, use_container_width=True, height=height)
    else:
        st.dataframe(styled_df, use_container_width=True)
    df_raw = styled_df.data if hasattr(styled_df, 'data') else styled_df
    try:
        excel_data = get_excel_bytes(df_raw)
        st.download_button(label="📥 Tải xuống Excel", data=excel_data, file_name=f"{filename}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=filename)
    except Exception as e:
        st.error(f"Lỗi xuất Excel: {e}")

# Cấu hình trang web
st.set_page_config(page_title="Dashboard Đối Soát Kho Thịt Cá", page_icon="🥩", layout="wide")

st.title("🥩 Báo Cáo Đối Soát Kho Thịt Cá")
st.markdown("Dữ liệu tự động cập nhật từ Hệ thống Google Sheets")

# Hàm làm sạch số liệu thông minh (Xử lý lẫn lộn định dạng Anh/Việt)
def clean_number(x):
    if pd.isna(x) or x == '':
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        x = x.strip()
        if x == '':
            return 0.0
        num_dots = x.count('.')
        num_commas = x.count(',')
        
        if num_dots > 0 and num_commas > 0:
            last_dot = x.rfind('.')
            last_comma = x.rfind(',')
            if last_comma > last_dot: # Format VN: 1.234,50
                x = x.replace('.', '').replace(',', '.')
            else: # Format EN: 1,234.50
                x = x.replace(',', '')
        elif num_commas > 0:
            parts = x.split(',')
            if num_commas > 1:
                x = x.replace(',', '')
            else:
                x = x.replace(',', '.') # VN format: comma is decimal
        elif num_dots > 0:
            parts = x.split('.')
            if num_dots > 1:
                x = x.replace('.', '')
            else:
                if len(parts[1]) == 3 and parts[0] not in ['0', '-0']:
                    x = x.replace('.', '') # VD: 16.000 -> 16000
                else:
                    pass # VD: 5.5 -> 5.5
    try:
        return float(x)
    except:
        return 0.0

@st.cache_data(ttl=600)  # Tự động tải lại sau mỗi 10 phút nếu có người truy cập
def load_data():
    url_meat_fish = "https://docs.google.com/spreadsheets/d/1IpDiHyr262LilJV_gjXUxdtIEh-6peQ4QosRLV0ZRdA/export?format=csv"
    
    def read_csv_with_retry(url, max_retries=3):
        import time
        for i in range(max_retries):
            try:
                response = requests.get(url, timeout=30, verify=False)
                response.raise_for_status()
                # Decode bytes to string correctly and ignore byte order mark (BOM)
                content = response.content.decode('utf-8-sig', errors='ignore')
                return pd.read_csv(io.StringIO(content), skiprows=1, dtype=str)
            except Exception as e:
                if i == max_retries - 1:
                    raise e
                time.sleep(2)
                
    df = read_csv_with_retry(url_meat_fish)
    
    # Strip spaces from column names
    df.columns = [str(c).strip() for c in df.columns]
    
    # Rename columns to standard ones if needed
    df.rename(columns={
        'ST': 'ID ST',
        'SL chênh lệch CXD': 'SL chênh lệch CXD'
    }, inplace=True)
    
    # Clean numeric columns
    for col in ['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tổng GT', 'Tổng hao hụt', 'Tổng ST', 'Tổng kho Thịt Cá', 'Tổng chưng chưa xác định']:
        # Note: the check handles possible typos in the sheet like 'Tổng chưng chưa xác định'
        matched_cols = [c for c in df.columns if col in c or c.startswith('Tổng')]
        for c in matched_cols:
            df[c] = df[c].apply(clean_number)
            
    df['Số lượng chuyển'] = df['Số lượng chuyển'].apply(clean_number)
    df['Số lượng nhận'] = df['Số lượng nhận'].apply(clean_number)
    df['Chênh lệch'] = df['Chênh lệch'].apply(clean_number)
    
    if 'Chi nhánh nhận' in df.columns:
        df['Chi nhánh nhận'] = df['Chi nhánh nhận'].astype(str).str.replace(',', '.', regex=False)
            
    # Lọc lý do
    df['LyDo_HaoHut'] = df['Hao hụt'].astype(str).str.strip().str.lower()
    df['LyDo_SieuThi'] = df['Siêu thị'].astype(str).str.strip().str.lower()
    
    col_kho = [c for c in df.columns if 'KHO TH' in c][0]
    df['LyDo_Kho'] = df[col_kho].astype(str).str.strip().str.lower()
    
    df['Qty_N'] = df['Hạo hụt tự nhiên'].apply(clean_number) if 'Hạo hụt tự nhiên' in df.columns else df['Tổng hao hụt']
    df['Qty_O'] = df['SL trả tồn về ST'].apply(clean_number) if 'SL trả tồn về ST' in df.columns else df['Tổng ST']
    df['Qty_P'] = df['SL chênh lệch CXD'].apply(clean_number) if 'SL chênh lệch CXD' in df.columns else 0.0
    
    df['Hao hụt tự nhiên'] = np.where(df['LyDo_HaoHut'].str.contains('hao hụt'), df['Qty_N'], 0)
    df['Lỗi Siêu thị'] = np.where(df['LyDo_SieuThi'].str.contains('siêu thị'), df['Qty_O'], 0)
    df['Lỗi Kho thịt cá'] = np.where(df['LyDo_Kho'].str.contains('kho thịt cá'), df['Qty_P'], 0)
    df['Chưa xác định'] = np.where(df['LyDo_Kho'].str.contains('chưa xác định'), df['Qty_P'], 0)
    
    # Parse dates
    # We try both %d/%m/%Y and %m/%d/%Y formats
    df['Ngày_parsed'] = pd.to_datetime(df['Ngành' if 'Ngành' in df.columns else 'Ngày'], format='%d/%m/%Y', errors='coerce')
    idx_null = df['Ngày_parsed'].isna()
    if idx_null.any():
        df.loc[idx_null, 'Ngày_parsed'] = pd.to_datetime(df.loc[idx_null, 'Ngành' if 'Ngành' in df.columns else 'Ngày'], format='%m/%d/%Y', errors='coerce')
        
    df['Ngày_str'] = df['Ngày_parsed'].dt.strftime('%d/%m/%Y').fillna(df['Ngành' if 'Ngành' in df.columns else 'Ngày'])
    df = df[df['Ngày_parsed'].notna()]
    
    # SKU representation
    ten_hang_col = [c for c in df.columns if 'Tên hàng' in c or 'Tên Hàng' in c][0]
    df['SKU_Full'] = df['Mã hàng'].fillna('').astype(str) + " - " + df[ten_hang_col].fillna('').astype(str)
    
    return df

try:
    df_all = load_data()
    st.success("Tải dữ liệu thành công!")
except Exception as e:
    st.error(f"Lỗi tải dữ liệu: {e}")
    st.stop()

# ----------------- SIDEBAR FILTER -----------------
st.sidebar.header("📊 Bộ Lọc Dữ Liệu")

# Chọn khoảng ngày
dates = sorted(df_all['Ngày_parsed'].unique())
if len(dates) > 0:
    start_date, end_date = st.sidebar.select_slider(
        "Chọn khoảng thời gian",
        options=dates,
        value=(dates[0], dates[-1]),
        format_func=lambda x: pd.to_datetime(x).strftime('%d/%m/%Y')
    )
    df_filtered = df_all[(df_all['Ngày_parsed'] >= start_date) & (df_all['Ngày_parsed'] <= end_date)].copy()
else:
    df_filtered = df_all.copy()

# Lọc siêu thị
stores = sorted(df_filtered['Chi nhánh nhận'].dropna().unique())
selected_stores = st.sidebar.multiselect("Chọn Siêu Thị", options=stores, default=[])
if selected_stores:
    df_filtered = df_filtered[df_filtered['Chi nhánh nhận'].isin(selected_stores)]

# Lọc lý do lỗi
error_types = ["Hao hụt tự nhiên", "Lỗi Siêu thị", "Lỗi Kho thịt cá", "Chưa xác định"]
selected_errors = st.sidebar.multiselect("Lọc theo nhóm lỗi thực tế", options=error_types, default=[])
if selected_errors:
    condition = pd.Series(False, index=df_filtered.index)
    for err in selected_errors:
        condition = condition | (df_filtered[err] != 0)
    df_filtered = df_filtered[condition]

# ----------------- MAIN METRICS -----------------
total_transfers = len(df_filtered)
total_qty_trans = df_filtered['Số lượng chuyển'].sum()
total_qty_recv = df_filtered['Số lượng nhận'].sum()
total_diff = df_filtered['Chênh lệch'].sum()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Tổng Số Dòng", f"{total_transfers:,}")
m2.metric("Số Lượng Chuyển", f"{total_qty_trans:,.2f}")
m3.metric("Số Lượng Nhận", f"{total_qty_recv:,.2f}")
m4.metric("Chênh Lệch", f"{total_diff:,.2f}", delta_color="inverse")

# ----------------- DETAIL TABS -----------------
tab1, tab2, tab3 = st.tabs(["📊 Báo Cáo Siêu Thị", "📅 Báo Cáo Theo Ngày", "🔍 Chi Tiết Giao Dịch"])

with tab1:
    st.subheader("Báo cáo chênh lệch và nguồn lỗi theo từng Siêu Thị")
    
    df_store = df_filtered.groupby('Chi nhánh nhận').agg(
        SL_Chuyen=('Số lượng chuyển', 'sum'),
        SL_Nhan=('Số lượng nhận', 'sum'),
        Chenh_Lech=('Chênh lệch', 'sum'),
        Hao_Hut=('Hao hụt tự nhiên', 'sum'),
        Loi_ST=('Lỗi Siêu thị', 'sum'),
        Loi_Kho=('Lỗi Kho thịt cá', 'sum'),
        Chua_XD=('Chưa xác định', 'sum')
    ).reset_index()
    
    df_store.columns = ['Siêu thị', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt tự nhiên', 'Lỗi Siêu thị', 'Lỗi Kho thịt cá', 'Chưa xác định']
    display_df_with_download(df_store, "Bao_cao_theo_Sieu_Thi")

with tab2:
    st.subheader("Báo cáo chênh lệch và nguồn lỗi theo Ngày")
    
    df_day = df_filtered.groupby('Ngày_str').agg(
        SL_Chuyen=('Số lượng chuyển', 'sum'),
        SL_Nhan=('Số lượng nhận', 'sum'),
        Chenh_Lech=('Chênh lệch', 'sum'),
        Hao_Hut=('Hao hụt tự nhiên', 'sum'),
        Loi_ST=('Lỗi Siêu thị', 'sum'),
        Loi_Kho=('Lỗi Kho thịt cá', 'sum'),
        Chua_XD=('Chưa xác định', 'sum')
    ).reset_index()
    
    df_day.columns = ['Ngày', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt tự nhiên', 'Lỗi Siêu thị', 'Lỗi Kho thịt cá', 'Chưa xác định']
    display_df_with_download(df_day, "Bao_cao_theo_Ngay")

with tab3:
    st.subheader("Chi tiết từng dòng giao dịch bị lệch")
    
    df_detail_view = df_filtered[['Ngày_str', 'Chi nhánh nhận', 'SKU_Full', 'ĐVT', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt tự nhiên', 'Lỗi Siêu thị', 'Lỗi Kho thịt cá', 'Chưa xác định', 'PT chuyển', 'Mã thùng', 'Lỗi', 'NOTE']].copy()
    df_detail_view.columns = ['Ngày', 'Siêu thị', 'Mã & Tên hàng', 'ĐVT', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt tự nhiên', 'Lỗi Siêu thị', 'Lỗi Kho thịt cá', 'Chưa xác định', 'PT chuyển', 'Mã thùng', 'Lỗi Hệ thống', 'Ghi chú']
    
    display_df_with_download(df_detail_view, "Chi_tiet_lech_thit_ca")
