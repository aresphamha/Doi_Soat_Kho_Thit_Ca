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

# HÀM XỬ LÝ SỐ AN TOÀN CHO BÁO CÁO
def to_numeric(series):
    if series.dtype == 'object':
        return pd.to_numeric(series.astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    return pd.to_numeric(series, errors='coerce').fillna(0)

@st.cache_data(ttl=600)  # Tự động tải lại sau mỗi 10 phút nếu có người truy cập
def load_data():
    url_meat_fish = "https://docs.google.com/spreadsheets/d/1IpDiHyr262LilJV_gjXUxdtIEh-6peQ4QosRLV0ZRdA/export?format=csv"
    
    def read_csv_with_retry(url, max_retries=3):
        import time
        for i in range(max_retries):
            try:
                response = requests.get(url, timeout=30, verify=False)
                response.raise_for_status()
                content = response.content.decode('utf-8-sig', errors='ignore')
                return pd.read_csv(io.StringIO(content), skiprows=1, dtype=str)
            except Exception as e:
                if i == max_retries - 1:
                    raise e
                time.sleep(2)
                
    df = read_csv_with_retry(url_meat_fish)
    
    # Strip spaces from column names and handle encoding
    df.columns = [str(c).strip() for c in df.columns]
    
    # Rename columns to standard ones if needed
    df.rename(columns={
        'ST': 'ID ST',
        'SL chênh lệch CXD': 'SL chênh lệch CXD'
    }, inplace=True)
    
    # Clean numeric columns
    for col in ['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tổng GT', 'Tổng hao hụt', 'Tổng ST', 'Tổng kho Thá»‹t CÃ¡', 'TÃ\x94NG KHO THá»\x8aT CÃ\x81', 'Tá»\x95ng kho Thá»‹t CÃ¡', 'Tổng chưa xác định', 'Tá»•ng chÆ°a xÃ¡c Ä‘á»‹nh']:
        matched_cols = [c for c in df.columns if col in c or c.startswith('Tổng') or c.startswith('Tá»•ng')]
        for c in matched_cols:
            df[c] = df[c].apply(clean_number)
            
    df['Số lượng chuyển'] = df['Số lượng chuyển'].apply(clean_number)
    df['Số lượng nhận'] = df['Số lượng nhận'].apply(clean_number)
    df['Chênh lệch'] = df['Chênh lệch'].apply(clean_number)
    
    if 'Chi nhánh nhận' in df.columns:
        df['Chi nhánh nhận'] = df['Chi nhánh nhận'].astype(str).str.replace(',', '.', regex=False)
            
    # Lọc lý do chênh lệch
    df['LyDo_HaoHut'] = df['Hao hụt'].astype(str).str.strip().str.lower()
    df['LyDo_SieuThi'] = df['Siêu thị'].astype(str).str.strip().str.lower()
    
    col_kho = [c for c in df.columns if 'KHO TH' in c or 'kho th' in c][0]
    df['LyDo_Kho'] = df[col_kho].astype(str).str.strip().str.lower()
    
    df['Qty_N'] = df['Hạo hụt tự nhiên'].apply(clean_number) if 'Hạo hụt tự nhiên' in df.columns else df['Tổng hao hụt']
    df['Qty_O'] = df['SL trả tồn về ST'].apply(clean_number) if 'SL trả tồn về ST' in df.columns else df['Tổng ST']
    df['Qty_P'] = df['SL chênh lệch CXD'].apply(clean_number) if 'SL chênh lệch CXD' in df.columns else 0.0
    
    # Kết hợp các cột tổng chênh lệch
    col_total_kho = [c for c in df.columns if 'kho Th' in c or 'kho th' in c or 'KHO TH' in c][0]
    col_total_cxd = [c for c in df.columns if 'chưa xác' in c or 'chÆ°a xÃ¡c' in c][0]
    
    df['Hao hụt'] = np.where(df['LyDo_HaoHut'].str.contains('hao hụt'), df['Qty_N'], 0)
    df['BS_ST'] = np.where(df['LyDo_SieuThi'].str.contains('siêu thị'), df['Qty_O'], 0)
    df['Kho_Rau'] = np.where(df['LyDo_Kho'].str.contains('kho thịt cá'), df['Qty_P'], 0)
    df['CXD'] = np.where(df['LyDo_Kho'].str.contains('chưa xác định'), df['Qty_P'], 0)
    
    df['Tổng kho rau'] = df[col_total_kho].apply(clean_number)
    df['Tổng chưa xác định'] = df[col_total_cxd].apply(clean_number)
    
    # Parse dates
    date_col = 'Ngành' if 'Ngành' in df.columns else 'Ngày'
    df['Ngày_parsed'] = pd.to_datetime(df[date_col], format='%d/%m/%Y', errors='coerce')
    idx_null = df['Ngày_parsed'].isna()
    if idx_null.any():
        df.loc[idx_null, 'Ngày_parsed'] = pd.to_datetime(df.loc[idx_null, date_col], format='%m/%d/%Y', errors='coerce')
        
    df['Ngày_str'] = df['Ngày_parsed'].dt.strftime('%d/%m/%Y').fillna(df[date_col])
    df['Ngày'] = df['Ngày_parsed']
    df = df[df['Ngày_parsed'].notna()]
    
    # Categories & SKU
    clv2_col = 'CLV2' if 'CLV2' in df.columns else ('Loại hàng' if 'Loại hàng' in df.columns else 'Unnamed: 0')
    df['CLV2'] = df[clv2_col].fillna('Chưa phân loại')
    
    clv4_col = 'CLV4' if 'CLV4' in df.columns else 'CLV2'
    df['CLV4'] = df[clv4_col].fillna('Chưa phân loại')
    
    ten_hang_col = [c for c in df.columns if 'Tên hàng' in c or 'Tên Hàng' in c][0]
    df['SKU_Full'] = df['Mã hàng'].fillna('').astype(str) + " - " + df[ten_hang_col].fillna('').astype(str)
    
    return df

try:
    df_all = load_data()
    st.success("Tải dữ liệu từ Google Sheets thành công!")
except Exception as e:
    st.error(f"Lỗi tải dữ liệu: {e}")
    st.stop()

# Format màu đỏ cho số chênh lệch
def color_red_for_chenhlech(val):
    color = 'red' if isinstance(val, (int, float)) and val > 0 else ''
    return f'color: {color}'

# Format số theo chuẩn Việt Nam (1.000.000,00)
def format_vn(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float, np.integer, np.floating)):
        if val == int(val):
            return f"{int(val):,}".replace(',', '.')
        else:
            formatted = f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            if formatted.endswith(',00'):
                return formatted[:-3]
            return formatted
    return val

def format_money(val):
    if val >= 1000000:
        return f"{val/1000000:.1f} triệu".replace('.', ',')
    elif val >= 1000:
        return f"{val/1000:.1f} ngàn".replace('.', ',')
    return format_vn(val)

def format_custom_table_with_total(df, name_col, title_prefix):
    if df.empty: return
    
    tong_df = pd.DataFrame(index=[0])
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            tong_df[col] = df[col].sum()
        else:
            tong_df[col] = ''
            
    tuples = []
    for col in df.columns:
        val = tong_df.iloc[0][col]
        if val not in [None, 'Tổng', '', 0] and pd.notna(val):
            if pd.api.types.is_numeric_dtype(type(val)) or isinstance(val, (int, float)):
                total_str = f"🟡 {format_vn(val)}"
            else:
                total_str = f"🟡 {str(val)}"
        else:
            total_str = '⭐ TỔNG' if col == name_col else ''
            
        tuples.append((total_str, col))
        
    df_renamed = df.copy()
    df_renamed.columns = pd.MultiIndex.from_tuples(tuples)
    styler = df_renamed.style.format(format_vn).hide(axis="index")
    display_df_with_download(styler, f"Daily_{title_prefix}")

def compute_daily_summary(df, date_str):
    if df.empty:
        return None
    total_items = int(df['Chênh lệch'].sum())
    total_value = df['Tổng GT'].sum() if 'Tổng GT' in df.columns else 0.0
    
    df_kho_rau = pd.to_numeric(df['Kho_Rau'], errors='coerce').fillna(0)
    df_bs_st = pd.to_numeric(df['BS_ST'], errors='coerce').fillna(0)
    df_hao_hut = pd.to_numeric(df['Hao hụt'], errors='coerce').fillna(0)
    
    processed = int((df_kho_rau + df_bs_st + df_hao_hut).sum())
    returned = int(df_kho_rau.sum())
    created_bs = int(df_bs_st.sum())
    lost = int(df_hao_hut.sum())
    remaining = total_items - processed
    
    df_cxd = pd.to_numeric(df['CXD'], errors='coerce').fillna(0)
    xuly_status = df['Xử lý'].astype(str).str.strip().str.lower() if 'Xử lý' in df.columns else pd.Series('đang xử lý', index=df.index)
    pending = int(np.where(xuly_status == 'hoàn thành', df_cxd, 0).sum())
    unprocessed = remaining - pending
    if unprocessed < 0: unprocessed = 0
    
    cat_summary = {}
    for cat in df['CLV2'].dropna().unique():
        df_cat = df[df['CLV2'] == cat]
        c_items = int(df_cat['Chênh lệch'].sum())
        c_val = df_cat['Tổng GT'].sum() if 'Tổng GT' in df_cat.columns else 0.0
        
        c_kr = pd.to_numeric(df_cat['Kho_Rau'], errors='coerce').fillna(0)
        c_st = pd.to_numeric(df_cat['BS_ST'], errors='coerce').fillna(0)
        c_hh = pd.to_numeric(df_cat['Hao hụt'], errors='coerce').fillna(0)
        
        c_ret = int(c_kr.sum())
        c_bs = int(c_st.sum())
        c_lost = int(c_hh.sum())
        c_proc = c_ret + c_bs + c_lost
        c_rem = c_items - c_proc
        
        c_cxd = pd.to_numeric(df_cat['CXD'], errors='coerce').fillna(0)
        c_xuly = df_cat['Xử lý'].astype(str).str.strip().str.lower() if 'Xử lý' in df_cat.columns else pd.Series('đang xử lý', index=df_cat.index)
        c_pending = int(np.where(c_xuly == 'hoàn thành', c_cxd, 0).sum())
        c_unprocessed = c_rem - c_pending
        if c_unprocessed < 0: c_unprocessed = 0
        
        cause_ratio = c_ret / c_items if c_items > 0 else 0
        
        cat_summary[cat] = {
            "items": c_items,
            "value": c_val,
            "processed": c_proc,
            "return": c_ret,
            "bs": c_bs,
            "lost": c_lost,
            "remaining": c_rem,
            "pending": c_pending,
            "unprocessed": c_unprocessed,
            "cause_ratio": cause_ratio
        }
    return {
        "date": date_str,
        "total_items": total_items,
        "total_value": total_value,
        "processed": processed,
        "return": returned,
        "bs": created_bs,
        "lost": lost,
        "remaining": remaining,
        "pending": pending,
        "unprocessed": unprocessed,
        "cat_summary": cat_summary
    }

# ==========================================
# GIAO DIỆN CHIA TAB
# ==========================================
tab_main, tab_daily, tab_dc = st.tabs(["📊 Báo Cáo Tổng Quan", "📈 Báo Cáo Năng Suất Daily", "👨‍🔧 Tiến Độ DC Phản Hồi"])

with tab_main:
    # Nút cập nhật
    if st.button('🔄 Cập nhật dữ liệu mới nhất'):
        st.cache_data.clear()
        st.rerun()
        
    df_active = df_all.copy()

    # Process Dataframes
    pivot_ngay_sum = df_active.groupby('Ngày_str')[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tổng GT', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']].sum()
    pivot_ngay_count = df_active[df_active['Chênh lệch'].abs() > 0].groupby('Ngày_str').size().rename('SL line chênh lệch')
    pivot_ngay = pivot_ngay_sum.join(pivot_ngay_count).fillna(0).reset_index()
    
    pivot_ngay['Ngày_dt'] = pd.to_datetime(pivot_ngay['Ngày_str'], format='%d/%m/%Y', errors='coerce')
    pivot_ngay = pivot_ngay.sort_values(by='Ngày_dt').drop(columns=['Ngày_dt'])

    tong_row_ngay = pivot_ngay.sum(numeric_only=True).to_frame().T
    tong_row_ngay['Ngày_str'] = 'Tổng'

    pivot_ngay.rename(columns={
        'Tổng GT': 'Giá trị chênh lệch (VNĐ)',
        'BS_ST': 'SL đã tạo bs cho ST',
        'Kho_Rau': 'SL đã xác nhận được trả kho thịt cá',
        'Hao hụt': 'Số lượng hao hụt',
        'CXD': 'Số lượng chưa xác định'
    }, inplace=True)
    tong_row_ngay.rename(columns={
        'Tổng GT': 'Giá trị chênh lệch (VNĐ)',
        'BS_ST': 'SL đã tạo bs cho ST',
        'Kho_Rau': 'SL đã xác nhận được trả kho thịt cá',
        'Hao hụt': 'Số lượng hao hụt',
        'CXD': 'Số lượng chưa xác định'
    }, inplace=True)

    # Theo ngày (Giá trị)
    pivot_ngay_val = df_active.groupby('Ngày_str')[['Tổng GT', 'Tổng ST', 'Tổng kho rau', 'Tổng chưa xác định']].sum().reset_index()
    pivot_ngay_val['Ngày_dt'] = pd.to_datetime(pivot_ngay_val['Ngày_str'], format='%d/%m/%Y', errors='coerce')
    pivot_ngay_val = pivot_ngay_val.sort_values(by='Ngày_dt').drop(columns=['Ngày_dt'])

    tong_row_ngay_val = pivot_ngay_val.sum(numeric_only=True).to_frame().T
    if not tong_row_ngay_val.empty: tong_row_ngay_val['Ngày_str'] = 'Tổng'

    pivot_ngay_val.rename(columns={
        'Tổng GT': 'Giá trị chênh lệch (VNĐ)',
        'Tổng ST': 'Giá trị đã tạo bs cho ST (VNĐ)',
        'Tổng kho rau': 'Giá trị đã trả Kho Thịt Cá (VNĐ)',
        'Tổng chưa xác định': 'Giá trị chưa xác định (VNĐ)'
    }, inplace=True)
    if not tong_row_ngay_val.empty:
        tong_row_ngay_val.rename(columns={
            'Tổng GT': 'Giá trị chênh lệch (VNĐ)',
            'Tổng ST': 'Giá trị đã tạo bs cho ST (VNĐ)',
            'Tổng kho rau': 'Giá trị đã trả Kho Thịt Cá (VNĐ)',
            'Tổng chưa xác định': 'Giá trị chưa xác định (VNĐ)'
        }, inplace=True)

    # Theo CLV2
    pivot_clv2_sum = df_active.groupby('CLV2', dropna=False)[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch']].sum()
    pivot_clv2_count = df_active[df_active['Chênh lệch'].abs() > 0].groupby('CLV2', dropna=False).size().rename('Số lượng line')
    pivot_clv2 = pivot_clv2_sum.join(pivot_clv2_count).fillna(0).reset_index()
    pivot_clv2['Số lượng line'] = pivot_clv2['Số lượng line'].astype(int)
    pivot_clv2 = pivot_clv2.sort_values(by='Chênh lệch', ascending=False)
    
    tong_row_clv2 = pivot_clv2.sum(numeric_only=True).to_frame().T
    tong_row_clv2['CLV2'] = 'Tổng'

    # Top 5 CLV4
    clv4_sum = df_active.groupby('CLV4', dropna=False)[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch']].sum().reset_index()
    clv4_sum['Abs_ChenhLech'] = clv4_sum['Chênh lệch'].abs()
    pivot_clv4 = clv4_sum.sort_values(by='Abs_ChenhLech', ascending=False).drop(columns=['Abs_ChenhLech']).head(5)

    # Header Metrics
    st.write("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Tổng số lượng chuyển", format_vn(df_active['Số lượng chuyển'].sum()))
    with col2:
        st.metric("Tổng số lượng nhận", format_vn(df_active['Số lượng nhận'].sum()))
    with col3:
        st.metric("TỔNG CHÊNH LỆCH", format_vn(df_active['Chênh lệch'].sum()))

    def create_multiindex_headers(df, tong_df):
        if df.empty or tong_df.empty: return df
        tuples = []
        for i, col in enumerate(df.columns):
            if col in tong_df.columns:
                val = tong_df.iloc[0][col]
                if val not in [None, 'Tổng', '', 0] and pd.notna(val):
                    if pd.api.types.is_numeric_dtype(type(val)) or isinstance(val, (int, float)):
                        tuples.append((f"🟡 {format_vn(val)}", col))
                    else:
                        tuples.append((f"🟡 {str(val)}", col))
                else:
                    tuples.append(('⭐ TỔNG' if i == 0 else '', col))
            else:
                tuples.append(('⭐ TỔNG' if i == 0 else '', col))
        df_new = df.copy()
        df_new.columns = pd.MultiIndex.from_tuples(tuples)
        return df_new

    pivot_ngay_renamed = create_multiindex_headers(pivot_ngay, tong_row_ngay)
    pivot_ngay_val_renamed = create_multiindex_headers(pivot_ngay_val, tong_row_ngay_val)
    pivot_clv2_renamed = create_multiindex_headers(pivot_clv2, tong_row_clv2)

    # Layout các bảng
    st.write("---")
    st.subheader("📅 1. TỔNG HỢP THEO TỪNG NGÀY")
    tab_ngay_qty, tab_ngay_val = st.tabs(["📊 Số lượng (Từng Ngày)", "💰 Giá trị (Từng Ngày)"])

    with tab_ngay_qty:
        display_df_with_download(pivot_ngay_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in pivot_ngay_renamed.columns if 'Chênh lệch' in c[1]]), "Tong_Hop_Theo_Ngay_So_Luong")

    with tab_ngay_val:
        display_df_with_download(pivot_ngay_val_renamed.style.format(format_vn), "Tong_Hop_Theo_Ngay_Gia_Tri")

    st.write("---")
    col4, col5 = st.columns(2)
    with col4:
        st.subheader("🔥 2. TOP 5 CATE CHÊNH LỆCH LỚN NHẤT")
        display_df_with_download(pivot_clv4.style.format(format_vn).map(color_red_for_chenhlech, subset=['Chênh lệch']), "Top_5_CLV4")
    with col5:
        st.subheader("📦 3. TỔNG HỢP THEO NGÀNH HÀNG (CLV2)")
        display_df_with_download(pivot_clv2_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in pivot_clv2_renamed.columns if 'Chênh lệch' in c[1]]), "Tong_Hop_CLV2")

# ==========================================
# TRANG 2: BÁO CÁO DAILY MỚI
# ==========================================
with tab_daily:
    st.header("Báo Cáo Năng Suất Chi Tiết Mỗi Ngày")
    
    unique_daily_dates = df_all.sort_values(by='Ngày')['Ngày_str'].dropna().unique().tolist()
    if unique_daily_dates:
        options = ["Tất cả các ngày"] + unique_daily_dates
        selected_daily_date = st.selectbox("📅 Chọn ngày báo cáo (Daily):", options, index=len(options)-1)
        df_filtered = df_all if selected_daily_date == "Tất cả các ngày" else df_all[df_all['Ngày_str'] == selected_daily_date].copy()
    else:
        df_filtered = pd.DataFrame()
    
    def calculate_daily_metrics(data, group_by_col='CLV2'):
        if data.empty: return pd.DataFrame()
        
        data['Số lượng chuyển_clean'] = to_numeric(data['Số lượng chuyển'])
        data['Chênh lệch_clean'] = to_numeric(data['Chênh lệch'])
        data['Tổng_GT_num'] = to_numeric(data['Tổng GT']) if 'Tổng GT' in data.columns else 0.0
        data['Tổng_ST_num'] = to_numeric(data['Tổng ST']) if 'Tổng ST' in data.columns else 0.0
        data['Tổng_Kho_Rau_num'] = to_numeric(data['Tổng kho rau'])
        data['Tổng_Hao_Hut_num'] = to_numeric(data['Tổng hao hụt']) if 'Tổng hao hụt' in data.columns else 0.0
        data['Tổng_CXD_num'] = to_numeric(data['Tổng chưa xác định'])
        
        trang_thai = data['Xử lý'].astype(str).str.strip().str.lower() if 'Xử lý' in data.columns else pd.Series('đang xử lý', index=data.index)
        
        data['CXD_DangXuLy'] = np.where(trang_thai == 'hoàn thành', data['CXD'], 0)
        data['GT_CXD_DangXuLy'] = np.where(trang_thai == 'hoàn thành', data['Tổng_CXD_num'], 0)
        
        data['CXD_WriteOff'] = np.where(trang_thai == 'đang xử lý', data['CXD'], 0)
        data['CXD_WriteOff_Val'] = np.where(trang_thai == 'đang xử lý', data['Tổng_CXD_num'], 0)
        
        data['CXD_ChuaXuLy'] = np.where((trang_thai != 'hoàn thành') & (trang_thai != 'đang xử lý'), data['CXD'], 0)
        data['GT_CXD_ChuaXuLy'] = np.where((trang_thai != 'hoàn thành') & (trang_thai != 'đang xử lý'), data['Tổng_CXD_num'], 0)
        
        data['Is_Da_Xu_Ly'] = (trang_thai == 'hoàn thành')
        data['ST_Chenh_Lech'] = np.where(data['Chênh lệch_clean'].abs() > 0, data['ID ST'], np.nan)
        
        grouped = data.groupby(group_by_col, dropna=False).agg(
            Số_lượng_chuyển=('Số lượng chuyển_clean', 'sum'),
            Số_lượng_chênh_lệch=('Chênh lệch_clean', 'sum'),
            Số_lượng_ST_chênh_lệch=('ST_Chenh_Lech', 'nunique'),
            Số_lượng_line_chênh_lệch=('Mã hàng', 'count'),
            Số_lượng_line_hao_hụt=('Hao hụt', lambda x: (x > 0).sum()),
            Số_lượng_line_đã_xử_lý=('Is_Da_Xu_Ly', 'sum'),
            Số_lượng_hao_hụt=('Hao hụt', 'sum'),
            Số_lượng_bs_ST=('BS_ST', 'sum'),
            SL_bs_kho_rau=('Kho_Rau', 'sum'),
            Số_lượng_đang_xử_lý=('CXD_DangXuLy', 'sum'),
            Số_lượng_chưa_xác_định=('CXD_ChuaXuLy', 'sum'),
            Số_lượng_write_off=('CXD_WriteOff', 'sum'),
            Giá_trị_write_off=('CXD_WriteOff_Val', 'sum'),
            Giá_trị_chênh_lệch=('Tổng_GT_num', 'sum'),
            Giá_trị_hao_hụt=('Tổng_Hao_Hut_num', 'sum'),
            Giá_trị_bs_ST=('Tổng_ST_num', 'sum'),
            Giá_trị_bs_kho_rau=('Tổng_Kho_Rau_num', 'sum'),
            Giá_trị_đang_xử_lý=('Giá_trị_đang_xử_lý' if 'Giá_trị_đang_xử_lý' in data.columns else 'GT_CXD_DangXuLy', 'sum'),
            Giá_trị_chưa_xác_định=('Giá_trị_chưa_xác_định' if 'Giá_trị_chưa_xác_định' in data.columns else 'GT_CXD_ChuaXuLy', 'sum')
        ).reset_index()
        
        grouped['Tỷ lệ line đã xử lý'] = ((grouped['Số_lượng_line_đã_xử_lý'] + grouped['Số_lượng_line_hao_hụt']) / grouped['Số_lượng_line_chênh_lệch'] * 100).round(2).astype(str) + '%'
        grouped['Tỷ lệ hao hụt'] = np.where(grouped['Số_lượng_chuyển'] > 0, (grouped['Số_lượng_hao_hụt'] / grouped['Số_lượng_chuyển'] * 100).round(2).astype(str) + '%', '0.0%')
        
        grouped = grouped.rename(columns={
            'Số_lượng_chuyển': 'SL chuyển',
            'Số_lượng_chênh_lệch': 'SL chênh lệch',
            'Số_lượng_ST_chênh_lệch': 'SL ST chênh lệch',
            'Số_lượng_line_chênh_lệch': 'SL line chênh lệch',
            'Số_lượng_line_hao_hụt': 'SL line hao hụt',
            'Số_lượng_line_đã_xử_lý': 'SL line đã xử lý',
            'Số_lượng_hao_hụt': 'Số lượng hao hụt',
            'Số_lượng_bs_ST': 'SL bs ST',
            'SL_bs_kho_rau': 'SL bs kho rau',
            'Số_lượng_đang_xử_lý': 'Đang xử lý',
            'Số_lượng_chưa_xác_định': 'Chưa xử lý',
            'Số_lượng_write_off': 'Không xử lý (WRITE OFF)',
            'Giá_trị_write_off': 'Giá trị WRITE OFF',
            'Giá_trị_chênh_lệch': 'GT chênh lệch',
            'Giá_trị_hao_hụt': 'GT hao hụt',
            'Giá_trị_bs_ST': 'GT bs ST',
            'Giá_trị_bs_kho_rau': 'GT bs kho rau',
            'Giá_trị_đang_xử_lý': 'GT Đang xử lý',
            'Giá_trị_chưa_xác_định': 'GT Chưa xử lý'
        })
        
        return grouped

    def display_daily_table(df, cols, title_prefix, group_by_col='CLV2'):
        if df.empty: return
        df_show = df[cols].copy()
        tong_df = pd.DataFrame(index=[0])
        for col in cols:
            if col == group_by_col:
                tong_df[col] = 'Tổng'
            elif col == 'Tỷ lệ line đã xử lý':
                sum_line_xl = df['SL line đã xử lý'].sum()
                sum_line_hh = df['SL line hao hụt'].sum()
                sum_line_cl = df['SL line chênh lệch'].sum()
                tong_df[col] = f"{round(((sum_line_xl + sum_line_hh) / sum_line_cl) * 100, 2)}%" if sum_line_cl > 0 else '0.0%'
            elif col == 'Tỷ lệ hao hụt':
                sum_hh = df['Số lượng hao hụt'].sum()
                sum_ch = df['SL chuyển'].sum()
                tong_df[col] = f"{round((sum_hh / sum_ch) * 100, 2)}%" if sum_ch > 0 else '0.0%'
            elif pd.api.types.is_numeric_dtype(df[col]):
                tong_df[col] = df[col].sum()
            else:
                tong_df[col] = ''
                
        tuples = []
        for col in cols:
            val = tong_df.iloc[0][col]
            if val not in [None, 'Tổng', '', 0] and pd.notna(val):
                total_str = f"🟡 {format_vn(val)}"
            else:
                total_str = '⭐ TỔNG' if col == group_by_col else ''
            tuples.append((total_str, col))
            
        df_renamed = df_show.copy()
        df_renamed.columns = pd.MultiIndex.from_tuples(tuples)
        display_df_with_download(df_renamed.style.format(format_vn), f"Daily_{title_prefix}")

    st.subheader("Bảng 1: Đánh giá nhanh tình hình xử lý")
    df_b1 = calculate_daily_metrics(df_filtered)
    cols = ['CLV2', 'SL chuyển', 'SL chênh lệch', 'SL ST chênh lệch', 'SL line chênh lệch', 'SL line hao hụt', 'SL line đã xử lý', 'Tỷ lệ line đã xử lý', 'Số lượng hao hụt', 'Tỷ lệ hao hụt', 'SL bs ST', 'SL bs kho rau', 'Đang xử lý', 'Chưa xử lý', 'Không xử lý (WRITE OFF)']
    
    display_daily_table(df_b1, cols, "Bang_1")
    
    # Nhận xét nhanh tự động
    summary_daily = compute_daily_summary(df_filtered, selected_daily_date)
    if summary_daily:
        st.markdown(f"**BÁO CÁO CHÊNH LỆCH ĐỐI SOÁT NGÀY {summary_daily['date']}**")
        st.write(f"**Tổng: Lệch {summary_daily['total_items']} items (~{format_money(summary_daily['total_value'])} VNĐ).**")
        
        pct_done = round(summary_daily['processed'] / summary_daily['total_items'] * 100) if summary_daily['total_items'] else 0
        st.write(f"Đã xử lý: {summary_daily['processed']} items ({pct_done}%)")
        st.write(f"- Trả về Kho thịt cá {summary_daily['return']} items")
        st.write(f"- Tạo bs ST {summary_daily['bs']} items")
        st.write(f"- Hao hụt {summary_daily['lost']} items.")
        
        pct_remain = round(summary_daily['remaining'] / summary_daily['total_items'] * 100) if summary_daily['total_items'] else 0
        st.write(f"Tồn lại: {summary_daily['remaining']} items ({pct_remain}%)")

# ==========================================
# TRANG 3: TIẾN ĐỘ DC PHẢN HỒI
# ==========================================
with tab_dc:
    st.header("Tiến Độ Phản Hồi Từ Trung Tâm Phân Phối (DC)")
    
    df_dc_active = df_all[df_all['Kho_Rau'] > 0].copy()
    if df_dc_active.empty:
        st.info("Không có dữ liệu tiến độ DC phản hồi trong kỳ báo cáo này.")
    else:
        df_dc_active['DC_Xac_Nhan'] = df_dc_active['DC xác nhận'].fillna('Chưa xác nhận').replace('', 'Chưa xác nhận')
        df_dc_active['Nhom_Loi'] = df_dc_active['Lỗi'].fillna('Không phân loại').replace('', 'Không phân loại')
        
        st.subheader("Thống kê chi tiết phản hồi")
        pivot_loi = pd.pivot_table(df_dc_active, values='Kho_Rau', index='Nhom_Loi', columns='DC_Xac_Nhan', aggfunc='sum', fill_value=0).reset_index()
        display_df_with_download(pivot_loi, "Tien_do_phat_sinh_DC")
