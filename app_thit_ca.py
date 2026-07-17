import streamlit as st
import pandas as pd
import numpy as np
import io
import requests

# Cấu hình trang web
st.set_page_config(page_title="Dashboard Đối Soát Kho Thịt Cá", page_icon="🥩", layout="wide")

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

st.title("🥩 Báo Cáo Đối Soát Kho Thịt Cá")
st.markdown("Dữ liệu tự động cập nhật từ Hệ thống Google Sheets")

# Hàm làm sạch số lượng chênh lệch (Đơn vị nhỏ: cái, kg, hộp)
def clean_qty(x):
    if pd.isna(x) or x == '':
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        x = x.strip()
        if x == '':
            return 0.0
        # Ở cột số lượng, dấu phẩy thường là phân cách thập phân (VD: 5,5 -> 5.5) hoặc dấu chấm cũng vậy (VD: -1.000 -> -1.0)
        # Không có số lượng hàng nghìn trên 1 dòng ở kho thịt cá, nên ta quy về dạng chuẩn
        x = x.replace(',', '.')
        # Nếu có nhiều dấu chấm (lỗi định dạng), chỉ giữ lại dấu chấm cuối cùng
        if x.count('.') > 1:
            parts = x.split('.')
            x = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(x)
    except:
        return 0.0

# Hàm làm sạch giá trị tiền tệ (Đơn vị lớn: VNĐ)
def clean_val(x):
    if pd.isna(x) or x == '':
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        x = x.strip()
        if x == '':
            return 0.0
        # Xử lý tiền tệ VNĐ (VD: -2.045.634 hoặc -2,045,634.00 hoặc -13.914.047,04)
        num_dots = x.count('.')
        num_commas = x.count(',')
        if num_dots > 0 and num_commas > 0:
            last_dot = x.rfind('.')
            last_comma = x.rfind(',')
            if last_comma > last_dot: # Định dạng VN: 1.234.567,89
                x = x.replace('.', '').replace(',', '.')
            else: # Định dạng EN: 1,234,567.89
                x = x.replace(',', '')
        elif num_dots > 1: # Nhiều dấu chấm: 1.234.567 -> bỏ chấm
            x = x.replace('.', '')
        elif num_commas > 1: # Nhiều dấu phẩy: 1,234,567 -> bỏ phẩy
            x = x.replace(',', '')
        elif num_dots == 1:
            # Nếu chỉ có 1 dấu chấm, xem nó là thập phân hay hàng nghìn (VD: -13914047.04 hay 16.000 VNĐ)
            parts = x.split('.')
            if len(parts[1]) == 3 and parts[0] not in ['0', '-0']:
                x = x.replace('.', '') # 16.000 -> 16000 VNĐ
            else:
                pass # 12.5 -> 12.5
        elif num_commas == 1:
            parts = x.split(',')
            if len(parts[1]) == 3 and parts[0] not in ['0', '-0']:
                x = x.replace(',', '') # 16,000 -> 16000 VNĐ
            else:
                x = x.replace(',', '.') # 12,5 -> 12.5
    try:
        return float(x)
    except:
        return 0.0

def clean_number(x):
    # Hàm dự phòng giữ nguyên tương thích ngược
    return clean_val(x)

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
                # Google Sheets CSV exports are UTF-8 encoded
                content = response.content.decode('utf-8')
                return pd.read_csv(io.StringIO(content), skiprows=1, dtype=str)
            except Exception as e:
                if i == max_retries - 1:
                    raise e
                time.sleep(2)
                
    df = read_csv_with_retry(url_meat_fish)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Rename columns to standard ones if needed
    df.rename(columns={
        'ST': 'ID ST',
        'SL chênh lệch CXD': 'SL chênh lệch CXD'
    }, inplace=True)
    
    # Clean numeric columns
    qty_cols = ['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Qty_N', 'Qty_O', 'Qty_P', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']
    
    # Clean money columns
    for col in ['Tổng GT', 'Tổng hao hụt', 'Tổng ST', 'Tổng kho Thịt Cá', 'Tổng chưa xác định']:
        matched_cols = [c for c in df.columns if col in c]
        for c in matched_cols:
            df[c] = df[c].apply(clean_val)
            
    df['Số lượng chuyển'] = df['Số lượng chuyển'].apply(clean_qty)
    df['Số lượng nhận'] = df['Số lượng nhận'].apply(clean_qty)
    df['Chênh lệch'] = df['Chênh lệch'].apply(clean_qty)
    
    if 'Chi nhánh nhận' in df.columns:
        df['Chi nhánh nhận'] = df['Chi nhánh nhận'].astype(str).str.replace(',', '.', regex=False)
            
    # Lọc lý do chênh lệch
    df['LyDo_HaoHut'] = df['Hao hụt'].astype(str).str.strip().str.lower()
    df['LyDo_SieuThi'] = df['Siêu thị'].astype(str).str.strip().str.lower()
    
    col_kho = [c for c in df.columns if 'KHO THỊT CÁ' in c or 'kho thịt cá' in c or 'Kho thịt cá' in c or 'KHO TH' in c][0]
    df['LyDo_Kho'] = df[col_kho].astype(str).str.strip().str.lower()
    df['LyDo_Loi'] = df['Lỗi'].astype(str).str.strip().str.lower() if 'Lỗi' in df.columns else ''
    
    col_hao_hut_qty = 'Hạo hụt tự nhiê' if 'Hạo hụt tự nhiê' in df.columns else ('Hạo hụt tự nhiên' if 'Hạo hụt tự nhiên' in df.columns else None)
    df['Qty_N'] = df[col_hao_hut_qty].apply(clean_qty) if col_hao_hut_qty else df['Tổng hao hụt']
    df['Qty_O'] = df['SL trả tồn về ST'].apply(clean_qty) if 'SL trả tồn về ST' in df.columns else df['Tổng ST']
    df['Qty_P'] = df['SL chênh lệch CXD'].apply(clean_qty) if 'SL chênh lệch CXD' in df.columns else 0.0
    
    # Kết hợp các cột tổng chênh lệch
    col_total_kho = [c for c in df.columns if 'Tổng kho Thịt Cá' in c or 'Tổng kho thịt cá' in c][0]
    col_total_cxd = [c for c in df.columns if 'Tổng chưa xác định' in c or 'chưa xác định' in c][0]
    
    df['Hao hụt'] = np.where(df['LyDo_HaoHut'].str.contains('hao hụt'), df['Qty_N'], 0)
    df['BS_ST'] = np.where(df['LyDo_SieuThi'].str.contains('siêu thị'), df['Qty_O'], 0)
    df['ST_NhapThieu'] = np.where(df['LyDo_SieuThi'].str.contains('siêu thị') & df['LyDo_Loi'].str.contains('thiếu'), df['Qty_O'], 0)
    df['ST_SaiQT'] = np.where(df['LyDo_SieuThi'].str.contains('siêu thị') & ~df['LyDo_Loi'].str.contains('thiếu'), df['Qty_O'], 0)
    df['Kho_Rau'] = np.where(df['LyDo_Kho'].str.contains('kho thịt cá'), df['Qty_P'], 0)
    df['CXD'] = np.where(df['LyDo_Kho'].str.contains('chưa xác định'), df['Qty_P'], 0)
    
    df['Tổng kho rau'] = df[col_total_kho].apply(clean_val)
    df['Tổng chưa xác định'] = df[col_total_cxd].apply(clean_val)
    
    # Parse dates strictly (Google Sheets sends dates in MM/DD/YYYY format)
    date_col = 'Ngày' if 'Ngày' in df.columns else ('Ngành' if 'Ngành' in df.columns else 'Ngày')
    
    df['Ngày_parsed'] = pd.to_datetime(df[date_col], format='%m/%d/%Y', errors='coerce')
    df['Ngày_str'] = df['Ngày_parsed'].dt.strftime('%d/%m/%Y')
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
# HÀM TÍNH TOÁN NĂNG SUẤT DAILY MỚI
def calculate_daily_metrics(data, group_by_col='CLV2'):
    if data.empty:
        return pd.DataFrame(columns=[
            group_by_col, 'SL chuyển', 'SL chênh lệch', 'GT chênh lệch', 
            'SL ST chênh lệch', 'SL line chênh lệch', 'SL line hao hụt', 
            'SL line đã xử lý', 'Tỷ lệ line đã xử lý', 'Số lượng hao hụt', 
            'GT hao hụt', 'Tỷ lệ hao hụt', 'SL bs ST', 'GT bs ST', 
            'SL bs kho rau', 'GT bs kho rau', 'Đang xử lý', 'GT Đang xử lý', 
            'Chưa xử lý', 'GT Chưa xử lý', 'Không xử lý (WRITE OFF)', 'Giá trị WRITE OFF',
            'Lỗi ST (Nhập thiếu)', 'Lỗi ST (Sai QT)', 'GT Lỗi ST (Nhập thiếu)', 'GT Lỗi ST (Sai QT)'
        ])
    
    df = data.copy()
    df['SL_chuyen_num'] = to_numeric(df['Số lượng chuyển'])
    df['CL_num'] = to_numeric(df['Chênh lệch'])
    df['GT_num'] = to_numeric(df['Tổng GT'])
    df['HH_qty'] = to_numeric(df['Hao hụt'])
    df['HH_val'] = to_numeric(df['Tổng hao hụt'])
    df['ST_qty'] = to_numeric(df['BS_ST'])
    df['ST_val'] = to_numeric(df['Tổng ST'])
    df['Kho_qty'] = to_numeric(df['Kho_Rau'])
    df['Kho_val'] = to_numeric(df['Tổng kho rau'])
    df['CXD_qty'] = to_numeric(df['CXD'])
    df['CXD_val'] = to_numeric(df['Tổng chưa xác định'])
    
    df['ST_NhapThieu_qty'] = to_numeric(df['ST_NhapThieu']) if 'ST_NhapThieu' in df.columns else 0.0
    df['ST_SaiQT_qty'] = to_numeric(df['ST_SaiQT']) if 'ST_SaiQT' in df.columns else 0.0
    
    price_col = 'Giá nhập \n( -VAT)' if 'Giá nhập \n( -VAT)' in df.columns else None
    if price_col:
        df['ST_NhapThieu_val'] = df['ST_NhapThieu_qty'] * to_numeric(df[price_col])
        df['ST_SaiQT_val'] = df['ST_SaiQT_qty'] * to_numeric(df[price_col])
    else:
        df['ST_NhapThieu_val'] = 0.0
        df['ST_SaiQT_val'] = 0.0
        
    df['Xuly_clean'] = df['Xử lý'].fillna('').astype(str).str.strip().str.lower()
    
    groups = df.groupby(group_by_col, dropna=False)
    rows = []
    
    for g_name, g_df in groups:
        cl_df = g_df[g_df['CL_num'].abs() > 0]
        sl_chuyen = g_df['SL_chuyen_num'].sum()
        sl_cl = g_df['CL_num'].sum()
        gt_cl = g_df['GT_num'].sum()
        
        sl_st_cl = cl_df['ID ST'].nunique()
        sl_line_cl = len(cl_df)
        sl_line_hh = len(g_df[g_df['HH_qty'].abs() > 0])
        
        done_df = g_df[g_df['Xuly_clean'].str.contains('hoàn thành')]
        sl_line_done = len(done_df)
        tyle_line_done = f"{(sl_line_done / sl_line_cl * 100):.2f}%" if sl_line_cl > 0 else "0.00%"
        
        sl_hh = g_df['HH_qty'].sum()
        gt_hh = g_df['HH_val'].sum()
        tyle_hh = f"{(sl_hh / sl_chuyen * 100):.2f}%" if sl_chuyen > 0 else "0.00%"
        
        sl_bs_st = g_df['ST_qty'].sum()
        gt_bs_st = g_df['ST_val'].sum()
        sl_bs_kho = g_df['Kho_qty'].sum()
        gt_bs_kho = g_df['Kho_val'].sum()
        
        sl_st_nhap = g_df['ST_NhapThieu_qty'].sum()
        sl_st_sai = g_df['ST_SaiQT_qty'].sum()
        gt_st_nhap = g_df['ST_NhapThieu_val'].sum()
        gt_st_sai = g_df['ST_SaiQT_val'].sum()
        
        # Đang xử lý
        dang_xl_df = g_df[g_df['Xuly_clean'].str.contains('đang chuyển') | g_df['Xuly_clean'].str.contains('đang xử lý')]
        sl_dang_xl = dang_xl_df['CL_num'].sum()
        gt_dang_xl = dang_xl_df['GT_num'].sum()
        
        # Không xử lý (Write Off)
        write_off_df = g_df[g_df['Xuly_clean'].str.contains('không xử lý') | g_df['Xuly_clean'].str.contains('write off')]
        sl_write_off = write_off_df['CL_num'].sum()
        gt_write_off = write_off_df['GT_num'].sum()
        
        # Chưa xử lý
        chua_xl_df = g_df[~g_df['Xuly_clean'].str.contains('hoàn thành') & 
                          ~g_df['Xuly_clean'].str.contains('đang chuyển') & 
                          ~g_df['Xuly_clean'].str.contains('đang xử lý') & 
                          ~g_df['Xuly_clean'].str.contains('không xử lý') & 
                          ~g_df['Xuly_clean'].str.contains('write off')]
        sl_chua_xl = chua_xl_df['CL_num'].sum()
        gt_chua_xl = chua_xl_df['GT_num'].sum()
        
        row = {
            group_by_col: g_name,
            'SL chuyển': sl_chuyen,
            'SL chênh lệch': sl_cl,
            'GT chênh lệch': gt_cl,
            'SL ST chênh lệch': sl_st_cl,
            'SL line chênh lệch': sl_line_cl,
            'SL line hao hụt': sl_line_hh,
            'SL line đã xử lý': sl_line_done,
            'Tỷ lệ line đã xử lý': tyle_line_done,
            'Số lượng hao hụt': sl_hh,
            'GT hao hụt': gt_hh,
            'Tỷ lệ hao hụt': tyle_hh,
            'SL bs ST': sl_bs_st,
            'GT bs ST': gt_bs_st,
            'SL bs kho rau': sl_bs_kho,
            'GT bs kho rau': gt_bs_kho,
            'Đang xử lý': sl_dang_xl,
            'GT Đang xử lý': gt_dang_xl,
            'Chưa xử lý': sl_chua_xl,
            'GT Chưa xử lý': gt_chua_xl,
            'Không xử lý (WRITE OFF)': sl_write_off,
            'Giá trị WRITE OFF': gt_write_off,
            'Lỗi ST (Nhập thiếu)': sl_st_nhap,
            'Lỗi ST (Sai QT)': sl_st_sai,
            'GT Lỗi ST (Nhập thiếu)': gt_st_nhap,
            'GT Lỗi ST (Sai QT)': gt_st_sai
        }
        rows.append(row)
        
    return pd.DataFrame(rows)

# HÀM HIỂN THỊ BẢNG DAILY
def display_daily_table(df, cols, title_prefix, group_by_col='CLV2'):
    if df.empty:
        st.info("Không có dữ liệu.")
        return
    df_to_show = df.copy()
    for col in cols:
        if col not in df_to_show.columns:
            df_to_show[col] = 0.0
    df_to_show = df_to_show[cols]
    format_custom_table_with_total(df_to_show, group_by_col, title_prefix)

# HÀM TÍNH TỔNG QUAN DAILY DẠNG TEXT
def compute_daily_summary(df, date_str):
    if df.empty:
        return None
    
    cl_qty = to_numeric(df['Chênh lệch'])
    gt_val = to_numeric(df['Tổng GT'])
    
    total_items = cl_qty.abs().sum()
    total_value = gt_val.abs().sum()
    
    xuly_clean = df['Xử lý'].fillna('').astype(str).str.strip().str.lower()
    df_done = df[xuly_clean.str.contains('hoàn thành')]
    
    ret_qty = to_numeric(df_done['Kho_Rau']).abs().sum()
    bs_qty = to_numeric(df_done['BS_ST']).abs().sum()
    lost_qty = to_numeric(df_done['Hao hụt']).abs().sum()
    
    processed_qty = ret_qty + bs_qty + lost_qty
    remaining_qty = total_items - processed_qty
    if remaining_qty < 0:
        remaining_qty = 0.0
        
    return {
        'date': date_str,
        'total_items': int(total_items),
        'total_value': total_value,
        'processed': int(processed_qty),
        'return': int(ret_qty),
        'bs': int(bs_qty),
        'lost': int(lost_qty),
        'remaining': int(remaining_qty)
    }

# Insight Generators y hệt bên Kho Rau
def generate_insights(df_raw, table_type, df_grouped=None, df_metrics=None, date_str=None):
    if df_raw.empty and (df_grouped is None or df_grouped.empty):
        return "Không có dữ liệu trong kỳ báo cáo này."
    
    def get_hh_insight():
        if df_metrics is not None and not df_metrics.empty and 'Số lượng chuyển' in df_metrics.columns and 'Số lượng hao hụt' in df_metrics.columns:
            tong_chuyen = df_metrics['Số lượng chuyển'].sum()
            tong_hh = df_metrics['Số lượng hao hụt'].sum()
            if tong_chuyen > 0:
                return f"\n- Tỷ lệ hao hụt ghi nhận: {round((tong_hh / tong_chuyen) * 100, 2)}%."
        return ""
    
    try:
        if table_type == "Bảng 1":
            df_raw_tmp = df_raw.copy()
            df_raw_tmp['Chênh_lệch_num'] = to_numeric(df_raw_tmp['Chênh lệch'])
            df_raw_tmp['Kho_Rau_num'] = to_numeric(df_raw_tmp['Kho_Rau'])
            df_raw_tmp['BS_ST_num'] = to_numeric(df_raw_tmp['BS_ST'])
            
            total_lines = len(df_raw_tmp)
            total_chenh_lech = df_raw_tmp['Chênh_lệch_num'].sum()
            
            def fmt(val):
                try: return f"{int(val):,}".replace(',', '.') if float(val).is_integer() else f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return str(val)
                
            msg = f"Trong kỳ có tổng cộng {total_lines} dòng phát sinh chênh lệch (Tổng chênh lệch: {fmt(total_chenh_lech)}).\n"
            msg += "\n- Phân tích 3 nhóm ngành hàng (CLV4) phát sinh chênh lệch cao nhất:\n"
            
            clv4_lines = df_raw_tmp['CLV4'].value_counts()
            top3_clv4 = clv4_lines.head(3)
            
            for clv4, lines in top3_clv4.items():
                sub_df = df_raw_tmp[df_raw_tmp['CLV4'] == clv4]
                sub_cl = sub_df['Chênh_lệch_num'].sum()
                sub_kr = sub_df['Kho_Rau_num'].sum()
                msg += f"  + [{clv4}]: {lines} dòng (Tổng chênh lệch: {fmt(sub_cl)} | Trả về Kho thịt cá: {fmt(sub_kr)})\n"
                
            msg += "\n- Phân bổ trả về Siêu Thị (ST):\n"
            st_by_clv4 = df_raw_tmp.groupby('CLV4')['BS_ST_num'].sum().sort_values(ascending=False)
            st_by_clv4 = st_by_clv4[st_by_clv4 > 0]
            
            if not st_by_clv4.empty:
                top_st_clv4 = st_by_clv4.index[0]
                top_st_val = st_by_clv4.iloc[0]
                total_st = st_by_clv4.sum()
                if top_st_val > (total_st * 0.3) and len(st_by_clv4) > 1:
                    msg += f"  Số lượng trả về ST tập trung nhiều nhất ở nhóm [{top_st_clv4}] ({fmt(top_st_val)}).\n"
                elif len(st_by_clv4) > 1:
                    msg += f"  Số lượng trả về ST nằm rải rác lẻ tẻ (cao nhất là [{top_st_clv4}] với {fmt(top_st_val)}).\n"
                else:
                    msg += f"  Số lượng trả về ST thuộc về nhóm [{top_st_clv4}] ({fmt(top_st_val)}).\n"
            else:
                msg += "  Không phát sinh số lượng chênh lệch trả về ST trong kỳ.\n"
                
            msg += "  -> Nguyên nhân: Do ST thao tác sai nên phải tạo lại thôi."
            msg += get_hh_insight()
            return msg
            
        elif table_type == "Bảng 1.1":
            if df_grouped is not None and not df_grouped.empty:
                top_nguon = df_grouped.iloc[0]['Nguồn xác nhận']
                top_sl = df_grouped.iloc[0]['Tổng (Kho thịt cá + ST)'] if 'Tổng (Kho thịt cá + ST)' in df_grouped.columns else df_grouped.iloc[0]['Tổng (Kho Rau + ST)']
                
                def fmt(val):
                    try: return f"{int(val):,}".replace(',', '.') if float(val).is_integer() else f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    except: return str(val)
                
                if top_nguon == 'Check camera':
                    return f"Nguồn thông tin được dùng để xác định chênh lệch trả về các điểm nhận nhiều nhất là [{top_nguon}] (Số lượng: {fmt(top_sl)}).\n- Việc dựa phần lớn vào Check camera cho thấy tình trạng ST báo thiếu/dư hàng nhưng không cung cấp đủ hình ảnh xác thực đang khá cao. Cần nhắc nhở ST tuân thủ quy định chụp hình."
                else:
                    return f"Nguồn thông tin được dùng để xác định chênh lệch trả về các điểm nhận nhiều nhất là dựa vào [{top_nguon}] (Số lượng: {fmt(top_sl)}).\n- Điều này phản ánh cơ sở dữ liệu chính yếu mà DC dùng để đối soát và phân bổ lượng hàng chênh lệch trong kỳ."
            
        elif table_type == "Bảng 2.1_New":
            if df_metrics is not None and not df_metrics.empty:
                t_cl = df_metrics['SL chênh lệch'].sum()
                if t_cl > 0:
                    l_st_nhap = df_metrics.get('Lỗi ST (Nhập thiếu)', pd.Series([0])).sum()
                    l_st_sai = df_metrics.get('Lỗi ST (Sai QT)', pd.Series([0])).sum()
                    l_st_tong = l_st_nhap + l_st_sai
                    pct_st = (l_st_tong / t_cl) * 100
                    pct_nhap = (l_st_nhap / t_cl) * 100
                    pct_sai = (l_st_sai / t_cl) * 100
                    
                    giao_thieu_5 = df_metrics.get('<= 5%', pd.Series([0])).sum()
                    giao_thieu_10 = df_metrics.get('5-10%', pd.Series([0])).sum()
                    giao_thieu_15 = df_metrics.get('10-15%', pd.Series([0])).sum()
                    giao_thieu_15_plus = df_metrics.get('> 15%', pd.Series([0])).sum()
                    
                    pct_5 = (giao_thieu_5 / t_cl) * 100
                    pct_10 = (giao_thieu_10 / t_cl) * 100
                    pct_15 = (giao_thieu_15 / t_cl) * 100
                    pct_15_plus = (giao_thieu_15_plus / t_cl) * 100
                    
                    d_str = "kỳ báo cáo"
                    if date_str and date_str != "Tất cả các ngày":
                        try:
                            d_str = "ngày " + date_str.split('/')[0] + "." + date_str.split('/')[1]
                        except:
                            d_str = "ngày " + str(date_str)
                            
                    def get_top_clv4(col):
                        if col in df_metrics.columns and df_metrics[col].max() > 0:
                            top_row = df_metrics.loc[df_metrics[col].idxmax()]
                            return f"[{top_row['CLV4']}] - {int(top_row[col])} item"
                        return ""
                        
                    top_5_clv4 = get_top_clv4('<= 5%')
                    top_10_clv4 = get_top_clv4('5-10%')
                    top_15_clv4 = get_top_clv4('10-15%')
                    top_15_plus_clv4 = get_top_clv4('> 15%')
                            
                    msg = f"Hàng KG có số lượng nhập nhưng phát sinh chênh lệch ghi nhận {d_str}\n"
                    msg += f"- Lỗi ST chiếm {pct_st:.1f}%: trong đó nhập sót {pct_nhap:.1f}% và sai QT chiếm {pct_sai:.1f}%\n"
                    if l_st_sai > 0:
                        msg += f"  + SL ST sai QT: {int(l_st_sai)}\n"
                    msg += f"- Giao thiếu:\n"
                    msg += f"  + Nhóm <= 5%: {pct_5:.1f}%\n"
                    if top_5_clv4: msg += f"    -> Nhóm lệch nhiều nhất: {top_5_clv4}\n"
                    msg += f"  + Nhóm 5 - 10%: {pct_10:.1f}%\n"
                    if top_10_clv4: msg += f"    -> Nhóm lệch nhiều nhất: {top_10_clv4}\n"
                    msg += f"  + Nhóm 10 - 15%: {pct_15:.1f}%\n"
                    if top_15_clv4: msg += f"    -> Nhóm lệch nhiều nhất: {top_15_clv4}\n"
                    msg += f"  + Nhóm > 15%: {pct_15_plus:.1f}%"
                    if top_15_plus_clv4: msg += f"\n    -> Nhóm lệch nhiều nhất: {top_15_plus_clv4}"
                    return msg
            return "Chưa có đủ dữ liệu để đánh giá."
            
        elif table_type == "Bảng 2.1":
            clv4_counts = df_raw['CLV4'].value_counts()
            top3_clv4_str = ", ".join([f"[{k}] ({v} dòng)" for k, v in clv4_counts.head(3).items()]) if not clv4_counts.empty else 'Không xác định'
            
            sku_counts = df_raw['SKU_Full'].value_counts()
            top_sku = sku_counts.index[0] if not sku_counts.empty else 'Không xác định'
            top_sku_count = sku_counts.iloc[0] if not sku_counts.empty else 0
            
            return f"- Top 3 ngành hàng (CLV4) chiếm đa số chênh lệch: {top3_clv4_str}.\n- Đáng chú ý, mã hàng bị ảnh hưởng nhiều nhất là [{top_sku}] với {top_sku_count} dòng phát sinh." + get_hh_insight()
            
        elif table_type == "Bảng 3":
            clv4_counts = df_raw['CLV4'].value_counts()
            top3_clv4_str = ", ".join([f"[{k}] ({v} dòng)" for k, v in clv4_counts.head(3).items()]) if not clv4_counts.empty else 'Không xác định'
            
            sku_counts = df_raw['SKU_Full'].value_counts()
            top_sku = sku_counts.index[0] if not sku_counts.empty else 'Không xác định'
            top_sku_count = sku_counts.iloc[0] if not sku_counts.empty else 0
            
            base_msg = f"- Top 3 ngành hàng (CLV4) chiếm đa số chênh lệch: {top3_clv4_str}.\n- Đáng chú ý, mã hàng bị ảnh hưởng nhiều nhất là [{top_sku}] với {top_sku_count} dòng phát sinh."
            
            df_kr = df_raw.copy()
            df_kr['Kho_Rau_num'] = to_numeric(df_kr['Kho_Rau'])
            kr_by_clv4 = df_kr.groupby('CLV4')['Kho_Rau_num'].sum().sort_values(ascending=False)
            kr_by_clv4 = kr_by_clv4[kr_by_clv4 > 0]
            
            def fmt(val):
                try: return f"{int(val):,}".replace(',', '.') if float(val).is_integer() else f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return str(val)
                
            if not kr_by_clv4.empty:
                top3 = kr_by_clv4.head(3)
                top3_msg = "\n- Top nhóm ngành hàng (CLV4) đang có lượng chênh lệch trả về Kho thịt cá cao nhất:\n"
                for i, (clv4, val) in enumerate(top3.items(), 1):
                    top3_msg += f"  {i}. {clv4}: {fmt(val)}\n"
            else:
                top3_msg = "\n- Không ghi nhận hàng Pack nào có chênh lệch trả về Kho thịt cá trong kỳ."
                
            return base_msg + top3_msg.rstrip() + get_hh_insight()
            
        elif table_type == "Bảng 2.2":
            clv4_counts = df_raw['CLV4'].value_counts()
            top3_clv4_str = ", ".join([f"[{k}] ({v} dòng)" for k, v in clv4_counts.head(3).items()]) if not clv4_counts.empty else 'Không xác định'
            
            sku_counts = df_raw['SKU_Full'].value_counts()
            top_sku = sku_counts.index[0] if not sku_counts.empty else 'Không xác định'
            top_sku_count = sku_counts.iloc[0] if not sku_counts.empty else 0
            
            sum_kr = to_numeric(df_raw['Kho_Rau']).sum()
            sum_st = to_numeric(df_raw['BS_ST']).sum()
            
            def fmt(val):
                try: return f"{int(val):,}".replace(',', '.') if float(val).is_integer() else f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                except: return str(val)
                
            return (f"- Top 3 ngành hàng (CLV4) chiếm đa số chênh lệch: {top3_clv4_str}.\n"
                    f"- Đáng chú ý, mã hàng bị ảnh hưởng nhiều nhất là [{top_sku}] với {top_sku_count} dòng phát sinh.\n"
                    f"- Vấn đề chênh lệch này được phân bổ xử lý như sau:\n"
                    f"  + Trả về ST (Số lượng: {fmt(sum_st)}): Lý do là DC giao bù do ban đầu giao sai điểm.\n"
                    f"  + Trả Kho thịt cá (Số lượng: {fmt(sum_kr)}): Do có ST khác nhận dư số này và có ST nhận thiếu."
                    f"{get_hh_insight()}")
            
        elif table_type == "Bảng 4":
            if df_grouped is not None and not df_grouped.empty:
                top_sku = df_grouped.iloc[0]['Mã & Tên hàng']
                top_hh = df_grouped.iloc[0]['Tổng số lượng hao hụt']
                
                clv4_counts = df_raw['CLV4'].value_counts()
                top3_clv4_str = ", ".join([f"[{k}] ({v} dòng)" for k, v in clv4_counts.head(3).items()]) if not clv4_counts.empty else 'Không xác định'
                
                return f"- Top 3 ngành hàng (CLV4) phát sinh hao hụt nhiều nhất: {top3_clv4_str}.\n- Mã hàng có sản lượng hao hụt nghiêm trọng nhất là [{top_sku}] (Hao hụt: {top_hh} KG).\n- Khuyến nghị: Cần ưu tiên kiểm tra chất lượng thực tế và quy trình đóng gói đối với mã hàng này."

        elif table_type == "Bảng 6":
            if df_grouped is not None and not df_grouped.empty:
                top_dc = df_grouped.iloc[0]['DC xác nhận']
                top_loi = df_grouped.iloc[0]['Lỗi']
                top_sl = df_grouped.iloc[0]['Tổng số lượng']
                
                def fmt(val):
                    try: return f"{int(val):,}".replace(',', '.') if float(val).is_integer() else f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    except: return str(val)
                
                return f"- Dựa trên xác nhận của DC, lỗi [{top_loi}] được ghi nhận nhiều nhất từ [{top_dc}] với tổng số lượng trả về Kho thịt cá là {fmt(top_sl)}.\n- Khuyến nghị: DC cần kiểm tra lại quy trình xuất hàng và kiểm đếm để giảm thiểu tình trạng này."

    except Exception as e:
        return "Chưa đủ dữ liệu để tạo nhận xét tự động."
        
    return ""

def render_dc_feedback_progress_report(df, tab_id=""):
    st.write("---")
    
    if df.empty:
        st.info("Không có dữ liệu tiến độ DC phản hồi.")
        return
        
    # Tính toán daily summary (Toàn hệ thống)
    df['GT_chuyen_temp'] = to_numeric(df.get('Số lượng chuyển', pd.Series(0, index=df.index))) * to_numeric(df.get('Giá trị ĐV', pd.Series(0, index=df.index)))
    daily_summary = df.groupby('Ngày_str').agg(
        SL_chuyen=('Số lượng chuyển', lambda x: to_numeric(x).sum()),
        SL_chenh_lech=('Chênh lệch', lambda x: to_numeric(x).sum()),
        GT_chuyen=('GT_chuyen_temp', 'sum'),
        GT_chenh_lech=('Tổng GT', lambda x: to_numeric(x).sum())
    ).reset_index()

    df_dc = df[to_numeric(df.get('Kho_Rau', pd.Series(0, index=df.index))) > 0].copy()
    df_dc['SL_CXD'] = to_numeric(df_dc.get('Kho_Rau', pd.Series(0, index=df_dc.index)))
    df_dc['GT_CXD'] = to_numeric(df_dc.get('Tổng kho rau', pd.Series(0, index=df_dc.index)))
        
    if df_dc.empty:
        st.info("Không có dữ liệu tiến độ DC phản hồi trong kỳ báo cáo này.")
        return
        
    df_dc['DC_Xac_Nhan'] = df_dc['DC xác nhận'].fillna('Chưa xác nhận')
    df_dc['DC_Xac_Nhan'] = df_dc['DC_Xac_Nhan'].apply(lambda x: 'Chưa xác nhận' if str(x).strip() == '' else x)
    df_dc['Nhom_Loi'] = df_dc['Lỗi'].fillna('Không phân loại').replace('', 'Không phân loại')
    df_dc['Chi_Tiet_Loi'] = 'Không ghi chú'
    
    df_chua_xn = df_dc[df_dc['DC_Xac_Nhan'] == 'Chưa xác nhận']
    tong_chua_xn = df_chua_xn['SL_CXD'].sum()
    tong_gt_chua_xn = df_chua_xn['GT_CXD'].sum()
    
    top_loi_name = "Không có"
    if not df_chua_xn.empty:
        loi_sum = df_chua_xn.groupby('Nhom_Loi')['SL_CXD'].sum()
        if not loi_sum.empty and loi_sum.max() > 0:
            top_loi_name = f"{loi_sum.idxmax()} ({int(loi_sum.max())} item)"
            
    # Hiển thị Metrics
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="🔴 Tổng chờ DC xác nhận", value=f"{int(tong_chua_xn)} item", delta=f"{format_vn(tong_gt_chua_xn)} VNĐ", delta_color="off")
    with col2:
        st.metric(label="🔥 Top 1 Lỗi chờ phản hồi", value=top_loi_name)
    
    st.write("### 📌 Bảng chi tiết (Tiến độ DC)")
    tab_ngay, tab_loi = st.tabs(["📅 Góc nhìn 1: Theo Ngày", "⚠️ Góc nhìn 2: Theo Nhóm Lỗi"])
    
    # Góc nhìn 1
    with tab_ngay:
        st.markdown("**1. Bảng Số Lượng (Item)**")
        pivot_ngay = pd.pivot_table(df_dc, values='SL_CXD', index='Ngày_str', columns='DC_Xac_Nhan', aggfunc='sum', fill_value=0).reset_index()
        dc_cols = [c for c in pivot_ngay.columns if c != 'Ngày_str']
        pivot_ngay['SL KHO THỊT CÁ'] = pivot_ngay[dc_cols].sum(axis=1)
        
        final_ngay = pd.merge(daily_summary[['Ngày_str', 'SL_chuyen', 'SL_chenh_lech']], pivot_ngay, on='Ngày_str', how='right')
        sorted_dc_cols = [c for c in dc_cols if c != 'Chưa xác nhận']
        sorted_dc_cols.sort()
        if 'Chưa xác nhận' in dc_cols:
            sorted_dc_cols = ['Chưa xác nhận'] + sorted_dc_cols
            
        col_order = ['Ngày_str', 'SL_chuyen', 'SL_chenh_lech', 'SL KHO THỊT CÁ'] + sorted_dc_cols
        final_ngay = final_ngay[[c for c in col_order if c in final_ngay.columns]]
        
        final_ngay['% Chưa xác nhận'] = final_ngay.apply(
            lambda r: f"{(r.get('Chưa xác nhận', 0) / r['SL KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('SL KHO THỊT CÁ', 0) > 0 else "0,00%", axis=1
        )
        final_ngay['% Tiến độ phản hồi'] = final_ngay.apply(
            lambda r: f"{((r['SL KHO THỊT CÁ'] - r.get('Chưa xác nhận', 0)) / r['SL KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('SL KHO THỊT CÁ', 0) > 0 else "100,00%", axis=1
        )
        
        final_ngay.rename(columns={
            'Ngày_str': 'Ngày chuyển hàng',
            'SL_chuyen': 'SL chuyển',
            'SL_chenh_lech': 'SL chênh lệch'
        }, inplace=True)
        format_custom_table_with_total(final_ngay, 'Ngày chuyển hàng', f"Tien_Do_DC_Theo_Ngay_SL_{tab_id}")
        
        st.markdown("**2. Bảng Giá Trị (VNĐ)**")
        pivot_ngay_gt = pd.pivot_table(df_dc, values='GT_CXD', index='Ngày_str', columns='DC_Xac_Nhan', aggfunc='sum', fill_value=0).reset_index()
        pivot_ngay_gt['GT KHO THỊT CÁ'] = pivot_ngay_gt[dc_cols].sum(axis=1)
        final_ngay_gt = pd.merge(daily_summary[['Ngày_str', 'GT_chuyen', 'GT_chenh_lech']], pivot_ngay_gt, on='Ngày_str', how='right')
        
        col_order_gt = ['Ngày_str', 'GT_chuyen', 'GT_chenh_lech', 'GT KHO THỊT CÁ'] + sorted_dc_cols
        final_ngay_gt = final_ngay_gt[[c for c in col_order_gt if c in final_ngay_gt.columns]]
        
        final_ngay_gt['% Chưa xác nhận'] = final_ngay_gt.apply(
            lambda r: f"{(r.get('Chưa xác nhận', 0) / r['GT KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('GT KHO THỊT CÁ', 0) > 0 else "0,00%", axis=1
        )
        final_ngay_gt['% Tiến độ phản hồi'] = final_ngay_gt.apply(
            lambda r: f"{((r['GT KHO THỊT CÁ'] - r.get('Chưa xác nhận', 0)) / r['GT KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('GT KHO THỊT CÁ', 0) > 0 else "100,00%", axis=1
        )
        
        final_ngay_gt.rename(columns={
            'Ngày_str': 'Ngày chuyển hàng',
            'GT_chuyen': 'GT chuyển (VNĐ)',
            'GT_chenh_lech': 'GT chênh lệch (VNĐ)'
        }, inplace=True)
        format_custom_table_with_total(final_ngay_gt, 'Ngày chuyển hàng', f"Tien_Do_DC_Theo_Ngay_GT_{tab_id}")
        
    # Góc nhìn 2
    with tab_loi:
        df_dc['Nhóm Lỗi & Chi tiết'] = df_dc['Nhom_Loi'] + " | " + df_dc['Chi_Tiet_Loi']
        
        st.markdown("**1. Bảng Số Lượng (Item)**")
        pivot_loi = pd.pivot_table(df_dc, values='SL_CXD', index='Nhóm Lỗi & Chi tiết', columns='DC_Xac_Nhan', aggfunc='sum', fill_value=0).reset_index()
        pivot_loi['SL KHO THỊT CÁ'] = pivot_loi[dc_cols].sum(axis=1)
        col_order_loi = ['Nhóm Lỗi & Chi tiết', 'SL KHO THỊT CÁ'] + sorted_dc_cols
        pivot_loi = pivot_loi[[c for c in col_order_loi if c in pivot_loi.columns]]
        
        pivot_loi['% Chưa xác nhận'] = pivot_loi.apply(
            lambda r: f"{(r.get('Chưa xác nhận', 0) / r['SL KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('SL KHO THỊT CÁ', 0) > 0 else "0,00%", axis=1
        )
        pivot_loi['% Tiến độ phản hồi'] = pivot_loi.apply(
            lambda r: f"{((r['SL KHO THỊT CÁ'] - r.get('Chưa xác nhận', 0)) / r['SL KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('SL KHO THỊT CÁ', 0) > 0 else "100,00%", axis=1
        )
        
        format_custom_table_with_total(pivot_loi, 'Nhóm Lỗi & Chi tiết', f"Tien_Do_DC_Theo_Loi_SL_{tab_id}")
        
        st.markdown("**2. Bảng Giá Trị (VNĐ)**")
        pivot_loi_gt = pd.pivot_table(df_dc, values='GT_CXD', index='Nhóm Lỗi & Chi tiết', columns='DC_Xac_Nhan', aggfunc='sum', fill_value=0).reset_index()
        pivot_loi_gt['GT KHO THỊT CÁ'] = pivot_loi_gt[dc_cols].sum(axis=1)
        pivot_loi_gt = pivot_loi_gt[[c for c in col_order_loi if c in pivot_loi_gt.columns]]
        pivot_loi_gt.rename(columns={'SL KHO THỊT CÁ': 'GT KHO THỊT CÁ'}, inplace=True)
        
        pivot_loi_gt['% Chưa xác nhận'] = pivot_loi_gt.apply(
            lambda r: f"{(r.get('Chưa xác nhận', 0) / r['GT KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('GT KHO THỊT CÁ', 0) > 0 else "0,00%", axis=1
        )
        pivot_loi_gt['% Tiến độ phản hồi'] = pivot_loi_gt.apply(
            lambda r: f"{((r['GT KHO THỊT CÁ'] - r.get('Chưa xác nhận', 0)) / r['GT KHO THỊT CÁ'] * 100):.2f}%".replace('.', ',') if r.get('GT KHO THỊT CÁ', 0) > 0 else "100,00%", axis=1
        )
        
        format_custom_table_with_total(pivot_loi_gt, 'Nhóm Lỗi & Chi tiết', f"Tien_Do_DC_Theo_Loi_GT_{tab_id}")

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

# ==========================================
# GIAO DIỆN CHIA TAB
# ==========================================
tab_main, tab_daily, tab_dc = st.tabs(["📊 Báo Cáo Tổng Quan", "📈 Báo Cáo Năng Suất Daily", "👨‍🔧 Tiến Độ DC Phản Hồi"])

# ==========================================
# TRANG 1: BÁO CÁO TỔNG QUAN
# ==========================================
with tab_main:
    # Nút Cập nhật dữ liệu mới nhất
    if st.button('🔄 Cập nhật dữ liệu mới nhất'):
        st.cache_data.clear()
        st.rerun()
        
    df_active = df_all.copy()

    # Process Dataframes
    pivot_ngay_sum = df_active.groupby('Ngày_str')[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tổng GT', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']].sum()
    pivot_ngay = pivot_ngay_sum.fillna(0).reset_index()
    
    pivot_ngay['Ngày_dt'] = pd.to_datetime(pivot_ngay['Ngày_str'], format='%d/%m/%Y', errors='coerce')
    pivot_ngay = pivot_ngay.sort_values(by='Ngày_dt').drop(columns=['Ngày_dt'])

    # Tính phần trăm đã phân bổ / chênh lệch cho từng ngày
    sum_dist = pivot_ngay['Hao hụt'] + pivot_ngay['BS_ST'] + pivot_ngay['Kho_Rau'] + pivot_ngay['CXD']
    pct_vals = np.where(pivot_ngay['Chênh lệch'].abs() > 0, (sum_dist / pivot_ngay['Chênh lệch'].abs()) * 100, 0.0)
    pivot_ngay['% Cột Tổng'] = [f"{v:.2f}%".replace('.', ',') for v in pct_vals]

    tong_row_ngay = pivot_ngay.sum(numeric_only=True).to_frame().T
    tong_row_ngay['Ngày_str'] = 'Tổng'
    
    # Tính phần trăm đã phân bổ / chênh lệch cho hàng Tổng
    total_sum_dist = tong_row_ngay['Hao hụt'].iloc[0] + tong_row_ngay['BS_ST'].iloc[0] + tong_row_ngay['Kho_Rau'].iloc[0] + tong_row_ngay['CXD'].iloc[0]
    total_cl = abs(tong_row_ngay['Chênh lệch'].iloc[0])
    total_pct = (total_sum_dist / total_cl) * 100 if total_cl > 0 else 0.0
    tong_row_ngay['% Cột Tổng'] = f"{total_pct:.2f}%".replace('.', ',')

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

    # Bảng SỐ LƯỢNG Chi tiết Từng Ngày - Siêu Thị
    pivot_qty_sum = df_active.groupby(['Ngày_str', 'ID ST', 'Chi nhánh nhận'], dropna=False)[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']].sum()
    pivot_qty_count = df_active[df_active['Chênh lệch'].abs() > 0].groupby(['Ngày_str', 'ID ST', 'Chi nhánh nhận'], dropna=False).size().rename('SL line chênh lệch')
    pivot_qty_nhap0 = df_active[(df_active['Số lượng nhận'] == 0) & (df_active['Chênh lệch'].abs() > 0)].groupby(['Ngày_str', 'ID ST', 'Chi nhánh nhận'], dropna=False).size().rename('SL line nhập=0')

    pivot_qty = pivot_qty_sum.join(pivot_qty_count).join(pivot_qty_nhap0).fillna(0).reset_index()
    pivot_qty.rename(columns={
        'BS_ST': 'SL đã tạo bs cho ST',
        'Kho_Rau': 'SL đã xác nhận được trả kho thịt cá',
        'Hao hụt': 'Số lượng hao hụt',
        'CXD': 'Số lượng chưa xác định'
    }, inplace=True)
    pivot_qty['Tỷ lệ (%)'] = np.where(pivot_qty['Số lượng chuyển'] > 0, (pivot_qty['Chênh lệch'] / pivot_qty['Số lượng chuyển']) * 100, 0)
    pivot_qty['Abs_ChenhLech'] = pivot_qty['Chênh lệch'].abs()
    pivot_qty = pivot_qty.sort_values(by='Abs_ChenhLech', ascending=False).drop(columns=['Abs_ChenhLech'])

    pivot_qty['SL line chênh lệch'] = pivot_qty['SL line chênh lệch'].astype(int)
    pivot_qty['SL line nhập=0'] = pivot_qty['SL line nhập=0'].astype(int)
    pivot_qty.insert(3, 'SL SKU NHẬP = 0/SL SKU CHÊNH LỆCH', pivot_qty['SL line nhập=0'].astype(str) + " / " + pivot_qty['SL line chênh lệch'].astype(str))
    pivot_qty = pivot_qty[['Ngày_str', 'ID ST', 'Chi nhánh nhận', 'SL SKU NHẬP = 0/SL SKU CHÊNH LỆCH', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tỷ lệ (%)', 'SL đã tạo bs cho ST', 'SL đã xác nhận được trả kho thịt cá', 'Số lượng hao hụt', 'Số lượng chưa xác định']]

    # Bảng GIÁ TRỊ Chi tiết Từng Ngày - Siêu Thị
    pivot_val_sum = df_active.groupby(['Ngày_str', 'ID ST', 'Chi nhánh nhận'], dropna=False)[['Tổng GT', 'Tổng ST', 'Tổng kho rau', 'Tổng chưa xác định']].sum().reset_index()
    pivot_val_sum.rename(columns={'Tổng GT': 'Giá trị chênh lệch (VNĐ)'}, inplace=True)

    # Thẻ thông tin (Metrics)
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
    
    if not pivot_ngay.empty:
        top_day = pivot_ngay.sort_values(by='Chênh lệch', ascending=False).iloc[0]
        st.info(f"🔹 **Ngày biến động nhất**: **{top_day['Ngày_str']}** ghi nhận mức chênh lệch cao nhất ({format_vn(top_day['Chênh lệch'])} item).")

    tab_ngay_qty, tab_ngay_val = st.tabs(["📊 Số lượng (Từng Ngày)", "💰 Giá trị (Từng Ngày)"])

    with tab_ngay_qty:
        display_df_with_download(pivot_ngay_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in pivot_ngay_renamed.columns if 'Chênh lệch' in c[1]]), "Tong_Hop_Theo_Ngay_So_Luong")

    with tab_ngay_val:
        display_df_with_download(pivot_ngay_val_renamed.style.format(format_vn), "Tong_Hop_Theo_Ngay_Gia_Tri")

    st.write("---")
    col4, col5 = st.columns(2)
    with col4:
        st.subheader("🔥 2. TOP 5 CATE CHÊNH LỆCH LỚN NHẤT")
        if not pivot_clv4.empty:
            top_clv4 = pivot_clv4.iloc[0]
            st.info(f"🔹 **Mã hàng (CLV4) cảnh báo đỏ**: **{top_clv4['CLV4']}** đang dẫn đầu với mức chênh lệch {format_vn(top_clv4['Chênh lệch'])}.")
        display_df_with_download(pivot_clv4.style.format(format_vn).map(color_red_for_chenhlech, subset=['Chênh lệch']), "Top_5_CLV4")
    with col5:
        st.subheader("📦 3. TỔNG HỢP THEO NGÀNH HÀNG (CLV2)")
        if not pivot_clv2.empty:
            top_clv2 = pivot_clv2.iloc[0]
            st.info(f"🔹 **Ngành hàng (CLV2) trọng điểm**: **{top_clv2['CLV2']}** chiếm số lượng chênh lệch cao nhất ({format_vn(top_clv2['Chênh lệch'])}).")
        display_df_with_download(pivot_clv2_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in pivot_clv2_renamed.columns if 'Chênh lệch' in c[1]]), "Tong_Hop_CLV2")

    st.write("---")
    # Sort dates chronologically
    sorted_dates_dt = sorted(pd.to_datetime(pivot_ngay['Ngày_str'], format='%d/%m/%Y', errors='coerce').dropna().unique())
    sorted_dates = [d.strftime('%d/%m/%Y') for d in sorted_dates_dt if d.strftime('%d/%m/%Y') != 'Tổng']
    dates = ["Tất cả các ngày"] + sorted_dates

    # 4. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO NHÓM HÀNG (CLV4)
    st.subheader("🛒 4. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO NHÓM HÀNG (CLV4)")
    item_qty_sum = df_active.groupby(['Ngày_str', 'CLV4'], dropna=False)[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']].sum()
    item_qty_count = df_active[df_active['Chênh lệch'].abs() > 0].groupby(['Ngày_str', 'CLV4'], dropna=False).size().rename('SL ST chênh lệch')
    item_qty_nhap0 = df_active[(df_active['Số lượng nhận'] == 0) & (df_active['Chênh lệch'].abs() > 0)].groupby(['Ngày_str', 'CLV4'], dropna=False).size().rename('SL ST nhập=0')

    pivot_qty_item = item_qty_sum.join(item_qty_count).join(item_qty_nhap0).fillna(0).reset_index()
    pivot_qty_item.rename(columns={
        'CLV4': 'Mã hàng (CLV4)',
        'BS_ST': 'SL đã tạo bs cho ST',
        'Kho_Rau': 'SL đã xác nhận được trả kho thịt cá',
        'Hao hụt': 'Số lượng hao hụt',
        'CXD': 'Số lượng chưa xác định'
    }, inplace=True)
    pivot_qty_item['Tỷ lệ (%)'] = np.where(pivot_qty_item['Số lượng chuyển'] > 0, (pivot_qty_item['Chênh lệch'] / pivot_qty_item['Số lượng chuyển']) * 100, 0)
    pivot_qty_item['Abs_ChenhLech'] = pivot_qty_item['Chênh lệch'].abs()
    pivot_qty_item = pivot_qty_item.sort_values(by='Abs_ChenhLech', ascending=False).drop(columns=['Abs_ChenhLech'])

    pivot_qty_item['SL ST chênh lệch'] = pivot_qty_item['SL ST chênh lệch'].astype(int)
    pivot_qty_item['SL ST nhập=0'] = pivot_qty_item['SL ST nhập=0'].astype(int)
    pivot_qty_item.insert(2, 'SL ST NHẬP = 0/SL ST CHÊNH LỆCH', pivot_qty_item['SL ST nhập=0'].astype(str) + " / " + pivot_qty_item['SL ST chênh lệch'].astype(str))
    pivot_qty_item = pivot_qty_item[['Ngày_str', 'Mã hàng (CLV4)', 'SL ST NHẬP = 0/SL ST CHÊNH LỆCH', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tỷ lệ (%)', 'SL đã tạo bs cho ST', 'SL đã xác nhận được trả kho thịt cá', 'Số lượng hao hụt', 'Số lượng chưa xác định']]

    pivot_val_item = df_active.groupby(['Ngày_str', 'CLV4'], dropna=False)[['Tổng GT', 'Tổng ST', 'Tổng kho rau', 'Tổng chưa xác định']].sum().reset_index()
    pivot_val_item.rename(columns={'Tổng GT': 'Giá trị chênh lệch (VNĐ)', 'CLV4': 'Mã hàng (CLV4)'}, inplace=True)

    selected_date_item = st.selectbox("🔍 Lọc theo Ngày (Mã hàng):", dates)
    if selected_date_item != "Tất cả các ngày":
        filtered_qty_item = pivot_qty_item[pivot_qty_item['Ngày_str'] == selected_date_item]
        filtered_val_item = pivot_val_item[pivot_val_item['Ngày_str'] == selected_date_item]
    else:
        filtered_qty_item = pivot_qty_item
        filtered_val_item = pivot_val_item

    tong_qty_item = pd.DataFrame() if filtered_qty_item.empty else filtered_qty_item.sum(numeric_only=True).to_frame().T
    if not tong_qty_item.empty: tong_qty_item['Ngày_str'] = 'Tổng'
    filtered_qty_item_renamed = create_multiindex_headers(filtered_qty_item, tong_qty_item)

    tong_val_item = pd.DataFrame() if filtered_val_item.empty else filtered_val_item.sum(numeric_only=True).to_frame().T
    if not tong_val_item.empty: tong_val_item['Ngày_str'] = 'Tổng'
    filtered_val_item_renamed = create_multiindex_headers(filtered_val_item, tong_val_item)

    tab3, tab4 = st.tabs(["📊 Chi Tiết SỐ LƯỢNG (Mã Hàng)", "💰 Chi Tiết GIÁ TRỊ (Mã Hàng)"])
    with tab3:
        display_df_with_download(filtered_qty_item_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in filtered_qty_item_renamed.columns if 'Chênh lệch' in c[1]]), "Chi_Tiet_SL_CLV4", height=600)
    with tab4:
        display_df_with_download(filtered_val_item_renamed.style.format(format_vn), "Chi_Tiet_GT_CLV4", height=600)

    # 5. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO MÃ HÀNG (SKU)
    st.write("---")
    st.subheader("🏷️ 5. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO MÃ HÀNG (SKU)")
    sku_qty_sum = df_active.groupby(['Ngày_str', 'SKU_Full'], dropna=False)[['Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Hao hụt', 'BS_ST', 'Kho_Rau', 'CXD']].sum()
    sku_qty_count = df_active[df_active['Chênh lệch'].abs() > 0].groupby(['Ngày_str', 'SKU_Full'], dropna=False).size().rename('SL ST chênh lệch')
    sku_qty_nhap0 = df_active[(df_active['Số lượng nhận'] == 0) & (df_active['Chênh lệch'].abs() > 0)].groupby(['Ngày_str', 'SKU_Full'], dropna=False).size().rename('SL ST nhập=0')

    pivot_qty_sku = sku_qty_sum.join(sku_qty_count).join(sku_qty_nhap0).fillna(0).reset_index()
    pivot_qty_sku.rename(columns={
        'SKU_Full': 'Mã hàng (SKU)',
        'BS_ST': 'SL đã tạo bs cho ST',
        'Kho_Rau': 'SL đã xác nhận được trả kho thịt cá',
        'Hao hụt': 'Số lượng hao hụt',
        'CXD': 'Số lượng chưa xác định'
    }, inplace=True)
    pivot_qty_sku['Tỷ lệ (%)'] = np.where(pivot_qty_sku['Số lượng chuyển'] > 0, (pivot_qty_sku['Chênh lệch'] / pivot_qty_sku['Số lượng chuyển']) * 100, 0)
    pivot_qty_sku['Abs_ChenhLech'] = pivot_qty_sku['Chênh lệch'].abs()
    pivot_qty_sku = pivot_qty_sku.sort_values(by='Abs_ChenhLech', ascending=False).drop(columns=['Abs_ChenhLech'])

    pivot_qty_sku['SL ST chênh lệch'] = pivot_qty_sku['SL ST chênh lệch'].astype(int)
    pivot_qty_sku['SL ST nhập=0'] = pivot_qty_sku['SL ST nhập=0'].astype(int)
    pivot_qty_sku.insert(2, 'SL ST NHẬP = 0/SL ST CHÊNH LỆCH', pivot_qty_sku['SL ST nhập=0'].astype(str) + " / " + pivot_qty_sku['SL ST chênh lệch'].astype(str))
    pivot_qty_sku = pivot_qty_sku[['Ngày_str', 'Mã hàng (SKU)', 'SL ST NHẬP = 0/SL ST CHÊNH LỆCH', 'Số lượng chuyển', 'Số lượng nhận', 'Chênh lệch', 'Tỷ lệ (%)', 'SL đã tạo bs cho ST', 'SL đã xác nhận được trả kho thịt cá', 'Số lượng hao hụt', 'Số lượng chưa xác định']]

    pivot_val_sku = df_active.groupby(['Ngày_str', 'SKU_Full'], dropna=False)[['Tổng GT', 'Tổng ST', 'Tổng kho rau', 'Tổng chưa xác định']].sum().reset_index()
    pivot_val_sku.rename(columns={'Tổng GT': 'Giá trị chênh lệch (VNĐ)', 'SKU_Full': 'Mã hàng (SKU)'}, inplace=True)

    selected_date_sku = st.selectbox("🔍 Lọc theo Ngày (SKU):", dates)
    if selected_date_sku != "Tất cả các ngày":
        filtered_qty_sku = pivot_qty_sku[pivot_qty_sku['Ngày_str'] == selected_date_sku]
        filtered_val_sku = pivot_val_sku[pivot_val_sku['Ngày_str'] == selected_date_sku]
    else:
        filtered_qty_sku = pivot_qty_sku
        filtered_val_sku = pivot_val_sku

    tong_qty_sku = pd.DataFrame() if filtered_qty_sku.empty else filtered_qty_sku.sum(numeric_only=True).to_frame().T
    if not tong_qty_sku.empty: tong_qty_sku['Ngày_str'] = 'Tổng'
    filtered_qty_sku_renamed = create_multiindex_headers(filtered_qty_sku, tong_qty_sku)

    tong_val_sku = pd.DataFrame() if filtered_val_sku.empty else filtered_val_sku.sum(numeric_only=True).to_frame().T
    if not tong_val_sku.empty: tong_val_sku['Ngày_str'] = 'Tổng'
    filtered_val_sku_renamed = create_multiindex_headers(filtered_val_sku, tong_val_sku)

    tab5, tab6 = st.tabs(["📊 Chi Tiết SỐ LƯỢNG (SKU)", "💰 Chi Tiết GIÁ TRỊ (SKU)"])
    with tab5:
        display_df_with_download(filtered_qty_sku_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in filtered_qty_sku_renamed.columns if 'Chênh lệch' in c[1]]), "Chi_Tiet_SL_SKU", height=600)
    with tab6:
        display_df_with_download(filtered_val_sku_renamed.style.format(format_vn), "Chi_Tiet_GT_SKU", height=600)

    # 6. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO SIÊU THỊ
    st.write("---")
    st.subheader("🏬 6. CHI TIẾT SỐ LƯỢNG & GIÁ TRỊ THEO SIÊU THỊ")
    selected_date = st.selectbox("🔍 Lọc theo Ngày:", dates)
    if selected_date != "Tất cả các ngày":
        filtered_qty = pivot_qty[pivot_qty['Ngày_str'] == selected_date]
        filtered_val = pivot_val_sum[pivot_val_sum['Ngày_str'] == selected_date]
    else:
        filtered_qty = pivot_qty
        filtered_val = pivot_val_sum

    tong_qty = pd.DataFrame() if filtered_qty.empty else filtered_qty.sum(numeric_only=True).to_frame().T
    if not tong_qty.empty: tong_qty['Ngày_str'] = 'Tổng'
    filtered_qty_renamed = create_multiindex_headers(filtered_qty, tong_qty)

    tong_val = pd.DataFrame() if filtered_val.empty else filtered_val.sum(numeric_only=True).to_frame().T
    if not tong_val.empty: tong_val['Ngày_str'] = 'Tổng'
    filtered_val_renamed = create_multiindex_headers(filtered_val, tong_val)

    tab1, tab2 = st.tabs(["📊 Chi Tiết SỐ LƯỢNG", "💰 Chi Tiết GIÁ TRỊ (VNĐ)"])
    with tab1:
        display_df_with_download(filtered_qty_renamed.style.format(format_vn).map(color_red_for_chenhlech, subset=[c for c in filtered_qty_renamed.columns if 'Chênh lệch' in c[1]]), "Chi_Tiet_SL_Sieu_Thi", height=600)
    with tab2:
        display_df_with_download(filtered_val_renamed.style.format(format_vn), "Chi_Tiet_GT_Sieu_Thi", height=600)

# ==========================================
# TRANG 2: BÁO CÁO DAILY MỚI
# ==========================================
with tab_daily:
    st.header("Báo Cáo Năng Suất Chi Tiết Mỗi Ngày")
    
    # Sort dates chronologically
    unique_daily_dates_dt = sorted(pd.to_datetime(df_all['Ngày_str'], format='%d/%m/%Y', errors='coerce').dropna().unique())
    unique_daily_dates = [d.strftime('%d/%m/%Y') for d in unique_daily_dates_dt]
    
    if unique_daily_dates:
        options = ["Tất cả các ngày"] + unique_daily_dates
        selected_daily_date = st.selectbox("📅 Chọn ngày báo cáo (Daily):", options, index=len(options)-1, key="daily_selectbox")
        df_filtered = df_all if selected_daily_date == "Tất cả các ngày" else df_all[df_all['Ngày_str'] == selected_daily_date].copy()
    else:
        df_filtered = pd.DataFrame()
    
    df_b1 = calculate_daily_metrics(df_filtered)
    cols = ['CLV2', 'SL chuyển', 'SL chênh lệch', 'SL ST chênh lệch', 'SL line chênh lệch', 'SL line hao hụt', 'SL line đã xử lý', 'Tỷ lệ line đã xử lý', 'Số lượng hao hụt', 'Tỷ lệ hao hụt', 'SL bs ST', 'SL bs kho rau', 'Đang xử lý', 'Chưa xử lý', 'Không xử lý (WRITE OFF)']
    cols_gt = ['CLV2', 'SL chuyển', 'GT chênh lệch', 'SL ST chênh lệch', 'SL line chênh lệch', 'SL line hao hụt', 'SL line đã xử lý', 'Tỷ lệ line đã xử lý', 'GT hao hụt', 'Tỷ lệ hao hụt', 'GT bs ST', 'GT bs kho rau', 'GT Đang xử lý', 'GT Chưa xử lý', 'Giá trị WRITE OFF']
    
    st.subheader("Bảng 1: Đánh giá nhanh tình hình xử lý")
    t1_sl, t1_gt = st.tabs(["📊 Số Lượng", "💰 Giá Trị"])
    with t1_sl:
        display_daily_table(df_b1, cols, "Bang_1")
    with t1_gt:
        display_daily_table(df_b1, cols_gt, "Bang_1_GT")
        
    nx_b1 = generate_insights(df_filtered, "Bảng 1", df_b1)
    st.text_area("Nhận xét Bảng 1:", value=nx_b1, key="nx_b1", height=100)
    
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
        
    # Bảng 1.1 Chi tiết đã xử lý qua Note
    st.write("---")
    st.subheader("Bảng 1.1: Chi tiết đã xử lý")
    df_note_data = df_filtered.copy()
    df_note_data['Kho_Rau_num'] = to_numeric(df_note_data['Kho_Rau'])
    df_note_data['BS_ST_num'] = to_numeric(df_note_data['BS_ST'])
    df_note_data = df_note_data[(df_note_data['Kho_Rau_num'] > 0) | (df_note_data['BS_ST_num'] > 0)].copy()
    
    if not df_note_data.empty:
        def map_note_to_category(note):
            note_str = str(note).lower().strip()
            if 'tele' in note_str or 'kdb' in note_str or 'hình' in note_str: return 'Hình ảnh ST'
            if 'st nhận' in note_str or 'giao sai' in note_str: return 'DC giao sai ST'
            if 'pick sai' in note_str or 'lấy sai' in note_str: return 'DC pick sai'
            return 'Check camera'
            
        df_note_data['Nguồn xác nhận'] = df_note_data['NOTE'].apply(map_note_to_category)
        df_note = df_note_data.groupby('Nguồn xác nhận').agg(
            SL_bs_kho_rau=('Kho_Rau_num', 'sum'),
            SL_bs_st=('BS_ST_num', 'sum'),
            So_lan=('Mã hàng', 'count')
        ).reset_index()
        df_note['Tổng (Kho thịt cá + ST)'] = df_note['SL_bs_kho_rau'] + df_note['SL_bs_st']
        df_note = df_note.sort_values(by='Tổng (Kho thịt cá + ST)', ascending=False)
        df_note.rename(columns={'SL_bs_kho_rau': 'SL bs kho thịt cá', 'SL_bs_st': 'SL bs ST', 'So_lan': 'Số line'}, inplace=True)
        format_custom_table_with_total(df_note, 'Nguồn xác nhận', "Chi_Tiet_Ly_Do_Xu_Ly")
    else:
        st.info("Không có dữ liệu lý do xử lý.")

    # Bảng 2.1 Đánh giá hàng KG
    st.write("---")
    st.subheader("Bảng 2: Đánh giá tình hình xử lý hàng theo ĐVT: KG")
    df_kg = df_filtered[df_filtered['Loại hàng'].astype(str).str.upper() == 'KG'].copy()
    
    st.markdown("**2.1 Đánh giá mức độ nghiêm trọng chênh lệch (Hàng KG Nhận > 0 - Theo CLV4)**")
    df_kg_nhan = df_kg[to_numeric(df_kg['Số lượng nhận']) > 0].copy()
    df_b21_new = calculate_daily_metrics(df_kg_nhan, group_by_col='CLV4')
    df_b21_new['Hao hụt (<=10%)'] = df_b21_new['Số lượng hao hụt']
    df_b21_new['Trả KR (Lỗi giao thiếu)'] = df_b21_new['SL bs kho rau']
    df_b21_new['Tổng Trả Kho Rau'] = df_b21_new['Hao hụt (<=10%)'] + df_b21_new['Trả KR (Lỗi giao thiếu)']
    df_b21_new['GT Hao hụt (<=10%)'] = df_b21_new['GT hao hụt']
    df_b21_new['GT Trả KR (Lỗi giao thiếu)'] = df_b21_new['GT bs kho rau']
    
    df_kg_nhan['Số lượng chuyển_clean'] = to_numeric(df_kg_nhan['Số lượng chuyển'])
    df_kg_nhan['Row_Tong_Tra_KR'] = to_numeric(df_kg_nhan['Hao hụt']) + to_numeric(df_kg_nhan['Kho_Rau'])
    df_kg_nhan['Row_Tong_Tra_KR_GT'] = to_numeric(df_kg_nhan['Tổng hao hụt']) + to_numeric(df_kg_nhan['Tổng kho rau'])
    
    df_kg_nhan['Tỷ lệ % lệch'] = np.where(df_kg_nhan['Số lượng chuyển_clean'] > 0, (df_kg_nhan['Row_Tong_Tra_KR'] / df_kg_nhan['Số lượng chuyển_clean']) * 100, 100)
    
    conditions = [
        (df_kg_nhan['Tỷ lệ % lệch'] == 0),
        (df_kg_nhan['Tỷ lệ % lệch'] > 0) & (df_kg_nhan['Tỷ lệ % lệch'] <= 5),
        (df_kg_nhan['Tỷ lệ % lệch'] > 5) & (df_kg_nhan['Tỷ lệ % lệch'] <= 10),
        (df_kg_nhan['Tỷ lệ % lệch'] > 10) & (df_kg_nhan['Tỷ lệ % lệch'] <= 15),
        (df_kg_nhan['Tỷ lệ % lệch'] > 15)
    ]
    choices = ['0%', '<= 5%', '5-10%', '10-15%', '> 15%']
    df_kg_nhan['Nhóm lệch'] = np.select(conditions, choices, default='> 15%')
    
    pivot_bucket = df_kg_nhan[df_kg_nhan['Nhóm lệch'] != '0%'].pivot_table(index='CLV4', columns='Nhóm lệch', values='Row_Tong_Tra_KR', aggfunc='sum', fill_value=0).reset_index()
    pivot_bucket_gt = df_kg_nhan[df_kg_nhan['Nhóm lệch'] != '0%'].pivot_table(index='CLV4', columns='Nhóm lệch', values='Row_Tong_Tra_KR_GT', aggfunc='sum', fill_value=0).reset_index()
    pivot_bucket_gt.rename(columns={c: f'GT {c}' for c in ['<= 5%', '5-10%', '10-15%', '> 15%']}, inplace=True)
    
    df_b21_new = pd.merge(df_b21_new, pivot_bucket, on='CLV4', how='left').fillna(0)
    df_b21_new = pd.merge(df_b21_new, pivot_bucket_gt, on='CLV4', how='left').fillna(0)
    
    for c in ['<= 5%', '5-10%', '10-15%', '> 15%']:
        if c not in df_b21_new.columns: df_b21_new[c] = 0
    for c in [f'GT {k}' for k in ['<= 5%', '5-10%', '10-15%', '> 15%']]:
        if c not in df_b21_new.columns: df_b21_new[c] = 0
        
    cols2_1 = ['CLV4', 'SL chuyển', 'SL chênh lệch', 'Lỗi ST (Nhập thiếu)', 'Lỗi ST (Sai QT)', 'Tổng Trả Kho Rau', 'Hao hụt (<=10%)', 'Trả KR (Lỗi giao thiếu)', '<= 5%', '5-10%', '10-15%', '> 15%', 'Đang xử lý', 'Chưa xử lý', 'Không xử lý (WRITE OFF)']
    cols2_1_gt = ['CLV4', 'SL chuyển', 'GT chênh lệch', 'GT Lỗi ST (Nhập thiếu)', 'GT Lỗi ST (Sai QT)', 'GT Tổng Trả Kho Rau', 'GT Hao hụt (<=10%)', 'GT Trả KR (Lỗi giao thiếu)', 'GT <= 5%', 'GT 5-10%', 'GT 10-15%', 'GT > 15%', 'GT Đang xử lý', 'GT Chưa xử lý', 'Giá trị WRITE OFF']
    
    t21_sl, t21_gt = st.tabs(["📊 Số Lượng", "💰 Giá Trị"])
    with t21_sl:
        display_daily_table(df_b21_new, cols2_1, "Bang_2_1_CLV4", group_by_col='CLV4')
    with t21_gt:
        display_daily_table(df_b21_new, cols2_1_gt, "Bang_2_1_CLV4_GT", group_by_col='CLV4')

    nx_b21_new = generate_insights(df_kg_nhan, "Bảng 2.1_New", df_metrics=df_b21_new, date_str=selected_daily_date)
    st.text_area("Nhận xét Bảng 2.1:", value=nx_b21_new, key="nx_b21_new", height=160)

    # 2.2 Hàng Pack
    st.write("---")
    st.subheader("Bảng 3: Đánh giá theo tình hình hàng Pack")
    df_pack = df_filtered[df_filtered['Loại hàng'].astype(str).str.upper() == 'PACK'].copy()
    df_b3 = calculate_daily_metrics(df_pack)
    
    t3_sl, t3_gt = st.tabs(["📊 Số Lượng", "💰 Giá Trị"])
    with t3_sl:
        display_daily_table(df_b3, cols, "Bang_3")
    with t3_gt:
        display_daily_table(df_b3, cols_gt, "Bang_3_GT")
        
    nx_b3 = generate_insights(df_pack, "Bảng 3", df_b3)
    st.text_area("Nhận xét Bảng 3:", value=nx_b3, key="nx_b3", height=120)

# ==========================================
# TRANG 3: TIẾN ĐỘ DC PHẢN HỒI
# ==========================================
with tab_dc:
    st.header("👨‍🔧 Theo Dõi Tiến Độ Xử Lý & Phản Hồi Của DC")
    render_dc_feedback_progress_report(df_active, "Tab_3")
