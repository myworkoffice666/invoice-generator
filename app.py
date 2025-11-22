import streamlit as st
import pandas as pd
import openpyxl
import os
import tempfile
import zipfile
from pathlib import Path

# ==========================================
# 1. 頁面配置
# ==========================================
st.set_page_config(page_title="顧問發票自動生成器", page_icon="📄", layout="wide")

st.title("📄 顧問發票自動生成系統 (雲端版)")
st.markdown("""
本系統協助您將三個月份的顧問資料合併，自動計算費用，並生成發票格式 (Excel)。
- **支援功能**：資料清洗、自動合併、費用計算、批量生成 Excel 發票。
- **輸出格式**：`.xlsx` (請下載後自行另存為 PDF)。
""")

# ==========================================
# 2. 側邊欄：檔案上傳
# ==========================================
st.sidebar.header("📂 1. 上傳檔案")

uploaded_file_1 = st.sidebar.file_uploader("上傳 第一個檔案 (July)", type=["xls", "xlsx"])
uploaded_file_2 = st.sidebar.file_uploader("上傳 第二個檔案 (August)", type=["xls", "xlsx"])
uploaded_file_3 = st.sidebar.file_uploader("上傳 第三個檔案 (September)", type=["xls", "xlsx"])

st.sidebar.header("📄 2. 上傳模板")
uploaded_template = st.sidebar.file_uploader("上傳發票模板 (CF_template.xlsx)", type=["xlsx"])

# 參數設定
st.sidebar.header("⚙️ 3. 參數設定")
EVALUATION_PERIOD = st.sidebar.text_input("Evaluation Period", value='07/01/2025 - 09/30/2025')

# ==========================================
# 3. 核心邏輯
# ==========================================

# 模板映射定義
DATA_TEMPLATE_MAPPING = [
    (1, "D12:E12", True),  (2, "D14:E14", True),  (3, "A7:F7", True),    
    (4, "D11:E11", True),  (5, "D13:E13", True),  (6, "A5:F5", True),    
    (7, "A8:F8", True),    (8, "A16:F16", True),
    (9, "B18:B18", False), (10, "C18:C18", False), (11, "D18:D18", False), (12, "E18:E18", False),
    (13, "B19:B19", False), (14, "C19:C19", False), (15, "D19:D19", False), (16, "E19:E19", False),
    (17, "B20:B20", False), (18, "C20:C20", False), (19, "D20:D20", False), (20, "E20:E20", False),
    (21, "E21:E21", False)
]

def process_data_streamlit(files_map):
    """讀取並處理資料"""
    dfs = []
    
    # 定義內部讀取函數
    def load_and_clean(file_obj, date_label):
        try:
            # Pandas 可以直接讀取 UploadedFile 物件
            df = pd.read_excel(file_obj, index_col=1, skiprows=6).iloc[:, 1:]
            
            # [修正 1] 移除欄位名稱的空白
            df.columns = df.columns.str.strip()
            
            # 安全檢查：確保有 Advisor 欄位
            if "Advisor" in df.columns:
                df = df.loc[~df["Advisor"].isna()]
                df = df.loc[df["Advisor"] != "Advisor"]
            
            df["Date"] = date_label
            return df
        except Exception as e:
            st.error(f"讀取錯誤 ({date_label}): {e}")
            return pd.DataFrame()

    # 依序讀取
    for label, file_obj in files_map.items():
        if file_obj is not None:
            dfs.append(load_and_clean(file_obj, label))
    
    if not dfs:
        return pd.DataFrame()

    all_data = pd.concat(dfs, axis=0, ignore_index=False).reset_index()
    if 'index' in all_data.columns:
        all_data.rename(columns={'index': 'Client'}, inplace=True)
    
    # 再次確保所有欄位去空白
    all_data.columns = all_data.columns.str.strip()

    # --- 資料清洗：將 Fee 與 Balance 轉為數字 ---
    cols_to_clean = ['Fee', 'Average Daily Balance']
    for col in cols_to_clean:
        if col in all_data.columns:
            all_data[col] = all_data[col].astype(str).str.replace(r'[$,]', '', regex=True)
            all_data[col] = pd.to_numeric(all_data[col], errors='coerce').fillna(0)
    # -----------------------------------------------------

    target_col = 'Client'
    # 檢查目標欄位是否存在
    if target_col not in all_data.columns:
        st.error(f"找不到 '{target_col}' 欄位，請檢查 Excel 格式。")
        return pd.DataFrame()

    all_data['count'] = all_data.groupby(target_col)[target_col].transform('count')
    df_exact_3 = all_data[all_data['count'] == 3].copy()
    
    # 處理不完整資料提示
    df_others = all_data[all_data['count'] != 3].copy()
    if not df_others.empty:
        st.warning(f"⚠️ 發現 {len(df_others)} 筆資料因非完整 3 個月而被排除 (Client: {df_others['Client'].unique()})")

    if df_exact_3.empty:
        st.error("❌ 沒有發現剛好 3 筆資料的客戶。")
        return pd.DataFrame()

    # Pivot 轉換
    df_exact_3['period_id'] = df_exact_3.groupby(target_col).cumcount() + 1
    fixed_cols = ['Client', 'Advisor', 'Unique Client ID']
    # 確保這些欄位存在
    fixed_cols = [c for c in fixed_cols if c in df_exact_3.columns]
    
    value_cols = ['Average Daily Balance', 'Days in Period', 'Fee', 'Date']
    
    df_wide = df_exact_3.pivot(index=fixed_cols, columns='period_id', values=value_cols)
    df_wide.columns = [f'{col[0]}{col[1]}' for col in df_wide.columns]
    df_wide = df_wide.reset_index()

    desired_columns = [
        'Client', 'Advisor', 'Unique Client ID',
        'Average Daily Balance1', 'Average Daily Balance2', 'Average Daily Balance3',
        'Days in Period1', 'Days in Period2', 'Days in Period3',
        'Fee1', 'Fee2', 'Fee3',
        'Date1', 'Date2', 'Date3'
    ]
    final_cols = [c for c in desired_columns if c in df_wide.columns]
    df_wide = df_wide[final_cols]
    
    # --- [修正 2] 終極防呆清洗：計算前再次確保 Fee1, Fee2, Fee3 是數字 ---
    for fee_col in ["Fee1", "Fee2", "Fee3"]:
        if fee_col in df_wide.columns:
            df_wide[fee_col] = pd.to_numeric(
                df_wide[fee_col].astype(str).str.replace(r'[$,]', '', regex=True), 
                errors='coerce'
            ).fillna(0)
    # -------------------------------------------------------------

    # 計算總和
    try:
        df_wide["Total"] = (df_wide.get("Fee1", 0) + df_wide.get("Fee2", 0) + df_wide.get("Fee3", 0)).round(2)
    except Exception as e:
        st.error(f"計算總金額時發生錯誤: {e}")
        df_wide["Total"] = 0

    df_wide["Eval"] = EVALUATION_PERIOD

    return df_wide

def generate_invoices_streamlit(df, template_path, output_dir):
    """生成 Excel 發票"""
    xlsx_dir = Path(output_dir) / "XLSX"
    xlsx_dir.mkdir(parents=True, exist_ok=True)
    
    generated_files = []
    
    progress_bar = st.progress(0)
    total_rows = len(df)
    
    for idx, row in enumerate(df.itertuples(index=False)):
        # 安全獲取欄位資料
        Client = getattr(row, "Client", "Unknown")
        Unique_Client_ID = getattr(row, "Unique_Client_ID", getattr(row, "_2", "")) 
        
        avg1 = getattr(row, "Average_Daily_Balance1", 0)
        avg2 = getattr(row, "Average_Daily_Balance2", 0)
        avg3 = getattr(row, "Average_Daily_Balance3", 0)
        
        days1 = getattr(row, "Days_in_Period1", 0)
        days2 = getattr(row, "Days_in_Period2", 0)
        days3 = getattr(row, "Days_in_Period3", 0)
        
        fee1 = getattr(row, "Fee1", 0)
        fee2 = getattr(row, "Fee2", 0)
        fee3 = getattr(row, "Fee3", 0)
        
        date1 = getattr(row, "Date1", "")
        date2 = getattr(row, "Date2", "")
        date3 = getattr(row, "Date3", "")
        
        Total = getattr(row, "Total", 0)
        Eval = getattr(row, "Eval", "")

        template_data = [
            Eval, f"${Total:,.2f}", f"Client Name(s): {Client}", str(Unique_Client_ID)[:10],
            "0.25%", f"Billing Cycle: {Eval}", "Address: ????", f"Fee Calculation {str(Unique_Client_ID)[:10]}",
            date1, avg1, days1, f"${fee1:,.2f}",
            date2, avg2, days2, f"${fee2:,.2f}",
            date3, avg3, days3, f"${fee3:,.2f}",
            f"${Total:,.2f}"
        ]

        output_path = xlsx_dir / f"CF_invoice_{Client}.xlsx"
        
        try:
            # 讀取模板並填入
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            for i, mapping in enumerate(DATA_TEMPLATE_MAPPING):
                index, cell_range, is_merged = mapping
                val = template_data[i]
                
                top_left = cell_range.split(':')[0]
                if is_merged:
                    try: ws.merge_cells(cell_range)
                    except ValueError: pass
                ws[top_left] = val
            
            wb.save(output_path)
            generated_files.append(output_path)
        except Exception as e:
            st.error(f"生成 Excel 失敗 {Client}: {e}")
        
        progress_bar.progress((idx + 1) / total_rows)
        
    return generated_files

def make_zip(source_dirs, output_filename):
    """將資料夾打包成 ZIP"""
    zip_path = Path(output_filename)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for folder in source_dirs:
            folder_path = Path(folder)
            if folder_path.exists():
                for file in folder_path.glob('*'):
                    zipf.write(file, arcname=f"{folder_path.name}/{file.name}")
    return zip_path

# ==========================================
# 4. 主執行流程
# ==========================================

start_button = st.sidebar.button("🚀 開始處理", type="primary")

if start_button:
    # 檢查檔案是否齊全
    if not (uploaded_file_1 and uploaded_file_2 and uploaded_file_3 and uploaded_template):
        st.error("請先上傳所有必要的檔案 (3個月份資料 + 1個模板)。")
    else:
        # 建立臨時工作目錄
        with tempfile.TemporaryDirectory() as tmpdirname:
            st.info(f"工作目錄已建立: {tmpdirname}")
            
            # 1. 儲存模板到臨時目錄
            temp_template_path = os.path.join(tmpdirname, "template.xlsx")
            with open(temp_template_path, "wb") as f:
                f.write(uploaded_template.getbuffer())
            
            # 2. 處理資料
            files_map = {
                'Jul 2025': uploaded_file_1,
                'Aug 2025': uploaded_file_2,
                'Sep 2025': uploaded_file_3
            }
            
            with st.spinner('Step 1: 正在讀取並合併資料...'):
                df_result = process_data_streamlit(files_map)
            
            if not df_result.empty:
                st.success(f"資料處理完成！共 {len(df_result)} 位合格客戶。")
                with st.expander("查看處理後的數據預覽"):
                    st.dataframe(df_result)
                
                # 3. 生成 Excel
                xlsx_output_dir = os.path.join(tmpdirname, "XLSX")
                with st.spinner('Step 2: 正在生成 Excel 發票...'):
                    generated_xlsx = generate_invoices_streamlit(df_result, temp_template_path, tmpdirname)
                
                st.success(f"已生成 {len(generated_xlsx)} 份 Excel 發票。")
                
                # 4. 打包下載 (只打包 XLSX)
                with st.spinner('正在打包檔案...'):
                    dirs_to_zip = [xlsx_output_dir]
                    
                    zip_filename = os.path.join(tmpdirname, "invoices_result.zip")
                    zip_path = make_zip(dirs_to_zip, zip_filename)
                    
                    # 讀取 ZIP 準備下載
                    with open(zip_path, "rb") as f:
                        zip_data = f.read()
                        
                    st.balloons()
                    st.header("🎉 處理完成！")
                    st.download_button(
                        label="📥 下載完整壓縮包 (Excel ZIP)",
                        data=zip_data,
                        file_name="consultant_invoices_xlsx.zip",
                        mime="application/zip"
                    )
            else:
                st.warning("沒有產生任何數據，請檢查上傳的檔案內容。")

st.markdown("---")
st.caption("Powered by Streamlit & Python | Designed for CF TransGlobal")