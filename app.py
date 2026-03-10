import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="報價單自動合併小工具", page_icon="📦", layout="wide")
st.title("📦 報價單萬用合併平台 (Excel 輸出升級版)")
st.write("請先選擇您要處理的專案類型，再上傳多份工廠報價單。系統將自動產出可直接貼圖片的 Excel 總表格式！")

# ================= 定義兩種輸出的標準欄位 =================
dt_master_cols = [
    '#', 'Vendor:', 'Vendor Stock #:', 'Item Description:', 'FOB:', 
    'Product Size (In/Cm):', 'Material Content All Parts:', 'Part Counts/Specs:', 
    'Master Pk:', 'Inner Pk:', 'Unit Package Type & Quality:', 'Display Type / Quality:', 
    'Notes / MOQ / Leadtime:', 'New Item ', 'Existing SKU#: ', 'Country of Origin: ', 
    'ELC:', 'Weight (G/Lb):', 'Size:'
]

tg_cost_cols = [
    'FACTORY#', 'VENDOR STYLE#', 'CATEGORY', 'DESCRIPTION', 'IMAGES', 
    'QUOTE REMARK', 'LINE PLAN QTY', 'MOLD FEE', 'FCA COST', 'GMI printed ', 
    'GMI-non printed', 'Film plate cost/mould fee', 'Cost of GMI printed package ', 
    'Cost of GIM-non printed', 'Total Cost', 'profit %', 'FCA', 'RETAIL ', 'IMU', 
    'Master\nCasepack\nUnit', 'Master\nLength (in)', 'Master\nWidth (in)', 
    'Master\nHeight (in)', 'Master\nWeight (lb)', 'Inner\nCasepack\nUnit', 
    'Inner\nLength (in)', 'Inner\nWidth (in)', 'Inner\nHeight (in)', 'Inner\nWeight (lb)'
]

def format_size(s):
    try: return f"{float(s):.2f}"
    except: return "" if pd.isna(s) else str(s).strip()

def process_single_df(df_temp, get_full_df_func, workflow):
    header_idx = -1
    file_type = None
    
    # 智慧尋找表頭
    for i in range(min(15, len(df_temp))):
        row_str = " ".join(df_temp.iloc[i].astype(str).tolist())
        if 'Target FTY BPM ID' in row_str or 'FTY Name' in row_str:
            header_idx = i
            file_type = 'TG_MASTER'
            break
        elif '工廠代碼/名稱' in row_str or '產品描述' in row_str:
            header_idx = i
            file_type = 'DT_MASTER'
            break
        elif '廠名/廠號' in row_str:
            header_idx = i
            file_type = 'AD450'
            break
        elif '工廠 / factory ID' in row_str:
            header_idx = i
            file_type = 'Others'
            break
            
    if header_idx == -1: return None
        
    df = get_full_df_func(header_idx)
    if df.empty: return None
    df = df.dropna(how='all')
    
    def get_col(kw, exact=False):
        if exact:
            matches = [c for c in df.columns if kw == str(c).strip()]
        else:
            matches = [c for c in df.columns if kw in str(c)]
        return matches[0] if matches else None

    rename_dict = {}
    
    # ================= 處理流程一：Target (TG) =================
    if workflow == "Target (TG) ➡️ Cost Analysis 成本分析表":
        if file_type != 'TG_MASTER': return None
        
        rename_dict = {
            get_col('FTY Name'): 'FACTORY#',
            get_col('Vendor Style#'): 'VENDOR STYLE#',
            get_col('Line', exact=True) or get_col('Line'): 'CATEGORY',
            get_col('Product Name'): 'DESCRIPTION',
            get_col('Product photo'): 'IMAGES',
            get_col("Line Plan q'ty"): 'LINE PLAN QTY',
            get_col('模具費'): 'MOLD FEE',
            get_col('FCA'): 'FCA COST',
            get_col('客人售價 Retail:') or get_col('Retail\n建議賣價'): 'RETAIL ',
            get_col('外箱數量'): 'Master\nCasepack\nUnit',
            get_col('外箱Length'): 'Master\nLength (in)',
            get_col('外箱Width'): 'Master\nWidth (in)',
            get_col('外箱Height'): 'Master\nHeight (in)',
            get_col('外箱Weight'): 'Master\nWeight (lb)',
            get_col('內箱數量'): 'Inner\nCasepack\nUnit',
            get_col('內箱Length'): 'Inner\nLength (in)',
            get_col('內箱Width'): 'Inner\nWidth (in)',
            get_col('內箱Height'): 'Inner\nHeight (in)'
        }
        
        inner_w_cols = [c for c in df.columns if 'Weight' in str(c) and '外箱' not in str(c)]
        if inner_w_cols:
            rename_dict[inner_w_cols[0]] = 'Inner\nWeight (lb)'
            
        rename_dict = {k: v for k, v in rename_dict.items() if k is not None}
        df = df.rename(columns=rename_dict)
        return df[[c for c in tg_cost_cols if c in df.columns]]

    # ================= 處理流程二：Dollar Tree (DT) =================
    elif workflow == "Dollar Tree (DT) ➡️ Master Sheet 總表":
        if file_type == 'TG_MASTER': return None
        
        if file_type == 'DT_MASTER':
            vendor_col = get_col('工廠代碼')
            desc_col = get_col('產品描述')
            fob_col = get_col('FOB                US$') or get_col('FOB')
            weight_col = get_col('產品重量')
            mat_col = get_col('材質分析')
            inner_col = get_col('內盒數量')
            master_col = get_col('外箱數量')
            pkg_col = get_col('包裝明細')
            lead_col = get_col('大貨生產天數')
            item_col = get_col('產品編號')
            
            l_col = get_col('L\n(INCH)', exact=True) or get_col('L', exact=False)
            w_col = get_col('W\n(INCH)', exact=True) or get_col('W', exact=False)
            h_col = get_col('H\n(INCH)', exact=True) or get_col('H', exact=False)
            
            if l_col and w_col and h_col:
                l = df[l_col].apply(format_size)
                w = df[w_col].apply(format_size)
                h = df[h_col].apply(format_size)
                size_str = l + ' x ' + w + ' x ' + h
                df['Product Size (In/Cm):'] = size_str.apply(lambda x: "" if x.replace('x', '').replace(' ', '') == '' else x)
            
            rename_dict = {
                vendor_col: 'Vendor:', item_col: 'Vendor Stock #:', desc_col: 'Item Description:',
                fob_col: 'FOB:', mat_col: 'Material Content All Parts:', inner_col: 'Inner Pk:',
                master_col: 'Master Pk:', pkg_col: 'Unit Package Type & Quality:', lead_col: 'Notes / MOQ / Leadtime:', weight_col: 'Weight (G/Lb):'
            }
        elif file_type == 'AD450':
            rename_dict = {
                get_col('廠名/廠號'): 'Vendor:', get_col('品名'): 'Item Description:', get_col('價格'): 'FOB:',
                get_col('產品尺寸'): 'Product Size (In/Cm):', get_col('產品材質'): 'Material Content All Parts:',
                get_col('外箱數量'): 'Master Pk:', get_col('內盒數量'): 'Inner Pk:', get_col('包裝明細'): 'Unit Package Type & Quality:',
                get_col('MOQ'): 'Notes / MOQ / Leadtime:', get_col('產品重量'): 'Weight (G/Lb):'
            }
        else:
            rename_dict = {
                get_col('工廠 / factory ID'): 'Vendor:', get_col('品名 & 內容描述'): 'Item Description:',
                get_col('價格 (FOB)'): 'FOB:', get_col('產品規格尺寸'): 'Product Size (In/Cm):',
                get_col('產品材質'): 'Material Content All Parts:', get_col('外箱數量'): 'Master Pk:',
                get_col('內盒數量'): 'Inner Pk:', get_col('包裝明細'): 'Unit Package Type & Quality:',
                get_col('MOQ'): 'Notes / MOQ / Leadtime:', get_col('產品重量'): 'Weight (G/Lb):'
            }
            
        rename_dict = {k: v for k, v in rename_dict.items() if k is not None}
        df = df.rename(columns=rename_dict)
        return df[[c for c in dt_master_cols if c in df.columns]]

def process_files(uploaded_files, workflow):
    all_dataframes = []
    target_cols = tg_cost_cols if "Target" in workflow else dt_master_cols
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df_temp = pd.read_csv(file, nrows=15, header=None)
                def get_full_csv(h_idx):
                    file.seek(0)
                    return pd.read_csv(file, header=h_idx)
                
                res_df = process_single_df(df_temp, get_full_csv, workflow)
                if res_df is not None:
                    all_dataframes.append(res_df)
            else:
                excel_file = pd.ExcelFile(file)
                for sheet_name in excel_file.sheet_names:
                    df_temp = excel_file.parse(sheet_name=sheet_name, nrows=15, header=None)
                    def get_full_excel(h_idx, s_name=sheet_name):
                        return excel_file.parse(sheet_name=s_name, header=h_idx)
                        
                    res_df = process_single_df(df_temp, get_full_excel, workflow)
                    if res_df is not None:
                        all_dataframes.append(res_df)
                        
        except Exception as e:
            st.error(f"處理檔案 {file.name} 時發生錯誤: {e}")
            
    if not all_dataframes: return None
        
    master_df = pd.concat(all_dataframes, ignore_index=True)
    
    for col in target_cols:
        if col not in master_df.columns:
            master_df[col] = None
            
    master_df = master_df[target_cols]
    
    if "Target" in workflow:
        master_df = master_df.dropna(subset=['DESCRIPTION', 'FACTORY#'], how='all')
    else:
        master_df = master_df.dropna(subset=['Item Description:', 'Vendor:'], how='all')
    
    return master_df

# ================= 網頁前端介面 =================
st.markdown("### 1️⃣ 選擇專案工作流程")
selected_workflow = st.selectbox("請選擇要執行的報價單轉換流程：", [
    "Target (TG) ➡️ Cost Analysis 成本分析表",
    "Dollar Tree (DT) ➡️ Master Sheet 總表"
])

st.markdown("### 2️⃣ 上傳檔案")
uploaded_files = st.file_uploader("📂 將工廠報價單拖曳至此", accept_multiple_files=True, type=['xlsx', 'xls', 'csv'])

if st.button("🚀 開始彙總產出 (Excel 格式)", type="primary"):
    if not uploaded_files:
        st.warning("請先上傳檔案喔！")
    else:
        with st.spinner("資料處理中，請稍候..."):
            result_df = process_files(uploaded_files, selected_workflow)
            
            if result_df is not None and not result_df.empty:
                st.success(f"轉換成功！共處理了 {len(result_df)} 筆商品資料。")
                st.dataframe(result_df)
                
                # ====== 產出真正的 Excel 檔案 (.xlsx) ======
                file_name = "Cost_Analysis_Data.xlsx" if "Target" in selected_workflow else "DT_Master_Sheet.xlsx"
                
                # 使用虛擬記憶體來存放 Excel 檔案，避免在雲端伺服器留下實體檔案
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Master Data')
                
                excel_data = buffer.getvalue()
                
                # 提供 Excel 格式的下載按鈕
                st.download_button(
                    label=f"📥 下載 {file_name}", 
                    data=excel_data, 
                    file_name=file_name, 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("沒有找到符合您所選專案格式的報價資料，請確認「工作流程」是否選擇正確，或檢查報價單內容。")

