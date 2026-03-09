import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="報價單自動合併小工具", page_icon="📦", layout="wide")
st.title("📦 報價單自動合併小工具 (新版)")
st.write("支援舊版獨立報價單，以及全新 **DT BUY TRIP 統一格式**！只要上傳包含多個分頁的 Excel，系統會自動幫您抓取全部資料合併！")

# 最終 Master Sheet 需要的欄位
master_cols = ['#', 'Vendor:', 'Vendor Stock #:', 'Item Description:', 'FOB:', 
               'Product Size (In/Cm):', 'Material Content All Parts:', 'Part Counts/Specs:', 
               'Master Pk:', 'Inner Pk:', 'Unit Package Type & Quality:', 'Display Type / Quality:', 
               'Notes / MOQ / Leadtime:', 'New Item ', 'Existing SKU#: ', 'Country of Origin: ', 
               'ELC:', 'Weight (G/Lb):', 'Size:']

def format_size(s):
    """協助把長寬高數字格式化到小數點第二位"""
    try: return f"{float(s):.2f}"
    except: return "" if pd.isna(s) else str(s).strip()

def process_single_df(df_temp, get_full_df_func):
    """處理單一個工作表的邏輯"""
    header_idx = -1
    file_type = None
    
    # 智慧尋找表頭
    for i in range(min(15, len(df_temp))):
        row_str = " ".join(df_temp.iloc[i].astype(str).tolist())
        if '工廠代碼/名稱' in row_str or '產品描述' in row_str:
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
    
    # 模糊搜尋欄位的輔助函數
    def get_col(kw, exact=False):
        if exact:
            matches = [c for c in df.columns if kw == str(c).strip()]
        else:
            matches = [c for c in df.columns if kw in str(c)]
        return matches[0] if matches else None

    rename_dict = {}
    
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
        
        # 處理產品尺寸 (L x W x H) - 會精準避開外箱或內盒的長寬高
        l_col = get_col('L\n(INCH)', exact=True) or get_col('L', exact=False)
        w_col = get_col('W\n(INCH)', exact=True) or get_col('W', exact=False)
        h_col = get_col('H\n(INCH)', exact=True) or get_col('H', exact=False)
        
        if l_col and w_col and h_col:
            l = df[l_col].apply(format_size)
            w = df[w_col].apply(format_size)
            h = df[h_col].apply(format_size)
            size_str = l + ' x ' + w + ' x ' + h
            # 移除空的尺寸組合
            df['Product Size (In/Cm):'] = size_str.apply(lambda x: "" if x.replace('x', '').replace(' ', '') == '' else x)
        
        rename_dict = {
            vendor_col: 'Vendor:',
            item_col: 'Vendor Stock #:',
            desc_col: 'Item Description:',
            fob_col: 'FOB:',
            mat_col: 'Material Content All Parts:',
            inner_col: 'Inner Pk:',
            master_col: 'Master Pk:',
            pkg_col: 'Unit Package Type & Quality:',
            lead_col: 'Notes / MOQ / Leadtime:',
            weight_col: 'Weight (G/Lb):'
        }
        
    elif file_type == 'AD450':
        rename_dict = {
            get_col('廠名/廠號'): 'Vendor:',
            get_col('品名'): 'Item Description:',
            get_col('價格'): 'FOB:',
            get_col('產品尺寸'): 'Product Size (In/Cm):',
            get_col('產品材質'): 'Material Content All Parts:',
            get_col('外箱數量'): 'Master Pk:',
            get_col('內盒數量'): 'Inner Pk:',
            get_col('包裝明細'): 'Unit Package Type & Quality:',
            get_col('MOQ'): 'Notes / MOQ / Leadtime:',
            get_col('產品重量'): 'Weight (G/Lb):'
        }
        
    else: # 其他舊版格式
        rename_dict = {
            get_col('工廠 / factory ID'): 'Vendor:',
            get_col('品名 & 內容描述'): 'Item Description:',
            get_col('價格 (FOB)'): 'FOB:',
            get_col('產品規格尺寸'): 'Product Size (In/Cm):',
            get_col('產品材質'): 'Material Content All Parts:',
            get_col('外箱數量'): 'Master Pk:',
            get_col('內盒數量'): 'Inner Pk:',
            get_col('包裝明細'): 'Unit Package Type & Quality:',
            get_col('MOQ'): 'Notes / MOQ / Leadtime:',
            get_col('產品重量'): 'Weight (G/Lb):'
        }
        
    # 清除沒找到的空欄位，並重新命名
    rename_dict = {k: v for k, v in rename_dict.items() if k is not None}
    df = df.rename(columns=rename_dict)
    
    # 僅留下 Master Sheet 需要的欄位
    return df[[c for c in master_cols if c in df.columns]]


def process_files(uploaded_files):
    all_dataframes = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df_temp = pd.read_csv(file, nrows=15, header=None)
                def get_full_csv(h_idx):
                    file.seek(0)
                    return pd.read_csv(file, header=h_idx)
                
                res_df = process_single_df(df_temp, get_full_csv)
                if res_df is not None:
                    all_dataframes.append(res_df)
                    
            else:
                excel_file = pd.ExcelFile(file)
                # 自動歷遍 Excel 內的所有分頁
                for sheet_name in excel_file.sheet_names:
                    df_temp = excel_file.parse(sheet_name=sheet_name, nrows=15, header=None)
                    def get_full_excel(h_idx, s_name=sheet_name):
                        return excel_file.parse(sheet_name=s_name, header=h_idx)
                        
                    res_df = process_single_df(df_temp, get_full_excel)
                    if res_df is not None:
                        all_dataframes.append(res_df)
                        
        except Exception as e:
            st.error(f"處理檔案 {file.name} 時發生錯誤: {e}")
            
    if not all_dataframes: return None
        
    master_df = pd.concat(all_dataframes, ignore_index=True)
    # 補齊可能缺失的空欄位
    for col in master_cols:
        if col not in master_df.columns:
            master_df[col] = ''
            
    master_df = master_df[master_cols]
    master_df = master_df.dropna(subset=['Item Description:', 'Vendor:'], how='all')
    
    return master_df

uploaded_files = st.file_uploader("📂 點擊或拖曳上傳報價單", accept_multiple_files=True, type=['xlsx', 'xls', 'csv'])

if st.button("🚀 產生 Master Sheet 總表", type="primary"):
    if not uploaded_files:
        st.warning("請先上傳檔案喔！")
    else:
        with st.spinner("資料處理中，請稍候..."):
            result_df = process_files(uploaded_files)
            
            if result_df is not None and not result_df.empty:
                st.success(f"轉換成功！共處理了 {len(result_df)} 筆商品。")
                st.dataframe(result_df)
                csv = result_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 下載 Master_Sheet.csv", data=csv, file_name="Master_Sheet.csv", mime="text/csv")
            else:
                st.warning("沒有找到符合格式的報價資料，請確認檔案內容。")