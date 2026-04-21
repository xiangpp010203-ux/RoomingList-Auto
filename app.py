import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os  # ★ 新增：用來處理檔案名稱的模組

st.set_page_config(page_title="OP 救星：分房表轉換神器", page_icon="🏨")
st.title("🏨 OP 救星：SCM_Rooming List 自動轉換神器")
st.write("支援上傳 .xlsx 與 .csv 格式，智慧閃避合併儲存格並自動命名！")

# ==========================================
# 安全寫入函式 (閃避合併儲存格)
# ==========================================
def safe_write(sheet, r, c, val):
    cell = sheet.cell(row=r, column=c)
    if type(cell).__name__ != 'MergedCell':
        cell.value = val

# ==========================================
# 上傳與主程式區塊
# ==========================================
uploaded_file_B = st.file_uploader("請上傳【內部系統分房表 (檔案B)】支援 Excel 或 CSV", type=["xlsx", "csv"])

if uploaded_file_B is not None:
    st.success(f"成功讀取檔案：{uploaded_file_B.name}！正在解析欄位...")
    
    try:
        # --- ★ 新增：動態檔名產生邏輯 ---
        # 取得上傳檔案的名稱，並使用 os.path.splitext 將主檔名與副檔名分開
        # 例如："ABC.xlsx" -> base_name="ABC", ext=".xlsx"
        base_name, ext = os.path.splitext(uploaded_file_B.name)
        # 組合出新的下載檔名
        output_filename = f"{base_name}_RoomingList.xlsx"
        
        # 智慧尋找標題列
        if uploaded_file_B.name.endswith('.csv'):
            df_temp = pd.read_csv(uploaded_file_B, header=None)
            header_idx = 0
            for i in range(min(5, len(df_temp))):
                if '房號' in str(df_temp.iloc[i].values) or '英文姓名' in str(df_temp.iloc[i].values):
                    header_idx = i
                    break
            uploaded_file_B.seek(0) 
            df_B = pd.read_csv(uploaded_file_B, header=header_idx)
            
        else: # 處理 xlsx
            df_temp = pd.read_excel(uploaded_file_B, header=None, engine='openpyxl')
            header_idx = 0
            for i in range(min(5, len(df_temp))):
                if '房號' in str(df_temp.iloc[i].values) or '英文姓名' in str(df_temp.iloc[i].values):
                    header_idx = i
                    break
            uploaded_file_B.seek(0)
            df_B = pd.read_excel(uploaded_file_B, header=header_idx, engine='openpyxl')

        df_B.columns = df_B.columns.str.strip()
        
        # 讀取飯店底稿
        wb = openpyxl.load_workbook("Template.xlsx")
        sheet = wb.active 
        
        # 修正：將起始列改為第 14 列，徹底避開綠底範例
        start_row = 14 
        
        for index, row in df_B.iterrows():
            if pd.isna(row.get('房號')):
                continue

            current_row = start_row + index
            
            # --- 姓名拆解 ---
            eng_name_raw = str(row.get('英文姓名', ''))
            title, last_name, first_name = "", "", ""
            
            if "MR " in eng_name_raw: title, eng_name_raw = "Mr", eng_name_raw.replace("MR ", "")
            elif "MS " in eng_name_raw: title, eng_name_raw = "Ms", eng_name_raw.replace("MS ", "")
            elif "MISS " in eng_name_raw: title, eng_name_raw = "Miss", eng_name_raw.replace("MISS ", "")
            
            if "/" in eng_name_raw:
                last_name = eng_name_raw.split("/")[0].strip()
                first_name = eng_name_raw.split("/")[1].strip()

            # --- 生日格式調整為 YYYY/MM/DD ---
            try:
                dob_raw = str(int(float(row.get('生日', '')))) 
            except:
                dob_raw = str(row.get('生日', ''))

            if len(dob_raw) == 8 and dob_raw.isdigit():
                dob_formatted = f"{dob_raw[0:4]}/{dob_raw[4:6]}/{dob_raw[6:8]}"
            else:
                dob_formatted = dob_raw
                
            # --- 中文姓名拆解 ---
            cht_name = str(row.get('中文姓名', ''))
            cht_name = cht_name if cht_name != 'nan' else ""
            cht_last, cht_first = "", ""
            if len(cht_name) >= 2:
                cht_last = cht_name[0]      
                cht_first = cht_name[1:]    
            
            # --- 寫入資料 ---
            safe_write(sheet, current_row, 1, row.get('房號', ''))
            safe_write(sheet, current_row, 2, row.get('No', ''))      
            safe_write(sheet, current_row, 3, title)                  
            safe_write(sheet, current_row, 4, last_name)              
            safe_write(sheet, current_row, 5, first_name)             
            safe_write(sheet, current_row, 6, cht_last)               
            safe_write(sheet, current_row, 7, cht_first)              
            safe_write(sheet, current_row, 10, row.get('護照號碼', ''))
            safe_write(sheet, current_row, 11, dob_formatted)         
            
            remark = str(row.get('備註', ''))
            remark = remark if remark != 'nan' else ""
            safe_write(sheet, current_row, 21, remark)                
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.balloons() 
        st.success(f"🎉 資料轉換成功！已為您生成專屬檔名：{output_filename}")
        
        # ★ 修正：將按鈕上的文字與下載的檔名，都替換成我們剛剛組合好的 output_filename
        st.download_button(
            label=f"📥 下載最終名單 ({output_filename})",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"系統發出微弱的求救訊號：\n{e}")