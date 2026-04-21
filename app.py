import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os

st.set_page_config(page_title="OP 救星：分房表轉換神器", page_icon="🏨")
st.title("🏨 OP 救星：SCM_Rooming List 自動轉換神器")
st.write("支援單人房留空、三人房動態新增列、自動命名！")

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
    st.success(f"成功讀取檔案：{uploaded_file_B.name}！正在進行智能分房排版...")
    
    try:
        # 動態檔名產生
        base_name, ext = os.path.splitext(uploaded_file_B.name)
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
        else: 
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
        
        start_row = 14 
        current_row = start_row
        
        # ==========================================
        # ★ 升級核心：房號群組化邏輯處理
        # ==========================================
        # 過濾掉完全沒有房號的空白列
        df_B_valid = df_B.dropna(subset=['房號'])
        
        # 為了保持名單原本的排序，我們把出現過的房號依序抓出來
        room_numbers = []
        for r in df_B_valid['房號']:
            if pd.notna(r) and r not in room_numbers:
                room_numbers.append(r)
                
        # 開始一間房間、一間房間處理
        for room_no in room_numbers:
            # 抓出這間房裡面的所有旅客名單
            room_group = df_B_valid[df_B_valid['房號'] == room_no]
            passengers = room_group.to_dict('records')
            num_people = len(passengers)
            
            # ★ 邏輯 2：預設每間房至少佔用 2 列 (若是單人房，第2列也會保留給它空著)
            rows_to_occupy = max(2, num_people)
            
            for i in range(rows_to_occupy):
                # ★ 邏輯 3：如果這間房超過 2 個人 (如第3人、第4人)，就在下方「動態插入新的一列」
                if i >= 2:
                    sheet.insert_rows(current_row)
                
                # 如果這個迴圈的 index 有對應的旅客 (不是單人房硬擠出來的空列)
                if i < num_people:
                    row = passengers[i]
                    
                    # --- 姓名拆解 ---
                    eng_name_raw = str(row.get('英文姓名', '')).strip()
                    if eng_name_raw == 'nan': eng_name_raw = ""
                    title, last_name, first_name = "", "", ""
                    
                    if "MR " in eng_name_raw: title, eng_name_raw = "Mr", eng_name_raw.replace("MR ", "")
                    elif "MS " in eng_name_raw: title, eng_name_raw = "Ms", eng_name_raw.replace("MS ", "")
                    elif "MISS " in eng_name_raw: title, eng_name_raw = "Miss", eng_name_raw.replace("MISS ", "")
                    elif "MSTR " in eng_name_raw: title, eng_name_raw = "Mstr", eng_name_raw.replace("MSTR ", "")
                    
                    if "/" in eng_name_raw:
                        last_name = eng_name_raw.split("/")[0].strip()
                        first_name = eng_name_raw.split("/")[1].strip()
                    else:
                        last_name = eng_name_raw

                    # --- 生日 ---
                    try:
                        dob_raw = str(int(float(row.get('生日', '')))) 
                    except:
                        dob_raw = str(row.get('生日', ''))

                    if len(dob_raw) == 8 and dob_raw.isdigit():
                        dob_formatted = f"{dob_raw[0:4]}/{dob_raw[4:6]}/{dob_raw[6:8]}"
                    else:
                        dob_formatted = dob_raw if dob_raw != 'nan' else ""
                        
                    # --- 中文姓名 ---
                    cht_name = str(row.get('中文姓名', ''))
                    cht_name = cht_name if cht_name != 'nan' else ""
                    cht_last, cht_first = "", ""
                    if len(cht_name) >= 2:
                        cht_last = cht_name[0]      
                        cht_first = cht_name[1:]    
                    elif len(cht_name) == 1:
                        cht_last = cht_name[0]
                        
                    # --- ★ 邏輯 1：客房編號與Guest No處理 ---
                    # 房號只在該房間的第一筆 (i==0) 顯示
                    try:
                        room_val = str(int(float(room_no))) if i == 0 else ""
                    except:
                        room_val = str(room_no) if i == 0 else ""
                        
                    # Guest No 直接抓取原始檔案的 No，去掉小數點
                    no_raw = row.get('No', '')
                    try:
                        no_val = str(int(float(no_raw))) if pd.notna(no_raw) and str(no_raw) != 'nan' else ""
                    except:
                        no_val = str(no_raw) if pd.notna(no_raw) and str(no_raw) != 'nan' else ""
                        
                    # 護照號碼 (避免 Excel 自動轉成浮點數如 367496721.0)
                    passport = str(row.get('護照號碼', ''))
                    if passport.endswith('.0'): passport = passport[:-2]
                    if passport == 'nan': passport = ""
                    
                    remark = str(row.get('備註', ''))
                    remark = remark if remark != 'nan' else ""

                    # --- 寫入資料 ---
                    safe_write(sheet, current_row, 1, room_val)
                    safe_write(sheet, current_row, 2, no_val)
                    safe_write(sheet, current_row, 3, title)
                    safe_write(sheet, current_row, 4, last_name)
                    safe_write(sheet, current_row, 5, first_name)
                    safe_write(sheet, current_row, 6, cht_last)
                    safe_write(sheet, current_row, 7, cht_first)
                    safe_write(sheet, current_row, 10, passport)
                    safe_write(sheet, current_row, 11, dob_formatted)
                    safe_write(sheet, current_row, 21, remark)
                    
                else:
                    # ★ 邏輯 2 實現：如果房間只有 1 人，第 2 個迴圈會走到這裡。
                    # 程式會直接跳過不寫入任何資料，製造出完美的「空白列」。
                    pass
                
                # 處理完一個人(或一個空白列)，行數往下加一
                current_row += 1
                
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.balloons() 
        st.success(f"🎉 資料轉換成功！已依據房號進行智能排版，並生成檔名：{output_filename}")
        
        st.download_button(
            label=f"📥 下載最終名單 ({output_filename})",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"系統發出微弱的求救訊號：\n{e}")