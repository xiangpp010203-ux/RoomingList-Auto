import streamlit as st
import pandas as pd
import openpyxl
import copy
from io import BytesIO
import os

st.set_page_config(page_title="OP 救星：分房表轉換神器", page_icon="🏨")
st.title("🏨 OP 救星：SCM_Rooming List 自動轉換神器")
st.write("支援單人房留空、三人房動態新增列與 L~V 欄自動合併格式化！")

# ==========================================
# ★ 輔助函式區
# ==========================================
def safe_write(sheet, r, c, val):
    """安全寫入：遇到合併儲存格自動閃避"""
    cell = sheet.cell(row=r, column=c)
    if type(cell).__name__ != 'MergedCell':
        cell.value = val

def copy_style(source_cell, target_cell):
    """格式複製：完美複製框線、字體、背景色"""
    if source_cell.has_style:
        target_cell.font = copy.copy(source_cell.font)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.number_format = copy.copy(source_cell.number_format)
        target_cell.alignment = copy.copy(source_cell.alignment)

def clean_str(val):
    """資料清洗：徹底消滅 nan 與隱形空白鍵"""
    if pd.isna(val):
        return ""
    s = str(val).strip() # 徹底清除前後空白 (含半形與全形)
    if s.lower() == 'nan' or s == '':
        return ""
    return s

def remerge_room_columns(sheet, start_row, end_row):
    """★ 最新優化：將指定房間區塊的 A欄與 L~V欄 進行垂直合併與置中"""
    if start_row >= end_row:
        return
        
    # 需要合併的欄位：A欄(房號=1), L~V欄(入住資訊=12~22)
    cols_to_merge = [1] + list(range(12, 23))
    
    for col in cols_to_merge:
        # 1. 先解除這個房間範圍內原有的合併 (避免舊格式衝突)
        for merged_range in list(sheet.merged_cells.ranges):
            if merged_range.min_col == col and merged_range.max_col == col:
                if merged_range.min_row >= start_row and merged_range.max_row <= end_row:
                    sheet.unmerge_cells(str(merged_range))
        
        # 2. 重新合併整個房間的列數
        sheet.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
        
        # 3. 設定視覺效果：垂直/水平置中對齊
        top_cell = sheet.cell(row=start_row, column=col)
        top_cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

# ==========================================
# 主程式區塊
# ==========================================
uploaded_file_B = st.file_uploader("請上傳【內部系統分房表 (檔案B)】支援 Excel 或 CSV", type=["xlsx", "csv"])

if uploaded_file_B is not None:
    st.success(f"成功讀取檔案：{uploaded_file_B.name}！正在進行智能分房排版與 L~V 欄格式重組...")
    
    try:
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
        
        wb = openpyxl.load_workbook("Template.xlsx")
        sheet = wb.active 
        
        start_row = 14 
        current_row = start_row
        
        df_B_valid = df_B.dropna(subset=['房號'])
        
        room_numbers = []
        for r in df_B_valid['房號']:
            if pd.notna(r) and r not in room_numbers:
                room_numbers.append(r)
                
        for room_no in room_numbers:
            room_group = df_B_valid[df_B_valid['房號'] == room_no]
            passengers = room_group.to_dict('records')
            num_people = len(passengers)
            
            rows_to_occupy = max(2, num_people)
            start_room_row = current_row  # 紀錄這間房間的起始列
            
            for i in range(rows_to_occupy):
                # 動態新增列並「完美複製」上一列的格式 (框線不跑位)
                if i >= 2:
                    sheet.insert_rows(current_row)
                    for col in range(1, sheet.max_column + 1):
                        source_c = sheet.cell(row=current_row - 1, column=col)
                        target_c = sheet.cell(row=current_row, column=col)
                        copy_style(source_c, target_c)
                
                if i < num_people:
                    row = passengers[i]
                    
                    eng_name_raw = clean_str(row.get('英文姓名'))
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

                    dob_raw = clean_str(row.get('生日'))
                    if dob_raw.endswith('.0'): dob_raw = dob_raw[:-2]
                    if len(dob_raw) == 8 and dob_raw.isdigit():
                        dob_formatted = f"{dob_raw[0:4]}/{dob_raw[4:6]}/{dob_raw[6:8]}"
                    else:
                        dob_formatted = dob_raw
                        
                    cht_name = clean_str(row.get('中文姓名'))
                    cht_last, cht_first = "", ""
                    if len(cht_name) >= 2:
                        cht_last = cht_name[0]      
                        cht_first = cht_name[1:]    
                    elif len(cht_name) == 1:
                        cht_last = cht_name[0]
                        
                    room_raw = clean_str(room_no)
                    if room_raw.endswith('.0'): room_raw = room_raw[:-2]
                    room_val = room_raw if i == 0 else ""
                        
                    no_raw = clean_str(row.get('No'))
                    if no_raw.endswith('.0'): no_raw = no_raw[:-2]
                    no_val = no_raw
                        
                    passport = clean_str(row.get('護照號碼'))
                    if passport.endswith('.0'): passport = passport[:-2]
                    
                    remark = clean_str(row.get('備註'))

                    # 寫入旅客資料
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
                    # 空白佔位列，不寫入任何資料
                    pass
                
                current_row += 1
                
            end_room_row = current_row - 1 # 紀錄這間房間的結束列
            
            # ★ 最新優化：這間房的資料填完後，將 L~V 欄與 A 欄依照人數進行垂直合併
            remerge_room_columns(sheet, start_room_row, end_room_row)
                
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.balloons() 
        st.success(f"🎉 格式重組完成！已為 L~V 欄建立完美的合併區塊。檔名：{output_filename}")
        
        st.download_button(
            label=f"📥 下載最終名單 ({output_filename})",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"系統發出微弱的求救訊號：\n{e}")