import streamlit as st
import pandas as pd
import io
import datetime

# Setup Streamlit layout
st.set_page_config(page_title="Attendance Auto-Fill Portal", layout="wide")
st.title("📊 Attendance Auto-Fill Portal")
st.markdown("Excel တွင် 0 သို့မဟုတ် သတ်မှတ်နာရီမပြည့်သော နေရာများတွင် Google Sheet မှ Data များ အလိုအလျောက် ဖြည့်ပေးမည့် Portal")

st.info("""
💡 **Logic Rules:**
- **0 ဖြစ်နေပါက:** Google Sheet တွင် `D` သို့မဟုတ် `ORC` ဆိုပါက `SI` ဟု ဖြည့်မည်။ (`NF`, `L`, `HL` စသည်ဆိုပါက ၎င်းအတိုင်း ဖြည့်မည်)
- **5 အောက် ငယ်နေပါက:** Google Sheet တွင် `NF` ဆိုပါက `NF` ဟု ဖြည့်မည်။
- **8 အောက် ငယ်နေပါက:** Google Sheet တွင် `HL` ဆိုပါက `HL` ဟု ဖြည့်မည်။
""")

# 1. Hardcoded Google Sheet CSV Export URL
GS_CSV_URL = "https://docs.google.com/spreadsheets/d/1_G-YC1wUsE0_VMn2YDA9cnyspBqG_5HsFIc0T8g1mCs/export?format=csv&gid=1448030562"

# 2. Only Excel File Uploader
excel_file = st.file_uploader("Upload your Attendance Excel File (.xlsx)", type=["xlsx", "xls"])

def parse_date_string(date_val):
    try:
        return pd.to_datetime(date_val).strftime('%Y-%m-%d')
    except:
        try:
            parts = str(date_val).split('-')
            if len(parts) >= 3:
                clean_str = '-'.join(parts[:3])
                return pd.to_datetime(clean_str).strftime('%Y-%m-%d')
        except:
            pass
    return None

if excel_file is not None:
    try:
        with st.spinner("Fetching data from Google Sheet..."):
            df_excel = pd.read_excel(excel_file)
            
            # Clean columns and add a hidden space at the end to force Excel to treat it as Text
            cleaned_columns = []
            for col in df_excel.columns:
                if hasattr(col, 'strftime') or '00:00:00' in str(col):
                    clean_date = str(col).split(' ')[0].strip()
                    cleaned_columns.append(f"{clean_date} ") # Space appended
                else:
                    cleaned_columns.append(str(col).strip())
                    
            df_excel.columns = cleaned_columns
            df_gs = pd.read_csv(GS_CSV_URL)
            
        st.success("✅ Data fetched successfully! Processing matching rules...")
        
        # --- Processing Logic ---
        excel_id_col = df_excel.columns[0]   # Column A (ID)
        excel_name_col = df_excel.columns[1] # Column B (Name)
        
        gs_id_col = df_gs.columns[2]         # Column C (ID)
        gs_name_col = df_gs.columns[3]       # Column D (Name)
        
        all_cols = list(df_excel.columns)
        date_start_idx = 5
        date_end_idx = len(all_cols)
        if 'Incomplete Working Hours Days' in all_cols:
            date_end_idx = all_cols.index('Incomplete Working Hours Days')
            
        excel_date_cols = all_cols[date_start_idx:date_end_idx]
        gs_date_cols = df_gs.columns[5:]
        
        for col in excel_date_cols:
            df_excel[col] = df_excel[col].apply(
                lambda x: round(float(x), 1) if pd.notna(x) and isinstance(x, (int, float)) 
                else (round(float(x), 1) if isinstance(x, str) and x.replace('.', '', 1).isdigit() else x)
            )
            df_excel[col] = df_excel[col].astype(object)
        
        gs_dict_by_id = {}
        gs_dict_by_name = {}
        
        for _, row in df_gs.iterrows():
            emp_id = str(row[gs_id_col]).strip()
            name = str(row[gs_name_col]).strip().lower()
            
            if emp_id and emp_id != 'nan':
                gs_dict_by_id[emp_id] = row
            if name and name != 'nan':
                gs_dict_by_name[name] = row
                
        # 🟢 FIX 12: Changed 'si_records' to a general 'modified_records' list
        modified_records = []
        
        for idx, row in df_excel.iterrows():
            emp_id = str(row[excel_id_col]).strip()
            name = str(row[excel_name_col]).strip().lower()
            
            gs_row = None
            if emp_id in gs_dict_by_id:
                gs_row = gs_dict_by_id[emp_id]
            elif name in gs_dict_by_name:
                gs_row = gs_dict_by_name[name]
                
            if gs_row is not None:
                for ex_col in excel_date_cols:
                    ex_val = row[ex_col]
                    
                    val_f = None
                    if pd.isna(ex_val) or str(ex_val).strip() in ['0', '0.0', '0.00']:
                        val_f = 0.0
                    else:
                        try:
                            val_f = float(ex_val)
                        except (ValueError, TypeError):
                            val_f = None
                    
                    if val_f is not None and val_f < 8.0:
                        matched_gs_col = None
                        ex_date_str = parse_date_string(ex_col.strip())
                        
                        if ex_date_str:
                            for gs_col in gs_date_cols:
                                gs_date_str = parse_date_string(gs_col)
                                if gs_date_str and ex_date_str == gs_date_str:
                                    matched_gs_col = gs_col
                                    break
                        
                        if matched_gs_col:
                            gs_val = str(gs_row[matched_gs_col]).strip()
                            original_val_str = str(val_f) if val_f != 0.0 else "0"
                            
                            # Rule 1: If value is 0
                            if val_f == 0.0:
                                if gs_val in ['D', 'ORC']:
                                    df_excel.at[idx, ex_col] = 'SI'
                                    modified_records.append({
                                        'Name': row[excel_name_col], 
                                        'Date': ex_date_str, 
                                        'Action': f'Changed {gs_val} to SI'
                                    })
                                elif gs_val in ['NF', 'L', 'AB', 'UU', 'HL', 'ML', 'Resign']:
                                    df_excel.at[idx, ex_col] = gs_val
                                    modified_records.append({
                                        'Name': row[excel_name_col], 
                                        'Date': ex_date_str, 
                                        'Action': f'Filled with {gs_val}'
                                    })
                            
                            # Rule 2: If value is < 5 and GS is 'NF'
                            elif val_f > 0 and val_f < 5.0 and gs_val == 'NF':
                                df_excel.at[idx, ex_col] = 'NF'
                                modified_records.append({
                                    'Name': row[excel_name_col], 
                                    'Date': ex_date_str, 
                                    'Action': f'Changed {original_val_str} to NF'
                                })
                                
                            # Rule 3: If value is < 8 and GS is 'HL'
                            elif val_f >= 5.0 and val_f < 8.0 and gs_val == 'HL':
                                df_excel.at[idx, ex_col] = 'HL'
                                modified_records.append({
                                    'Name': row[excel_name_col], 
                                    'Date': ex_date_str, 
                                    'Action': f'Changed {original_val_str} to HL'
                                })

        # --- Count NF and L ---
        df_excel['NF Count'] = df_excel.apply(lambda row: (row == 'NF').sum(), axis=1)
        df_excel['L Count'] = df_excel.apply(lambda row: (row == 'L').sum(), axis=1)

        # --- Add Grand Total Row ---
        total_nf = df_excel['NF Count'].sum()
        total_l = df_excel['L Count'].sum()
        
        grand_total_row = {col: None for col in df_excel.columns}
        grand_total_row[excel_name_col] = 'Grand Total'
        grand_total_row['NF Count'] = total_nf
        grand_total_row['L Count'] = total_l
        
        df_excel = pd.concat([df_excel, pd.DataFrame([grand_total_row])], ignore_index=True)

        # --- Show Results ---
        st.write("### 🔍 Modified Records (SI, NF, HL အဖြစ် ပြောင်းလဲသွားသော မှတ်တမ်းများ)")
        if modified_records:
            st.dataframe(pd.DataFrame(modified_records), use_container_width=True)
        else:
            st.info("No records were modified.")
            
        st.write("### Preview of Updated Excel Data")
        st.dataframe(df_excel, use_container_width=True, height=600)
            
        # Download Button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_excel.to_excel(writer, index=False, sheet_name='Attendance')
        processed_data = output.getvalue()
        
        st.download_button(
            label="⬇️ Download Updated Excel File",
            data=processed_data,
            file_name="AutoFilled_Attendance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except pd.errors.ParserError:
        st.error("❌ Google Sheet ကို ဖတ်လို့မရပါ။ Google Sheet ကို 'Anyone with the link can view' ပြောင်းထားရန် လိုအပ်ပါသည်။")
    except Exception as e:
        st.error(f"Error occurred: {e}")
