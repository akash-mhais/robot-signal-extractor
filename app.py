import streamlit as st
import pandas as pd
import re
import io
import zipfile
import os
from openpyxl.styles import PatternFill

# Page Configuration
st.set_page_config(page_title="Robot Signal Extractor", page_icon="🤖", layout="wide")

# Custom CSS
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; font-weight: bold; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #28a745; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# Extraction & Signal Merging Engine
def parse_content(content):
    di_entries, do_entries = [], []
    di_pattern = re.compile(r'DI\[(\d+)(?::([^\]]*))?\](?:\s*(?:=|==|<>|:)\s*(ON|OFF))?', re.IGNORECASE)
    do_pattern = re.compile(r'DO\[(\d+)(?::([^\]]*))?\](?:\s*(?:=|==|<>|:)\s*(ON|OFF))?', re.IGNORECASE)
    for line in content.splitlines():
        for match in di_pattern.finditer(line):
            index = int(match.group(1))
            comment = (match.group(2) or "").strip()
            state = (match.group(3) or "").upper()
            di_entries.append((index, comment, state))
        for match in do_pattern.finditer(line):
            index = int(match.group(1))
            comment = (match.group(2) or "").strip()
            state = (match.group(3) or "").upper()
            do_entries.append((index, comment, state))
    return di_entries, do_entries

def update_signal_map(robot_data, robot_name, di_list, do_list):
    if robot_name not in robot_data:
        robot_data[robot_name] = {'DI': {}, 'DO': {}}
    for index, comment, state in di_list:
        if not comment and not state: continue
        curr_com, curr_st = robot_data[robot_name]['DI'].get(index, ("", ""))
        new_com = comment or curr_com
        new_st = curr_st
        prefix_match = re.match(r'^(ON|OFF)\s*:', new_com, re.IGNORECASE)
        if prefix_match:
            new_st = prefix_match.group(1).upper()
            new_com = re.sub(r'^(ON|OFF)\s*:\s*', '', new_com, flags=re.IGNORECASE)
        elif state and not new_st: new_st = state
        if not curr_com or new_st or len(new_com) > len(curr_com):
            robot_data[robot_name]['DI'][index] = (new_com, new_st)
    for index, comment, state in do_list:
        if not comment and not state: continue
        curr_com, curr_st = robot_data[robot_name]['DO'].get(index, ("", ""))
        new_com = comment or curr_com
        new_st = curr_st
        prefix_match = re.match(r'^(ON|OFF)\s*:', new_com, re.IGNORECASE)
        if prefix_match:
            new_st = prefix_match.group(1).upper()
            new_com = re.sub(r'^(ON|OFF)\s*:\s*', '', new_com, flags=re.IGNORECASE)
        elif state and not new_st: new_st = state
        if not curr_com or new_st or len(new_com) > len(curr_com):
            robot_data[robot_name]['DO'][index] = (new_com, new_st)

def generate_excel_bytes(robot_data):
    output = io.BytesIO()
    gf, rf = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid'), PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for robot_name, signals in robot_data.items():
            all_di, all_do = signals['DI'], signals['DO']
            has_state = any(s for _, s in all_di.values()) or any(s for _, s in all_do.values())
            di_rows = [[f"DI {i}", all_di[i][1], all_di[i][0]] if has_state else [f"DI {i}", all_di[i][0]] for i in sorted(all_di.keys())]
            do_rows = [[f"DO {i}", all_do[i][1], all_do[i][0]] if has_state else [f"DO {i}", all_do[i][0]] for i in sorted(all_do.keys())]
            if not di_rows and not do_rows: continue
            df_di = pd.DataFrame(di_rows, columns=['DI', 'State', 'Comment']) if has_state else pd.DataFrame(di_rows, columns=['DI', 'Comment'])
            df_do = pd.DataFrame(do_rows, columns=['DO', 'State', 'Comment']) if has_state else pd.DataFrame(do_rows, columns=['DO', 'Comment'])
            max_r = max(len(df_di), len(df_do))
            col_sp = pd.DataFrame([''] * max_r, columns=[' '])
            final_df = pd.concat([df_do.reindex(range(max_r)).fillna(''), col_sp, df_di.reindex(range(max_r)).fillna('')], axis=1)
            sheet_name = str(robot_name)[:31]
            final_df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            if has_state:
                ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width = 11, 8, 30
                ws.column_dimensions['D'].width, ws.column_dimensions['E'].width, ws.column_dimensions['F'].width, ws.column_dimensions['G'].width = 5, 11, 8, 30
                for r_idx in range(2, max_r + 2):
                    if ws.cell(row=r_idx, column=2).value == 'ON': ws.cell(row=r_idx, column=2).fill = gf
                    elif ws.cell(row=r_idx, column=2).value == 'OFF': ws.cell(row=r_idx, column=2).fill = rf
                    if ws.cell(row=r_idx, column=6).value == 'ON': ws.cell(row=r_idx, column=6).fill = gf
                    elif ws.cell(row=r_idx, column=6).value == 'OFF': ws.cell(row=r_idx, column=6).fill = rf
            else:
                ws.column_dimensions['A'].width, ws.column_dimensions['B'].width, ws.column_dimensions['C'].width, ws.column_dimensions['D'].width, ws.column_dimensions['E'].width = 13, 27, 13, 13, 21
    return output.getvalue()

def process_and_zip(all_robot_line_data):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
        for line_name, robot_data in all_robot_line_data.items():
            excel_bytes = generate_excel_bytes(robot_data)
            zip_file.writestr(f"{line_name}.xlsx", excel_bytes)
    return zip_buffer.getvalue()

def run_zip_processing(uploaded_file):
    line_data = {}
    with zipfile.ZipFile(uploaded_file, 'r') as z:
        for name in z.namelist():
            if name.lower().endswith('.ls'):
                parts = [p for p in name.split('/') if p]
                if 'LS' in [p.upper() for p in parts]:
                    ls_idx = [i for i, p in enumerate(parts) if p.upper() == 'LS'][0]
                    robot_name = parts[ls_idx-1] if ls_idx > 0 else "Robot"
                    line_name = parts[ls_idx-2] if ls_idx > 1 else "Main"
                else: 
                    robot_name = parts[-2] if len(parts) > 1 else "Root"
                    line_name = parts[-3] if len(parts) > 2 else "Main"
                if line_name not in line_data: line_data[line_name] = {}
                try:
                    with z.open(name) as f:
                        content = f.read().decode('latin-1', errors='replace')
                        di_l, do_l = parse_content(content)
                        update_signal_map(line_data[line_name], robot_name, di_l, do_l)
                except: pass
    return line_data

# UI Application - Simple Cloud Version
st.title("🤖 Robot Signal Extractor")
st.markdown("##### Professional Multi-Line Extraction Tool")

# Only Zip Upload for GitHub Version
st.info("💡 Please upload a **.ZIP** file of your robot backup or the entire robot line folder.")
up_f = st.file_uploader("Upload robot backup ZIP", type="zip")

if up_f:
    with st.spinner("Processing multiple robot lines..."):
        line_data = run_zip_processing(up_f)
        if line_data:
            st.success(f"Successfully processed {len(line_data)} lines!")
            for ln, robots in line_data.items(): st.write(f"- **{ln}**: {len(robots)} robots found")
            
            # Generate the ZIP content
            zip_bytes = process_and_zip(line_data)
            
            # Clean the filename for the zip download
            # Use a simple filename and unique key to ensure browser compatibility
            st.download_button(
                label="📥 Download Separate Excel Files (ZIP)",
                data=zip_bytes,
                file_name="Signal_Table.zip",
                mime="application/zip",
                key="download_zip_button"
            )
        else:
            st.error("No valid robot backups found in the ZIP.")

st.markdown("---")
st.caption("Powered by Advanced Agentic Automation")
