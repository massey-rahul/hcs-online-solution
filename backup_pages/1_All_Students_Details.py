import os
import streamlit as st
import pandas as pd

st.set_page_config(page_title="All Students Details - Holy Child School", layout="wide")

st.markdown("<h2 style='text-align: center; color: #1e3d59;'>📚 All Registered Students Details Hub</h2>", unsafe_allow_html=True)
st.write("Yahan aap Excel file ke sabhi registered students ki complete details dekh aur search kar sakte hain.")
st.write("---")

excel_file = "Online_Registration.xlsx"
if not os.path.exists(excel_file):
    excel_file = "Online_Registration (2).xls"

if os.path.exists(excel_file):
    try:
        df = pd.read_excel(excel_file, header=None, engine='openpyxl')
    except Exception:
        try:
            df = pd.read_excel(excel_file, header=None, engine='xlrd')
        except Exception:
            df = None

    if df is not None:
        # Serial numbers start from 1
        df_display = df.iloc[1:].copy()
        df_display.index = range(1, len(df_display) + 1)
        
        search_term = st.text_input("🔍 Search Student (by Name, Reg No, or Class):")
        
        if search_term:
            filtered_rows = []
            for idx, row in df_display.iterrows():
                row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                if search_term.lower() in row_str.lower():
                    filtered_rows.append(row)
            if filtered_rows:
                df_filtered = pd.DataFrame(filtered_rows)
                df_filtered.index = range(1, len(df_filtered) + 1)
                st.success(f"Kul {len(df_filtered)} matching records mile hain:")
                st.dataframe(df_filtered, use_container_width=True)
            else:
                st.warning("⚠️ Koi matching student nahi mila.")
        else:
            st.info(f"Total Registered Students: {len(df_display)}")
            st.dataframe(df_display, use_container_width=True)
    else:
        st.error("⚠️ Excel file read karne mein error aayi.")
else:
    st.warning(f"⚠️ Excel file ('Online_Registration.xlsx') project folder mein nahi mili.")
