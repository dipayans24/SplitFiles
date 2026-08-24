import os
import random
import string
import tempfile
import zipfile
import streamlit as st 
import pandas as pd, re
from datetime import date, datetime
from math import ceil

st.markdown("""
    <style>
    div.stDownloadButton > button {
        background-color: #28a745;
        color: white;
        border: 1px solid #28a745;
    }
    div.stDownloadButton > button:hover {
        background-color: #218838;
        border: 1px solid #218838;
        color: white;
    }
    div.stDownloadButton > button:active {
        background-color: #1e7e34;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

@st.dialog("Error!!")
def raiseError(text):
    st.text(text)


def save_upload(uploaded_file):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, uploaded_file.name)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.read())
    return tmp_path


def identifier():
    return ''.join(random.choices(string.ascii_lowercase, k=4))


def exportFile(part_df, partName, CHUNK_SIZE):
    files = []
    UUID = ""#identifier()
    export_date = datetime.now().date().strftime("%d%m")
    for part_idx, start in enumerate(range(0, len(part_df), CHUNK_SIZE)):
        chunk = part_df.iloc[start:start + CHUNK_SIZE].copy()
        

        if partName is not None:
            mail_tag = f"{partName}{export_date}{part_idx + 1}".lower()
            chunk["mail to be sent"] = mail_tag
            
        else:
            mail_tag = f"file{export_date}{part_idx + 1}".lower()

        outputFileName = f"{mail_tag}.csv"
        chunk.to_csv(rf"{outputFileName}", index=False, sep=",")

        files.append(rf"{outputFileName}")

    return files 

def reset_columns():
    if "df_col" in st.session_state:
        st.session_state.df_col = None

def update_colums():
    if "df_col" in st.session_state:
        st.session_state.df_col = None

    if "sheet_num" in st.session_state:
        st.session_state.sheet_num = None

def splitFiles(filePath,fileextn, select_columns, export_columns, CHUNK_SIZE, manual = False):
    if fileextn == "xlsx":
        df = pd.read_excel(filePath)
    else:
        df= pd.read_csv(filePath,sep=",", encoding_errors="ignore", low_memory=False)

    st.write("File read complete.")
    PATTERN = re.compile(r"\W+")
    
    ziptemp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    ziptemp.close()

    with zipfile.ZipFile(ziptemp.name, "w", zipfile.ZIP_DEFLATED) as zf:
        if manual == False:
            Values = df[select_columns].value_counts().to_frame("index").reset_index(drop=False)
            Values["Parts"] = Values["index"].map(lambda x: ceil(x/CHUNK_SIZE))

        
            for funnel, _ in Values[[select_columns, "Parts"]].itertuples(index=False):
                funnel = str(funnel)
                if select_columns != "index":
                    part_df = (df.loc[df[select_columns] == funnel, export_columns].reset_index(drop=True))
                else:
                    part_df = (df.loc[:,  export_columns].reset_index(drop=True))

                clean_funnel = PATTERN.sub("", funnel)

                partName = f"{clean_funnel}"
                files = exportFile(part_df.iloc[1:, :], partName, CHUNK_SIZE)
                
                st.write("Packing in zip...")
                for file in files:
                    zf.write(file, arcname=file)
                    os.unlink(file)

        else:
            Values = df.iloc[1:, :][export_columns].reset_index(drop=True)
            files = exportFile(Values, select_columns, CHUNK_SIZE)

            st.write("Packing in zip...")
            for file in files:
                zf.write(file, arcname=file)
                os.unlink(file)
        

    return ziptemp.name

 
st.set_page_config(page_title="Split Files", page_icon="✂️", layout="wide")
st.title("✂️ Split Files")

fileLoc = st.file_uploader("Upload File", type=["csv", "xlsx"], on_change=update_colums)

if "df_col" not in st.session_state:
    st.session_state.df_col = None

if "sheet_num" not in st.session_state:
    st.session_state.sheet_num = None
    
if fileLoc:
    filePath = save_upload(fileLoc)
    fileName,  fileextn = os.path.basename(filePath).split(".")
    CHUNK_SIZE = int(st.number_input("Enter the number of rows", step=1.0))

    cleanPattern = re.compile(r"\s")
    
    if fileextn == "xlsx":
            if st.session_state.sheet_num is None:
                with pd.ExcelFile(filePath, engine="openpyxl") as file:
                    st.session_state.sheet_num = file.sheet_names
                
            if len(st.session_state.sheet_num) > 1:
                select_sheets = st.selectbox(label = "Multiple Sheets detected. Select a sheet to process", options=st.session_state.sheet_num, on_change=reset_columns)
                if select_sheets and st.session_state.df_col is None:
                    df = pd.read_excel(filePath,nrows=10 , sheet_name=select_sheets)
                    st.session_state.df_col = df.columns
            else:
                if st.session_state.df_col is None:
                    df = pd.read_excel(filePath,nrows=10 , sheet_name=st.session_state.sheet_num[0])
            
                    st.session_state.df_col = df.columns
    else:
            if st.session_state.df_col is None:
                df= pd.read_csv(filePath,sep=",", nrows = 10, encoding_errors="replace")

                st.session_state.df_col = df.columns

    chkbox = st.checkbox(label="Check to manually define 'mail to be sent' info")

    manual = False
    if chkbox:
        select_columns = st.text_input("Enter a custom identifier for the file name", max_chars=20, placeholder="Leave the field empty if 'mail to be sent' is not required")
        manual = True
    else:
        select_columns = st.selectbox(label = "Select a column to process for 'mail to be sent'", options=st.session_state.df_col, index=None)
 

     
    export_columns = st.multiselect(label = "Select columns to export.", options=st.session_state.df_col)
    
    if export_columns:
        btn = st.button(label="Generate Data", type="primary")
        
        if btn:
            with st.spinner("", show_time=True):
                with st.status("Processing...", expanded=True) as status:
                    if len(cleanPattern.sub("", select_columns)) == 0:
                        select_columns = None
                    zipFileName = splitFiles(filePath,fileextn, select_columns, export_columns, CHUNK_SIZE, manual)

                    status.update(state="complete", expanded=False )
                    outputFileName = f"{fileName}_frmt.zip"

                with open(zipFileName, "rb") as f:
                    st.download_button("Download Files", f,  file_name=outputFileName, on_click = "ignore")

                os.unlink(zipFileName) 
                

    
