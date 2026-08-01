import os
import random
import string
import tempfile
import zipfile
import streamlit as st 
import pandas as pd, re
from datetime import date, datetime
from math import ceil
 

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
        mail_tag = f"{partName}{export_date}{part_idx + 1}".lower()

        chunk["mail to be sent"] = mail_tag
        outputFileName = f"{mail_tag}.csv"
        chunk.to_csv(rf"{outputFileName}", index=False, sep=",")

        files.append(rf"{outputFileName}")

    return files 


def splitFiles(filePath,fileextn, select_columns, export_columns, CHUNK_SIZE, manual = False):
    if fileextn == "xlsx":
        df = pd.read_excel(filePath)
    else:
        df= pd.read_csv(filePath,sep=",", encoding_errors="ignore", low_memory=False)

    PATTERN = re.compile(r"\W+")
    
    ziptemp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    ziptemp.close()

    with zipfile.ZipFile(ziptemp.name, "w", zipfile.ZIP_DEFLATED) as zf:
        if manual == False :
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

                for file in files:
                    zf.write(file, arcname=file)
                    os.unlink(file)
        else:
            Values = df.iloc[1:, :][export_columns].reset_index(drop=True)
            files = exportFile(Values, select_columns, CHUNK_SIZE)

            for file in files:
                zf.write(file, arcname=file)
                os.unlink(file)

    return ziptemp.name

st.header("Split Files")

fileLoc = st.file_uploader("Upload File", type=["csv", "xlsx"])

if "df_col" not in st.session_state:
    st.session_state.df_col = None

if fileLoc:
    filePath = save_upload(fileLoc)
    fileName,  fileextn = os.path.basename(filePath).split(".")
    CHUNK_SIZE = int(st.number_input("Enter the number of rows", step=1.0))

    cleanPattern = re.compile(r"\s")
    
    if st.session_state.df_col is None:
        if fileextn == "xlsx":
            with pd.ExcelFile(filePath, engine="openpyxl") as file:
                sheets = file.sheet_names
                
            if len(sheets) > 1:
                select_sheets = st.selectbox(label = "Multiple Sheets detected. Select a sheet to process", options=sheets)
                if select_sheets:
                    df = pd.read_excel(filePath,nrows=100 , sheet_name=select_sheets)
            else:
                df = pd.read_excel(filePath,nrows=100 , sheet_name=sheets[0])
        else:
            df= pd.read_csv(filePath,sep=",", nrows = 100, encoding_errors="replace")

        st.session_state.df_col = df.columns

    chkbox = st.checkbox(label="Check to manually define 'mail to be sent' info")

    manual = False
    if chkbox:
        select_columns = st.text_input("Enter a custom identifier for the file name", max_chars=20)
        manual = True
    else:
        select_columns = st.selectbox(label = "Select a column to process for 'mail to be sent'", options=st.session_state.df_col, index=None)
 

    if select_columns:
        export_columns = st.multiselect(label = "Select columns to export.", options=st.session_state.df_col)
        
        if export_columns:
            btn = st.button(label="Generate Data", type="primary")
            
            if btn:
                if len(cleanPattern.sub("", select_columns)) > 0:
                    with st.spinner("Processing..",show_time=True):
                        zipFileName = splitFiles(filePath,fileextn, select_columns, export_columns, CHUNK_SIZE, manual)

                        outputFileName = f"{fileName}_frmt.zip"
                        with open(zipFileName, "rb") as f:
                                    st.download_button("Download Files", f,  file_name=outputFileName)

                        os.unlink(zipFileName) 
                else:
                    raiseError("Custom identifier cannot be blank or whitespace.")

    
