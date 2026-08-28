from base64 import b64encode

from datetime import datetime, timezone

from io import BytesIO

from pathlib import Path

import json


import pandas as pd

import streamlit as st

import streamlit.components.v1 as components

from supabase import create_client




# ---------------------------------------------------------

# PAGE CONFIGURATION

# ---------------------------------------------------------


st.set_page_config(

    page_title="Inventory Ageing Dashboard",

    page_icon="📊",

    layout="wide",

    initial_sidebar_state="collapsed",

)




# Reduce Streamlit spacing so only the dashboard is prominent.

st.markdown(

    """

    <style>

        .block-container {

            padding-top: 0.25rem;

            padding-left: 0.5rem;

            padding-right: 0.5rem;

            padding-bottom: 0;

            max-width: 100%;

        }


        [data-testid="stSidebar"] {

            min-width: 340px;

            max-width: 420px;

        }


        [data-testid="stSidebarContent"] {

            padding-top: 1rem;

        }


        iframe {

            width: 100%;

            border: none;

        }

    </style>

    """,

    unsafe_allow_html=True,

)




# ---------------------------------------------------------

# CONSTANTS

# ---------------------------------------------------------


BUCKET_NAME = "ageing-reports"

LATEST_FILE = "latest.xlsx"

METADATA_FILE = "latest_metadata.json"




# ---------------------------------------------------------

# SUPABASE CONNECTION

# ---------------------------------------------------------


@st.cache_resource

def get_supabase():

    return create_client(

        st.secrets["SUPABASE_URL"],

        st.secrets["SUPABASE_KEY"],

    )




# ---------------------------------------------------------

# EXCEL VALIDATION

# ---------------------------------------------------------


def read_excel_report(file_bytes):

    """

    Read the first non-empty worksheet and validate that it

    contains a material column and at least one ageing column.

    """


    workbook = pd.ExcelFile(

        BytesIO(file_bytes),

        engine="openpyxl",

    )


    if not workbook.sheet_names:

        raise ValueError(

            "The uploaded workbook contains no worksheets."

        )


    selected_dataframe = None

    selected_sheet = None


    for sheet_name in workbook.sheet_names:

        candidate = pd.read_excel(

            BytesIO(file_bytes),

            sheet_name=sheet_name,

            engine="openpyxl",

        )


        if not candidate.empty:

            selected_dataframe = candidate

            selected_sheet = sheet_name

            break


    if selected_dataframe is None:

        raise ValueError(

            "No data rows were found in the workbook."

        )


    selected_dataframe.columns = [

        str(column).strip()

        for column in selected_dataframe.columns

    ]


    normalized_columns = {

        str(column).strip().upper()

        for column in selected_dataframe.columns

    }


    material_headers = {

        "MATERIAL",

        "MATERIAL CODE",

    }


    if not normalized_columns.intersection(material_headers):

        raise ValueError(

            "Material or Material Code column was not found."

        )


    ageing_headers = []


    for column in selected_dataframe.columns:

        normalized_column = str(column).strip().upper()


        contains_amount_term = any(

            term in normalized_column

            for term in [

                "AMNT",

                "AMOUNT",

                "VALUE",

            ]

        )


        contains_number = any(

            character.isdigit()

            for character in normalized_column

        )


        if contains_amount_term and contains_number:

            ageing_headers.append(column)


    if not ageing_headers:

        raise ValueError(

            "No ageing amount columns were detected."

        )


    return (

        selected_dataframe,

        selected_sheet,

        ageing_headers,

    )




# ---------------------------------------------------------

# SUPABASE STORAGE FUNCTIONS

# ---------------------------------------------------------


def save_report_to_supabase(

    file_bytes,

    original_filename,

    row_count,

    worksheet,

    ageing_headers,

):

    """

    Save the latest Excel report and its metadata in Supabase.

    Existing files are replaced.

    """


    supabase = get_supabase()


    excel_content_type = (

        "application/vnd.openxmlformats-officedocument."

        "spreadsheetml.sheet"

    )


    supabase.storage.from_(BUCKET_NAME).upload(

        path=LATEST_FILE,

        file=file_bytes,

        file_options={

            "content-type": excel_content_type,

            "upsert": "true",

        },

    )


    metadata = {

        "fileName": original_filename,

        "updatedAt": datetime.now(

            timezone.utc

        ).isoformat(),

        "rowCount": int(row_count),

        "worksheet": worksheet,

        "ageingColumns": [

            str(column)

            for column in ageing_headers

        ],

    }


    metadata_bytes = json.dumps(

        metadata,

        ensure_ascii=False,

        indent=2,

    ).encode("utf-8")


    supabase.storage.from_(BUCKET_NAME).upload(

        path=METADATA_FILE,

        file=metadata_bytes,

        file_options={

            "content-type": "application/json",

            "upsert": "true",

        },

    )


    return metadata




def download_latest_report():

    """

    Download the latest shared Excel report.

    """


    try:

        supabase = get_supabase()


        return supabase.storage.from_(

            BUCKET_NAME

        ).download(LATEST_FILE)


    except Exception:

        return None




def download_latest_metadata():

    """

    Download metadata for the latest shared report.

    """


    try:

        supabase = get_supabase()


        metadata_bytes = supabase
 
