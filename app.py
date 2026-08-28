from base64 import b64encode

from datetime import datetime, timezone

from io import BytesIO

from pathlib import Path

import json
 
import pandas as pd

import streamlit as st

import streamlit.components.v1 as components

from supabase import create_client
 
st.set_page_config(

    page_title="Inventory Ageing Dashboard",

    page_icon="📊",

    layout="wide",

    initial_sidebar_state="collapsed",

)
 
st.markdown(

    """
<style>

    .block-container {

        padding: 0.25rem 0.5rem 0 0.5rem;

        max-width: 100%;

    }

    [data-testid="stSidebar"] {

        min-width: 340px;

        max-width: 420px;

    }

    [data-testid="stSidebarContent"] { padding-top: 1rem; }

    iframe { width: 100%; border: none; }
</style>

    """,

    unsafe_allow_html=True,

)
 
BUCKET_NAME = "ageing-reports"

LATEST_FILE = "latest.xlsx"

METADATA_FILE = "latest_metadata.json"
 
 
@st.cache_resource

def get_supabase():

    return create_client(

        st.secrets["SUPABASE_URL"],

        st.secrets["SUPABASE_KEY"],

    )
 
 
def read_excel_report(file_bytes):

    workbook = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")

    if not workbook.sheet_names:

        raise ValueError("The workbook contains no worksheets.")
 
    selected_dataframe = None

    selected_sheet = None

    for sheet_name in workbook.sheet_names:

        candidate = pd.read_excel(

            BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl"

        )

        if not candidate.empty:

            selected_dataframe = candidate

            selected_sheet = sheet_name

            break
 
    if selected_dataframe is None:

        raise ValueError("No data rows were found in the workbook.")
 
    selected_dataframe.columns = [

        str(column).strip() for column in selected_dataframe.columns

    ]

    normalized_columns = {

        str(column).strip().upper() for column in selected_dataframe.columns

    }
 
    if not normalized_columns.intersection({"MATERIAL", "MATERIAL CODE"}):

        raise ValueError("Material or Material Code column was not found.")
 
    ageing_headers = []

    for column in selected_dataframe.columns:

        normalized = str(column).strip().upper()

        has_amount_term = any(

            term in normalized for term in ("AMNT", "AMOUNT", "VALUE")

        )

        has_number = any(character.isdigit() for character in normalized)

        if has_amount_term and has_number:

            ageing_headers.append(column)
 
    if not ageing_headers:

        raise ValueError("No ageing amount columns were detected.")
 
    return selected_dataframe, selected_sheet, ageing_headers
 
 
def save_report_to_supabase(

    file_bytes, original_filename, row_count, worksheet, ageing_headers

):

    supabase = get_supabase()

    excel_mime = (

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )
 
    supabase.storage.from_(BUCKET_NAME).upload(

        path=LATEST_FILE,

        file=file_bytes,

        file_options={"content-type": excel_mime, "upsert": "true"},

    )
 
    metadata = {

        "fileName": original_filename,

        "updatedAt": datetime.now(timezone.utc).isoformat(),

        "rowCount": int(row_count),

        "worksheet": worksheet,

        "ageingColumns": [str(column) for column in ageing_headers],

    }

    metadata_bytes = json.dumps(

        metadata, ensure_ascii=False, indent=2

    ).encode("utf-8")
 
    supabase.storage.from_(BUCKET_NAME).upload(

        path=METADATA_FILE,

        file=metadata_bytes,

        file_options={"content-type": "application/json", "upsert": "true"},

    )

    return metadata
 
 
def download_latest_report():

    try:

        return get_supabase().storage.from_(BUCKET_NAME).download(LATEST_FILE)

    except Exception:

        return None
 
 
def download_latest_metadata():

    try:

        metadata_bytes = (

            get_supabase().storage.from_(BUCKET_NAME).download(METADATA_FILE)

        )

        return json.loads(metadata_bytes.decode("utf-8"))

    except Exception:

        return {}
 
 
def inject_shared_excel(dashboard_html, excel_bytes, metadata):

    encoded_excel = b64encode(excel_bytes).decode("ascii")

    filename_json = json.dumps(

        metadata.get("fileName", "Shared ageing report.xlsx")

    )
 
    shared_data_script = f"""
<script>

(function () {{

    const encodedExcel = "{encoded_excel}";

    const sharedFilename = {filename_json};
 
    function decodeBase64(value) {{

        const binary = atob(value);

        const bytes = new Uint8Array(binary.length);

        for (let i = 0; i < binary.length; i += 1) {{

            bytes[i] = binary.charCodeAt(i);

        }}

        return bytes;

    }}
 
    function startSharedReportLoad() {{

        if (typeof XLSX === "undefined") {{

            setTimeout(startSharedReportLoad, 250);

            return;

        }}

        if (

            typeof detect !== "function" ||

            typeof setB !== "function" ||

            typeof norm !== "function" ||

            typeof refresh !== "function" ||

            typeof render !== "function"

        ) {{

            setTimeout(startSharedReportLoad, 250);

            return;

        }}
 
        try {{

            const workbook = XLSX.read(decodeBase64(encodedExcel), {{

                type: "array",

                cellDates: true

            }});
 
            let rawRows = [];

            for (const sheetName of workbook.SheetNames) {{

                const candidateRows = XLSX.utils.sheet_to_json(

                    workbook.Sheets[sheetName],

                    {{ defval: "", raw: true }}

                );

                if (candidateRows.length) {{

                    rawRows = candidateRows;

                    break;

                }}

            }}
 
            if (!rawRows.length) {{

                throw new Error("The shared workbook has no data rows.");

            }}
 
            setB(detect(Object.keys(rawRows[0] || {{}})));

            const parsedRows = rawRows

                .map(norm)

                .filter(row => row.MATERIAL || row.DESCRIPTION);
 
            if (!parsedRows.length) {{

                throw new Error("No usable material rows were found.");

            }}
 
            DATA.splice(0, DATA.length, ...parsedRows);

            refresh();
 
            ["plant", "type", "group", "storage"].forEach(id => {{

                const element = document.getElementById(id);

                if (element) element.value = "";

            }});

            const search = document.getElementById("search");

            if (search) search.value = "";
 
            render();
 
            const source = document.getElementById("source");

            if (source) {{

                source.textContent =

                    "Shared report: " + sharedFilename + " | " +

                    parsedRows.length.toLocaleString("en-IN") + " rows";

            }}
 
            const oldUpload = document.querySelector(".upload");

            if (oldUpload) oldUpload.style.display = "none";

        }} catch (error) {{

            console.error("Unable to load shared ageing report:", error);

            const source = document.getElementById("source");

            if (source) {{

                source.textContent = "Unable to load shared report: " + error.message;

            }}

        }}

    }}
 
    if (document.readyState === "loading") {{

        document.addEventListener("DOMContentLoaded", startSharedReportLoad);

    }} else {{

        startSharedReportLoad();

    }}

}})();
</script>

"""
 
    if "</body>" in dashboard_html:

        return dashboard_html.replace(

            "</body>", shared_data_script + "</body>", 1

        )

    return dashboard_html + shared_data_script
 
 
# Load the current shared report before building the sidebar.

latest_report_bytes = download_latest_report()

latest_metadata = download_latest_metadata()
 
with st.sidebar:

    st.header("Dashboard controls")
 
    with st.expander("Publish latest ageing report", expanded=False):

        st.warning(

            "Publishing replaces the shared report for everyone using this dashboard."

        )

        publishing_password = st.text_input(

            "Publishing password", type="password", key="publishing_password"

        )

        uploaded_file = st.file_uploader(

            "Select the latest ageing Excel report",

            type=["xlsx"],

            key="ageing_report_file",

        )
 
        if st.button(

            "Publish report for everyone",

            type="primary",

            use_container_width=True,

        ):

            if publishing_password != st.secrets["UPLOAD_PASSWORD"]:

                st.error(

                    "Incorrect publishing password. The shared report was not changed."

                )

            elif uploaded_file is None:

                st.error("Select an Excel report before publishing.")

            else:

                try:

                    uploaded_bytes = uploaded_file.getvalue()

                    dataframe, worksheet, ageing_headers = read_excel_report(

                        uploaded_bytes

                    )

                    save_report_to_supabase(

                        file_bytes=uploaded_bytes,

                        original_filename=uploaded_file.name,

                        row_count=len(dataframe),

                        worksheet=worksheet,

                        ageing_headers=ageing_headers,

                    )

                    st.success(

                        f"{uploaded_file.name} was published successfully "

                        f"with {len(dataframe):,} rows."

                    )

                    st.rerun()

                except Exception as error:

                    st.error(f"Publishing failed: {error}")
 
    st.divider()

    st.subheader("Current shared report")
 
    if latest_report_bytes:

        st.write(f"**{latest_metadata.get('fileName', 'latest.xlsx')}**")

        row_count = latest_metadata.get("rowCount")

        if row_count is not None:

            st.caption(f"Rows: {int(row_count):,}")
 
        updated_at = latest_metadata.get("updatedAt", "")

        if updated_at:

            try:

                formatted = datetime.fromisoformat(

                    updated_at.replace("Z", "+00:00")

                ).strftime("%d-%m-%Y %H:%M UTC")

                st.caption(f"Last published: {formatted}")

            except ValueError:

                st.caption(f"Last published: {updated_at}")

    else:

        st.info("No shared report has been published yet.")
 
 
# Validate the current shared report without adding a large status box.

if latest_report_bytes:

    try:

        read_excel_report(latest_report_bytes)

    except Exception as error:

        st.error(f"The shared report could not be validated: {error}")
 
 
dashboard_path = Path(__file__).parent / "dashboard.html"

if not dashboard_path.exists():

    st.error("dashboard.html was not found in the repository.")

    st.stop()
 
dashboard_html = dashboard_path.read_text(encoding="utf-8")

if latest_report_bytes:

    dashboard_html = inject_shared_excel(

        dashboard_html=dashboard_html,

        excel_bytes=latest_report_bytes,

        metadata=latest_metadata,

    )
 
components.html(dashboard_html, height=2400, scrolling=True)

 
