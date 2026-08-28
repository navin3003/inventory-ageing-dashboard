from base64 import b64encode

from datetime import datetime, timezone

from pathlib import Path

from io import BytesIO

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




def save_report_to_supabase(

    file_bytes,

    original_filename,

    row_count,

    worksheet,

    ageing_headers,

):

    supabase = get_supabase()


    content_type = (

        "application/vnd.openxmlformats-officedocument."

        "spreadsheetml.sheet"

    )


    supabase.storage.from_(BUCKET_NAME).upload(

        path=LATEST_FILE,

        file=file_bytes,

        file_options={

            "content-type": content_type,

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

    supabase = get_supabase()


    try:

        return supabase.storage.from_(

            BUCKET_NAME

        ).download(LATEST_FILE)

    except Exception:

        return None




def download_latest_metadata():

    supabase = get_supabase()


    try:

        metadata_bytes = supabase.storage.from_(

            BUCKET_NAME

        ).download(METADATA_FILE)


        return json.loads(

            metadata_bytes.decode("utf-8")

        )

    except Exception:

        return {}




def inject_shared_excel(

    dashboard_html,

    excel_bytes,

    metadata,

):

    encoded_excel = b64encode(

        excel_bytes

    ).decode("ascii")


    filename = metadata.get(

        "fileName",

        "Shared ageing report.xlsx",

    )


    filename_json = json.dumps(filename)


    shared_data_script = f"""

<script>

(function loadSharedAgeingReport() {{

    const encodedExcel = "{encoded_excel}";

    const sharedFilename = {filename_json};


    function decodeBase64(base64Value) {{

        const binaryValue = atob(base64Value);

        const bytes = new Uint8Array(binaryValue.length);


        for (let index = 0;

             index < binaryValue.length;

             index += 1) {{

            bytes[index] = binaryValue.charCodeAt(index);

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

            typeof render !== "function"

        ) {{

            setTimeout(startSharedReportLoad, 250);

            return;

        }}


        try {{

            const excelBytes = decodeBase64(encodedExcel);


            const workbook = XLSX.read(

                excelBytes,

                {{

                    type: "array",

                    cellDates: true

                }}

            );


            let rawRows = [];

            let selectedSheet = "";


            for (const sheetName of workbook.SheetNames) {{

                const candidateRows =

                    XLSX.utils.sheet_to_json(

                        workbook.Sheets[sheetName],

                        {{

                            defval: "",

                            raw: true

                        }}

                    );


                if (candidateRows.length) {{

                    rawRows = candidateRows;

                    selectedSheet = sheetName;

                    break;

                }}

            }}


            if (!rawRows.length) {{

                throw new Error(

                    "The shared workbook has no data rows."

                );

            }}


            const detectedBuckets = detect(

                Object.keys(rawRows[0] || {{}})

            );


            setB(detectedBuckets);


            const parsedRows = rawRows

                .map(norm)

                .filter(

                    row =>

                        row.MATERIAL ||

                        row.DESCRIPTION

                );


            DATA.splice(

                0,

                DATA.length,

                ...parsedRows

            );


            refresh();


            [

                "plant",

                "type",

                "group",

                "storage"

            ].forEach(elementId => {{

                const element =

                    document.getElementById(elementId);


                if (element) {{

                    element.value = "";

                }}

            }});


            const searchElement =

                document.getElementById("search");


            if (searchElement) {{

                searchElement.value = "";

            }}


            render();


            const sourceElement =

                document.getElementById("source");


            if (sourceElement) {{

                sourceElement.textContent =

                    "Shared report: " +

                    sharedFilename +

                    " | " +

                    parsedRows.length.toLocaleString(

                        "en-IN"

                    ) +

                    " rows";

            }}


            const originalUpload =

                document.querySelector(".upload");


            if (originalUpload) {{

                originalUpload.style.display = "none";

            }}


            console.log(

                "Shared report loaded:",

                sharedFilename,

                selectedSheet,

                parsedRows.length

            );


        }} catch (error) {{

            console.error(

                "Unable to load shared ageing report:",

                error

            );


            const sourceElement =

                document.getElementById("source");


            if (sourceElement) {{

                sourceElement.textContent =

                    "Unable to load shared report: " +

                    error.message;

            }}

        }}

    }}


    if (document.readyState === "loading") {{

        document.addEventListener(

            "DOMContentLoaded",

            startSharedReportLoad

        );

    }} else {{

        startSharedReportLoad();

    }}

}})();

</script>

"""


    return dashboard_html.replace(

        "</body>",

        shared_data_script + "</body>",

    )




st.title("Inventory Ageing Dashboard")


with st.expander(

    "Publish latest ageing report",

    expanded=False,

):

    st.warning(

        "Publishing a report replaces the shared report "

        "for everyone using this dashboard."

    )


    publishing_password = st.text_input(

        "Publishing password",

        type="password",

        key="publishing_password",

    )


    uploaded_file = st.file_uploader(

        "Select the latest ageing Excel report",

        type=["xlsx"],

        key="ageing_report_file",

    )


    publish_button = st.button(

        "Publish report for everyone",

        type="primary",

        use_container_width=True,

    )


    if publish_button:

        if (

            publishing_password

            != st.secrets["UPLOAD_PASSWORD"]

        ):

            st.error(

                "Incorrect publishing password. "

                "The shared report was not changed."

            )


        elif uploaded_file is None:

            st.error(

                "Select an Excel report before publishing."

            )


        else:

            try:

                uploaded_bytes = uploaded_file.getvalue()


                (

                    uploaded_dataframe,

                    selected_worksheet,

                    detected_ageing_headers,

                ) = read_excel_report(uploaded_bytes)


                saved_metadata = save_report_to_supabase(

                    file_bytes=uploaded_bytes,

                    original_filename=uploaded_file.name,

                    row_count=len(uploaded_dataframe),

                    worksheet=selected_worksheet,

                    ageing_headers=detected_ageing_headers,

                )


                st.success(

                    f"{uploaded_file.name} was published "

                    f"successfully with "

                    f"{len(uploaded_dataframe):,} rows. "

                    f"The report is now available to all viewers."

                )


                st.rerun()


            except Exception as error:

                st.error(

                    f"Publishing failed: {error}"

                )




latest_report_bytes = download_latest_report()

latest_metadata = download_latest_metadata()


if latest_report_bytes:

    try:

        shared_dataframe, shared_sheet, _ = (

            read_excel_report(latest_report_bytes)

        )


        filename = latest_metadata.get(

            "fileName",

            "latest.xlsx",

        )


        updated_at = latest_metadata.get(

            "updatedAt",

            "",

        )


        st.success(

            f"Shared report: {filename} | "

            f"{len(shared_dataframe):,} rows"

        )


        if updated_at:

            st.caption(

                f"Last published: {updated_at}"

            )


    except Exception as error:

        st.error(

            f"The shared report could not be validated: "

            f"{error}"

        )

else:

    st.info(

        "No shared report has been published yet. "

        "Open the publishing section above to publish "

        "the first report."

    )




dashboard_path = (

    Path(__file__).parent / "dashboard.html"

)


if not dashboard_path.exists():

    st.error(

        "dashboard.html was not found in the repository."

    )

    st.stop()


dashboard_html = dashboard_path.read_text(

    encoding="utf-8"

)


if latest_report_bytes:

    dashboard_html = inject_shared_excel(

        dashboard_html=dashboard_html,

        excel_bytes=latest_report_bytes,

        metadata=latest_metadata,

    )


components.html(

    dashboard_html,

    height=2400,

    scrolling=True,

)
