import streamlit as st

from supabase import create_client




st.set_page_config(

    page_title="Inventory Ageing Dashboard",

    layout="wide",

)


st.title("Inventory Ageing Dashboard")


try:

    supabase = create_client(

        st.secrets["SUPABASE_URL"],

        st.secrets["SUPABASE_KEY"],

    )


    buckets = supabase.storage.list_buckets()

    bucket_names = [bucket.name for bucket in buckets]


    if "ageing-reports" in bucket_names:

        st.success(

            "Supabase is connected and the ageing-reports "

            "bucket is available."

        )

    else:

        st.warning(

            "Supabase is connected, but the ageing-reports "

            "bucket was not found."

        )


except Exception as error:

    st.error(f"Supabase connection failed: {error}")
