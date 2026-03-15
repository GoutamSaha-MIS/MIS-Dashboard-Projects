import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="MIS Dashboard", layout="wide")

st.title("MIS Performance Dashboard")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file:
    df = pd.read_excel(file)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Clean status text
    df["Call Status"] = df["Call Status"].astype(str).str.strip()

    # Pivot summary
    summary = pd.pivot_table(
        df,
        index="Bank Name",
        columns="Call Status",
        aggfunc="size",
        fill_value=0
    )

    st.subheader("Summary Table")
    st.dataframe(summary, width="stretch")

    # Only draw chart if data exists
    if not summary.empty:

        st.subheader("Call Status Chart")

        fig, ax = plt.subplots()
        summary.plot(kind="barh", ax=ax)

        st.pyplot(fig)

    else:
        st.warning("No summary data available to plot.")

    # Export Excel
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary")

    st.download_button(
        "Download Summary Excel",
        data=output.getvalue(),
        file_name="MIS_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )