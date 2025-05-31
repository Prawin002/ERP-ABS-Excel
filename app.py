import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter

st.title("ERP Data Processor")

uploaded_file = st.file_uploader("Upload Consumption.xlsx", type=["xlsx"])
if uploaded_file:
    consumption = pd.read_excel(uploaded_file)
    ERP = consumption[['Report Date', 'UTC Date & Time', 'Event', 'From Port',
                       'Steaming time (HRS)', 'Obs distance (NM)', 'Distance Travelled during HS (NM)',
                       'Time Spent at Anchorage (Hrs)', 'Time Spent at Drifting (Hrs)',
                       'Total cargo on board (MT)', 'AE LS MGO consumption (MT)',
                       'ME LS MGO consumption (MT)', 'BLR LS MGO consumption (MT)',
                       'AE VLSFO consumption (MT)', 'ME VLSFO consumption (MT)',
                       'BLR VLSFO consumption (MT)', 'Total Consumption LSMGO',
                       'Total Consumption VLSFO', 'ROB LS MGO', 'ROB VLSFO',
                       'Bunker Recvd LSMGO', 'Bunker Recvd HSMGO',
                       'Survey Correction LSMGO', 'Survey Correction VLSFO']]

    st.write("Processed Data:")
    st.dataframe(ERP)

    # Convert DataFrame to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        ERP.to_excel(writer, index=False)
        writer.close()

    # Provide download link
    st.download_button(label="Download Processed File",
                       data=output.getvalue(),
                       file_name="ERP.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
