import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ERP Data Processor", layout="centered")
st.title("ERP Data Processor")

uploaded_file = st.file_uploader("Upload Consumption.xlsx", type=["xlsx"])

if uploaded_file:
    consumption = pd.read_excel(uploaded_file)

    selected_columns = ['Report Date', 'UTC Date & Time', 'Event', 'From Port', 'To Port',
                        'Steaming time (HRS)', 'Obs distance (NM)', 'Distance Travelled during HS (NM)',
                        'Time Spent at Anchorage (Hrs)', 'Time Spent at Drifting (Hrs)',
                        'Total cargo on board (MT)', 'AE LS MGO consumption (MT)',
                        'ME LS MGO consumption (MT)', 'BLR LS MGO consumption (MT)',
                        'AE VLSFO consumption (MT)', 'ME VLSFO consumption (MT)',
                        'BLR VLSFO consumption (MT)', 'Total Consumption LSMGO',
                        'Total Consumption VLSFO', 'ROB LS MGO', 'ROB VLSFO',
                        'Bunker Recvd LSMGO', 'Bunker Recvd HSMGO',
                        'Survey Correction LSMGO', 'Survey Correction VLSFO']

    ERP = consumption[selected_columns]

    st.subheader("Processed ERP Data Table:")
    st.dataframe(ERP, use_container_width=True)

    # Excel output with borders and wrap text
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        ERP.to_excel(writer, index=False, sheet_name="ERP Data")
        workbook = writer.book
        worksheet = writer.sheets["ERP Data"]

        bordered_wrap_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'top'
        })

        (max_row, max_col) = ERP.shape
        worksheet.conditional_format(0, 0, max_row, max_col - 1,
                                     {'type': 'no_errors', 'format': bordered_wrap_format})
        worksheet.set_column(0, max_col - 1, 20)

    st.download_button(label="Download ERP File with Borders & Wrap Text",
                       data=output.getvalue(),
                       file_name="ERP_Processed_Styled.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
