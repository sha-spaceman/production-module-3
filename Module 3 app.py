import streamlit as st
import pandas as pd
import io

# Set up the page configuration
st.set_page_config(page_title="Exam File Parser", page_icon="📄", layout="centered")

st.title("Examination File Name Parser & Aggregator")
st.write("Upload your Excel list of PNG filenames to parse, aggregate, and generate the `QP_MS_Parts` output.")

# --- 1. UI: Examination Board Selection ---
board_options = {
    "1. Cambridge (CAIE)": 1,
    "2. Pearson Edexcel": 2,
    "3. IBDP": 3
}
selected_board_label = st.radio("Select the Examination Board:", options=list(board_options.keys()))
board_choice = board_options[selected_board_label]

st.divider()

# --- 2. UI: File Upload ---
uploaded_file = st.file_uploader(
    "Upload Input Excel File (single column, no header)", 
    type=["xlsx", "xls"]
)

# --- 3. Processing Logic ---
if uploaded_file is not None:
    if st.button("Process Data", type="primary"):
        with st.spinner("Parsing files and generating queries..."):
            try:
                # Load the Excel sheet
                df_input = pd.read_excel(uploaded_file, header=None)
                if df_input.empty:
                    st.error("The uploaded Excel file contains no data.")
                    st.stop()
                
                # Extract filenames from the first column
                filenames = df_input.iloc[:, 0].astype(str).tolist()
                
                table_data = []

                # Core Parsing Logic
                for name in filenames:
                    qp_ms = ""
                    q_num = ""
                    q_part = ""
                    length = len(name)

                    if board_choice == 1:
                        qp_ms = name[:14]
                        q_num = name[15:17] if length >= 17 else ""
                        q_part = name[18] if length == 23 else ""

                    elif board_choice == 2:
                        is_r = (name[7] == "r") if length >= 8 else False
                        if is_r:
                            qp_ms = name[:21]
                            q_num = name[22:24] if length >= 24 else ""
                        else:
                            qp_ms = name[:20]
                            q_num = name[21:23] if length >= 23 else ""

                        if length == 29:
                            q_part = name[24]
                        elif length == 30:
                            q_part = name[25]
                        else:
                            q_part = ""

                    elif board_choice == 3:
                        qp_ms = name[:29]
                        q_num = name[15:17] if length >= 17 else ""
                        q_part = name[18] if length == 23 else ""

                    table_data.append(
                        {
                            "Q paper/Mark Scheme": qp_ms,
                            "Q Number": q_num,
                            "Q part": q_part,
                        }
                    )

                # Convert to DataFrame
                df_info = pd.DataFrame(table_data)

                # Determine query filters based on the selected board
                if board_choice in [1, 3]:
                    filter_q = "qp"
                    filter_ms = "ms"
                else:  
                    filter_q = "que"
                    filter_ms = "rms"

                # Query 1 - QParts
                df_qparts = df_info[
                    df_info["Q paper/Mark Scheme"].str.contains(filter_q, case=False, na=False)
                ]
                df_qparts_grouped = (
                    df_qparts.groupby(["Q paper/Mark Scheme", "Q Number"], as_index=False)["Q part"]
                    .max()
                    .rename(columns={"Q part": "MaxOfQ part"})
                )
                # Replace blanks/nulls with 1
                df_qparts_grouped["MaxOfQ part"] = df_qparts_grouped["MaxOfQ part"].replace("", 1).fillna(1)

                # Query 2 - MSParts
                df_msparts = df_info[
                    df_info["Q paper/Mark Scheme"].str.contains(filter_ms, case=False, na=False)
                ]
                df_msparts_grouped = (
                    df_msparts.groupby(["Q paper/Mark Scheme", "Q Number"], as_index=False)["Q part"]
                    .max()
                    .rename(columns={"Q part": "MaxOfQ part"})
                )
                # Replace blanks/nulls with 1
                df_msparts_grouped["MaxOfQ part"] = df_msparts_grouped["MaxOfQ part"].replace("", 1).fillna(1)

                # --- 4. Export to Memory buffer ---
                # We use BytesIO to save the file in memory rather than on disk
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                    df_qparts_grouped.to_excel(writer, sheet_name="QParts", index=False)
                    df_msparts_grouped.to_excel(writer, sheet_name="MSParts", index=False)
                
                processed_data = output_buffer.getvalue()

                st.success("Queries processed successfully!")

                # --- 5. UI: Download Button ---
                st.download_button(
                    label="📥 Download QP_MS_Parts.xlsx",
                    data=processed_data,
                    file_name="QP_MS_Parts.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")