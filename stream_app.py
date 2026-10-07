import tempfile
from pathlib import Path

import streamlit as st

from roi_excel_calculator_v2_dynamic_premises_fixed_v8 import (
    create_excel_download,
    process_file,
)
from frame_chat import render_chat_panel


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Frame ROI Calculator",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# HEADER
# =========================================================

st.title("Frame Services Calculator")

st.write(
    "Upload the Services Project Excel file. "
    "The application calculates Current Services and Future Services "
    "row by row using the X/Y/Z assumptions from the Premises tab."
)

# st.info(
#     "No ROI percentage is calculated. "
#     "The output contains FTE, cost, value, yearly margin contribution, "
#     "and monthly margin contribution."
# )


# =========================================================
# INPUT
# =========================================================

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"],
    help="Upload the Excel workbook containing Services Project and Premises tabs.",
)


# =========================================================
# PROCESS
# =========================================================

if uploaded_file is not None:

    st.success(f"Uploaded: {uploaded_file.name}")

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=Path(uploaded_file.name).suffix,
    ) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        input_path = temp_file.name

    try:
        with st.spinner("Calculating Current and Future Services..."):
            (
                current_df,
                future_df,
                combined_df,
                summary_df,
                future_schedule_df,
                future_current_margin_df,
            ) = process_file(input_path)

        # -----------------------------------------------------
        # SUMMARY
        # -----------------------------------------------------

        # st.subheader("Calculation Summary")

        # if not summary_df.empty:
        #     st.dataframe(
        #         summary_df,
        #         use_container_width=True,
        #         hide_index=True,
        #     )

        # -----------------------------------------------------
        # CURRENT SERVICES
        # -----------------------------------------------------

        st.subheader("Current Services")

        if current_df.empty:
            st.warning("No Current Services rows were found.")
        else:
            st.dataframe(
                current_df,
                use_container_width=True,
                hide_index=True,
            )

        # -----------------------------------------------------
        # FUTURE SERVICES
        # -----------------------------------------------------

        st.subheader("Future Services")

        if future_df.empty:
            st.warning("No Future Services rows were found.")
        else:
            st.dataframe(
                future_df,
                use_container_width=True,
                hide_index=True,
            )

        # -----------------------------------------------------
        # COST AND SCHEDULE TO REACH FUTURE SERVICES
        # -----------------------------------------------------

        st.subheader("Cost and Schedule to reach Future services")
        st.dataframe(
            future_schedule_df,
            use_container_width=True,
            hide_index=True,
        )

        # -----------------------------------------------------
        # FUTURE-CURRENT SERVICE MARGIN CONTRIBUTION
        # -----------------------------------------------------

        st.subheader("Future-Current Service Margin Contribution")
        st.dataframe(
            future_current_margin_df,
            use_container_width=True,
            hide_index=True,
        )

        # -----------------------------------------------------
        # DOWNLOAD
        # -----------------------------------------------------

        excel_data = create_excel_download(
            current_df=current_df,
            future_df=future_df,
            combined_df=combined_df,
            summary_df=summary_df,
            future_schedule_df=future_schedule_df,
            future_current_margin_df=future_current_margin_df,
        )

        st.download_button(
            label="Download Calculated Excel",
            data=excel_data,
            file_name="frame_roi_calculated_output.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
        )

        st.caption(
            "The downloaded workbook contains Summary, Current Services, "
            "Future Services, All Calculations, and Premises Used sheets."
        )

        # -----------------------------------------------------
        # CHAT WITH RESULTS
        # -----------------------------------------------------

        render_chat_panel(
            input_path=input_path,
            file_key=f"{uploaded_file.name}:{uploaded_file.size}",
            current_df=current_df,
            future_df=future_df,
            schedule_df=future_schedule_df,
            margin_df=future_current_margin_df,
        )

    except Exception as exc:
        st.error(f"Error processing Excel file: {exc}")

    finally:
        try:
            Path(input_path).unlink(missing_ok=True)
        except Exception:
            pass