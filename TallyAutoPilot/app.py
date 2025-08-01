
import streamlit as st
import pandas as pd
import yaml
from core.tally_api import post_to_tally
from core.gpt_classifier import extract_from_invoice
from core.email_notify import send_email_notification

# Load config
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

st.markdown(
    """
    <style>
        .tally-header-bg {
            background: #111111;
            border-radius: 18px;
            padding: 38px 0 28px 0;
            box-shadow: 0 4px 32px 0 rgba(0,0,0,0.18);
            margin-top: 40px;
            margin-bottom: 30px;
        }
        .tally-title {
            font-size: 2.8em;
            color: #f8f8f8;
            letter-spacing: 1.5px;
            font-family: 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
            font-weight: 700;
            margin-bottom: 0.2em;
            text-shadow: 0 2px 12px #00000055;
        }
        .tally-subtitle {
            color: #b0b0b0;
            font-weight: 400;
            font-size: 1.25em;
            font-style: italic;
            font-family: 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
            margin-top: 0;
            letter-spacing: 0.5px;
        }
    </style>
    <div class="tally-header-bg" style="text-align: center;">
        <h1 class="tally-title">
            TallyAutoPilot
        </h1>
        <h3 class="tally-subtitle">
            AI-Powered Tally Data Entry Tool
        </h3>
    </div>
    """,
    unsafe_allow_html=True
)

mode = st.radio("Mode", ["Data Entry - Excel File Upload"])

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.dataframe(df)
    if st.button("Post to Tally"):
        for idx, row in df.iterrows():
            try:
                res = post_to_tally(row, config)
                send_email_notification(
                    subject="TallyEntry Success",
                    body=f"Posted {row['PartyName']} - Rs.{row['Amount']}"
                )
                st.success(f"Row {idx+1}: Success")
            except Exception as e:
                st.error(f"Row {idx+1}: Failed - {e}")

