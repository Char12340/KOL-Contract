import pandas as pd
from docxtpl import DocxTemplate
from datetime import date
import streamlit as st
import io
import zipfile

st.set_page_config(page_title="KOL Contract Generator", layout="centered", page_icon="📝")

st.markdown("""
<style>
    .main {
        background-color: #f9f9fb;
        font-family: 'Segoe UI', sans-serif;
    }
    .title {
        font-size: 2.2em;
        font-weight: bold;
        color: #4a4a4a;
    }
    .subtitle {
        font-size: 1.1em;
        color: #6c6c6c;
    }
    .stFileUploader > label {
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="title">📄 KOC Contract Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Coded by Char</div>', unsafe_allow_html=True)
st.markdown("---")

# Upload section
col1, col2 = st.columns(2)
with col1:
    uploaded_csv = st.file_uploader("📑 Upload CSV File", type=["csv"])
with col2:
    uploaded_template = st.file_uploader("📄 Upload Word Template (.docx)", type=["docx"])

# Process files
if uploaded_csv and uploaded_template:
    st.success("✅ Files uploaded successfully!")
    df = pd.read_csv(uploaded_csv)
    df.columns = df.columns.str.strip()
    today = date.today().isoformat()
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for index, row in df.iterrows():
            try:
                template = DocxTemplate(uploaded_template)

                # Updated context
                context = {
                    'Signatory_name': row.get('Agency Name', ''),
                    'Influencer_name': row.get('Name', ''),
                    'Influencer_email': row.get('Email', ''),
                    'Influencer_contact': row.get('Contact', ''),
                    'Influencer_address': row.get('Address', ''),
                    'platform': row.get('Platform', ''),
                    'platform_username': row.get('Platform username', ''),
                    'Influencer_links': row.get('Links', ''),
                    'promotion_date': row.get('Promotion Dates', ''),
                    'promotion_timeline': row.get('Timeline', ''),
                    'Videos': row.get('video', ''),
                    'video_rate': row.get('Total Costs', ''),
                    'Each_rate': row.get('each', ''),
                    'payment_method': row.get('Payment method', ''),
                    'payment_information': row.get('Payment Info', ''),
                    'payment_charges': row.get('payment charges', '')
                }

                template.render(context)

                safe_name = row.get('Name', f'Row_{index}').replace(" ", "_").replace("/", "-")
                filename = f'IO-ARETIS_{safe_name}_{today}.docx'

                doc_stream = io.BytesIO()
                template.save(doc_stream)
                doc_stream.seek(0)

                zip_file.writestr(filename, doc_stream.read())

            except Exception as e:
                st.error(f"❌ Error processing {row.get('Name', f'Row {index}')} (row {index}): {e}")

    zip_buffer.seek(0)
    st.markdown("### ✅ All contracts generated!")
    st.download_button(
        "📥 Download ZIP of All Contracts",
        zip_buffer,
        file_name=f"KOL_Contracts_{today}.zip",
        mime="application/zip"
    )

else:
    st.info("⬆️ Upload both files above to get started.")
