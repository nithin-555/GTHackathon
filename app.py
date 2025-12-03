import streamlit as st
import os
import sys
from dotenv import load_dotenv
import pandas as pd

# Load environment variables
load_dotenv()

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.ingestion import DataLoader
from src.processing import DataProcessor
from src.analysis import AIAnalyzer
from src.reporting_pdf import PDFReporter
from src.reporting_ppt import PPTReporter

# Page config
st.set_page_config(
    page_title="Automated Data Pipeline",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 Automated Data Analysis ")
st.markdown("Upload your data file and generate reports")

# Sidebar for configuration
st.sidebar.header("⚙️ Configuration")

# File uploader
uploaded_file = st.file_uploader(
    "Choose a data file",
    type=['csv', 'json', 'xlsx', 'xls'],
    help="Upload CSV, JSON, or Excel files"
)

# Output format selection
output_format = st.sidebar.radio(
    "Select Output Format",
    options=['Both PDF & PPT', 'PDF Only', 'PPT Only'],
    index=0
)

# API Key input (optional)
api_key_input = st.sidebar.text_input(
    "Google API Key (Optional)",
    type="password",
    help="Leave empty to use .env file"
)

# Generate button
generate_button = st.button("🚀 Generate Report", type="primary", use_container_width=True)

# Main processing logic
if generate_button:
    if uploaded_file is None:
        st.error("⚠️ Please upload a file first!")
    else:
        # Create progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # Save uploaded file temporarily
            temp_dir = "temp"
            os.makedirs(temp_dir, exist_ok=True)
            temp_file_path = os.path.join(temp_dir, uploaded_file.name)
            
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            status_text.text("📥 Loading data...")
            progress_bar.progress(20)
            
            # 1. Ingestion
            loader = DataLoader()
            df = loader.load_file(temp_file_path)
            
            st.success(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")
            
            # Show data preview
            with st.expander("👀 Data Preview"):
                st.dataframe(df.head(10), use_container_width=True)
            
            status_text.text("🔄 Processing data...")
            progress_bar.progress(40)
            
            # 2. Processing
            processor = DataProcessor()
            df_clean = processor.clean_data(df)
            summary = processor.summarize_data(df_clean)
            
            status_text.text("🤖 Analyzing with AI...")
            progress_bar.progress(60)
            
            # 3. Analysis
            if api_key_input:
                os.environ["GOOGLE_API_KEY"] = api_key_input
            
            analyzer = AIAnalyzer()
            insights = analyzer.analyze_summary(summary)
            
            # Show insights
            with st.expander("💡 AI Insights", expanded=True):
                st.markdown(insights)
            
            status_text.text("📄 Generating reports...")
            progress_bar.progress(80)
            
            # 4. Reporting
            output_dir = "output"
            os.makedirs(output_dir, exist_ok=True)
            base_name = os.path.splitext(uploaded_file.name)[0]
            
            generated_files = []
            
            if output_format in ['Both PDF & PPT', 'PDF Only']:
                pdf_path = os.path.join(output_dir, f"{base_name}_report.pdf")
                pdf_reporter = PDFReporter()
                pdf_reporter.generate_report(summary, insights, pdf_path, df_clean)
                generated_files.append(("PDF", pdf_path))
            
            if output_format in ['Both PDF & PPT', 'PPT Only']:
                ppt_path = os.path.join(output_dir, f"{base_name}_report.pptx")
                ppt_reporter = PPTReporter()
                ppt_reporter.generate_report(summary, insights, ppt_path, df_clean)
                generated_files.append(("PPT", ppt_path))
            
            progress_bar.progress(100)
            status_text.text("✅ Complete!")
            
            st.success("🎉 Reports generated successfully!")
            
            # Download buttons
            st.markdown("### 📥 Download Reports")
            cols = st.columns(len(generated_files))
            
            for idx, (file_type, file_path) in enumerate(generated_files):
                with open(file_path, "rb") as f:
                    cols[idx].download_button(
                        label=f"⬇️ Download {file_type}",
                        data=f,
                        file_name=os.path.basename(file_path),
                        mime="application/pdf" if file_type == "PDF" else "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
            
            # Clean up temp file
            os.remove(temp_file_path)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            progress_bar.empty()
            status_text.empty()

# Footer
st.sidebar.markdown("---")
st.sidebar.markdown("### 📖 Instructions")
st.sidebar.markdown("""
1. Upload your data file (CSV, JSON, Excel)
2. Select output format
3. Click 'Generate Report'
4. Download your reports!
""")

st.sidebar.markdown("---")
st.sidebar.info("💡 **Tip**: Your Google API key is loaded from `.env` file automatically")
