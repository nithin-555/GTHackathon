import argparse
import os
import sys
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Ensure src is in path if running from root
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.ingestion import DataLoader
from src.processing import DataProcessor
from src.analysis import AIAnalyzer
from src.reporting_pdf import PDFReporter
from src.reporting_ppt import PPTReporter

def main():
    parser = argparse.ArgumentParser(description="Automated Data Pipeline")
    parser.add_argument("--source", required=True, help="Path to input data file")
    parser.add_argument("--output-format", choices=['pdf', 'ppt', 'all'], default='all', help="Output report format")
    parser.add_argument("--output-dir", default="output", help="Directory to save reports")
    parser.add_argument("--api-key", help="Google API Key (optional, can also set GOOGLE_API_KEY env var)")
    
    args = parser.parse_args()
    
    if args.api_key:
        os.environ["GOOGLE_API_KEY"] = args.api_key
    
    # 1. Ingestion
    print(f"Starting pipeline for {args.source}...")
    loader = DataLoader()
    try:
        df = loader.load_file(args.source)
    except Exception as e:
        print(f"Failed to load data: {e}")
        return

    # 2. Processing
    print("Processing data...")
    processor = DataProcessor()
    df_clean = processor.clean_data(df)
    summary = processor.summarize_data(df_clean)
    
    # 3. Analysis
    print("Analyzing data...")
    analyzer = AIAnalyzer()
    insights = analyzer.analyze_summary(summary)
    print("Insights generated.")
    
    # 4. Reporting
    base_name = os.path.splitext(os.path.basename(args.source))[0]
    
    if args.output_format in ['pdf', 'all']:
        pdf_path = os.path.join(args.output_dir, f"{base_name}_report.pdf")
        print(f"Generating PDF report at {pdf_path}...")
        reporter = PDFReporter()
        reporter.generate_report(summary, insights, pdf_path, df_clean)
        
    if args.output_format in ['ppt', 'all']:
        ppt_path = os.path.join(args.output_dir, f"{base_name}_report.pptx")
        print(f"Generating PPT report at {ppt_path}...")
        reporter = PPTReporter()
        reporter.generate_report(summary, insights, ppt_path, df_clean)
        
    print("Pipeline completed successfully.")

if __name__ == "__main__":
    main()
