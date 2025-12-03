import os
import google.generativeai as genai
import json

class AIAnalyzer:
    """
    Interfaces with an LLM to generate insights from data summaries.
    """
    def __init__(self, api_key=None):
        self.api_key = api_key or os.getenv("GOOGLE_API_KEY")
        self.model = None
        if self.api_key:
            try:
                genai.configure(api_key=self.api_key)
                self.model = genai.GenerativeModel('gemini-2.5-flash')
            except Exception as e:
                print(f"Failed to initialize Google Gemini client: {e}")

    def analyze_summary(self, summary_stats):
        """
        Sends the data summary to the LLM and retrieves insights.
        """
        if not self.model:
            return "AI Analysis skipped: No API Key provided. Please set GOOGLE_API_KEY environment variable to enable AI insights."
        
        # Convert summary to string if it's a dict
        if isinstance(summary_stats, dict):
            summary_str = json.dumps(summary_stats, indent=2, default=str)
        else:
            summary_str = str(summary_stats)

        prompt = f"""
You are a senior consultant at McKinsey & Company. You have been provided with a dataset summary for analysis.

DATASET SUMMARY:
{summary_str}

INSTRUCTIONS:
- Conduct a comprehensive statistical and market-aligned analysis of this dataset
- Provide deep, actionable insights that a C-suite executive would find valuable
- Focus on patterns, trends, correlations, outliers, and business implications
- Use statistical reasoning and data-driven conclusions ONLY
- DO NOT hallucinate or make assumptions beyond what the data shows
- DO NOT summarize the dataset - provide analytical insights instead
- Structure your analysis professionally with clear sections

REQUIRED ANALYSIS SECTIONS:
1. Key Statistical Findings (distributions, central tendencies, variance)
2. Notable Patterns & Trends (what the data reveals)
3. Anomalies & Outliers (if any significant deviations exist)
4. Business Implications (what this means for decision-makers)
5. Recommended Actions (data-driven next steps)

Provide your analysis in a clear, professional format suitable for an executive report.
"""
        
        try:
            print("Sending data summary to AI for analysis...")
            response = self.model.generate_content(prompt)
            return response.text
        except Exception as e:
            print(f"Error during AI analysis: {e}")
            return f"Error during AI analysis: {e}"
