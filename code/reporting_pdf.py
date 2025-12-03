from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import os
import pandas as pd
import seaborn as sns
import numpy as np

class PDFReporter:
    """
    Generates professional PDF reports with visualizations.
    """
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        
    def _setup_custom_styles(self):
        """Setup custom paragraph styles"""
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=28,
            textColor=colors.HexColor('#1a1a1a'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))
        
        self.styles.add(ParagraphStyle(
            name='Subtitle',
            parent=self.styles['Normal'],
            fontSize=14,
            textColor=colors.HexColor('#666666'),
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))
        
        self.styles.add(ParagraphStyle(
            name='CustomHeading',
            parent=self.styles['Heading2'],
            fontSize=18,
            textColor=colors.HexColor('#0066CC'),
            spaceAfter=14,
            spaceBefore=16,
            fontName='Helvetica-Bold'
        ))
        
        self.styles.add(ParagraphStyle(
            name='CustomBody',
            parent=self.styles['BodyText'],
            fontSize=11,
            spaceAfter=10,
            alignment=TA_JUSTIFY,
            leading=16
        ))
        
        self.styles.add(ParagraphStyle(
            name='Narrative',
            parent=self.styles['BodyText'],
            fontSize=12,
            spaceAfter=12,
            alignment=TA_JUSTIFY,
            leading=18,
            textColor=colors.HexColor('#333333')
        ))
    
    def _generate_narrative_summary(self, summary_stats, df):
        """Generate natural language narrative about the dataset"""
        overview = summary_stats.get('dataset_overview', {})
        col_info = summary_stats.get('column_information', {})
        
        total_rows = overview.get('total_rows', 0)
        total_cols = overview.get('total_columns', 0)
        
        # Identify column types
        numeric_cols = [col for col, info in col_info.items() if 'statistics' in info]
        categorical_cols = [col for col, info in col_info.items() if 'statistics' not in info]
        
        # Build narrative
        narrative = f"This analysis examines a comprehensive dataset comprising {total_rows:,} observations "
        narrative += f"across {total_cols} distinct dimensions. "
        
        # Describe what the data represents based on column names
        if df is not None and len(numeric_cols) > 0:
            # Try to infer dataset purpose from column names
            sample_metrics = [col.replace('_', ' ').title() for col in numeric_cols[:3]]
            narrative += f"The data captures key performance metrics including {', '.join(sample_metrics)}. "
        
        # Data richness
        if len(categorical_cols) > 0:
            narrative += f"Spanning {len(categorical_cols)} categorical dimensions alongside {len(numeric_cols)} quantitative measurements, "
            narrative += "this dataset provides a multi-faceted view of the underlying business operations. "
        
        # Data quality
        null_percentages = [info.get('null_percentage', 0) for info in col_info.values()]
        avg_completeness = 100 - (sum(null_percentages) / len(null_percentages) if null_percentages else 0)
        
        if avg_completeness > 95:
            narrative += f"With {avg_completeness:.1f}% data completeness, the dataset exhibits exceptional quality and reliability. "
        elif avg_completeness > 80:
            narrative += f"The data maintains strong quality standards with {avg_completeness:.1f}% completeness. "
        else:
            narrative += f"The dataset shows {avg_completeness:.1f}% completeness, indicating opportunities for enhanced data collection. "
        
        # Business context
        narrative += "Through systematic analysis, we uncover actionable patterns and strategic insights that can inform executive decision-making and drive measurable business outcomes."
        
        return narrative
    
    def generate_report(self, summary_stats, insights, output_path, df=None):
        """
        Creates a professional PDF file with visualizations.
        """
        try:
            # Create output directory
            output_dir = os.path.dirname(output_path) if os.path.dirname(output_path) else 'output'
            os.makedirs(output_dir, exist_ok=True)
            
            doc = SimpleDocTemplate(output_path, pagesize=letter,
                                   rightMargin=72, leftMargin=72,
                                   topMargin=72, bottomMargin=50)
            
            story = []
            
            # Title Page
            story.append(Spacer(1, 1.5*inch))
            story.append(Paragraph("Executive Data Analysis", self.styles['CustomTitle']))
            story.append(Paragraph("Strategic Insights Report", self.styles['Subtitle']))
            story.append(Spacer(1, 0.5*inch))
            
            # Decorative line
            story.append(Paragraph("_" * 80, self.styles['Normal']))
            story.append(Spacer(1, 0.3*inch))
            
            # Narrative Summary
            story.append(Paragraph("Executive Overview", self.styles['CustomHeading']))
            narrative = self._generate_narrative_summary(summary_stats, df)
            story.append(Paragraph(narrative, self.styles['Narrative']))
            story.append(Spacer(1, 0.3*inch))
            
            # Key Highlights Box
            overview = summary_stats.get('dataset_overview', {})
            highlights = [
                ['Total Observations', f"{overview.get('total_rows', 'N/A'):,}"],
                ['Data Dimensions', f"{overview.get('total_columns', 'N/A')} attributes"],
                ['Dataset Size', f"{overview.get('memory_usage_mb', 0):.2f} MB"]
            ]
            
            t = Table(highlights, colWidths=[2.5*inch, 3*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f0f0')),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#333333')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
            ]))
            
            story.append(t)
            story.append(PageBreak())
            
            # Visualizations
            if df is not None and not df.empty:
                story.append(Paragraph("Data Visualization & Analysis", self.styles['CustomHeading']))
                story.append(Spacer(1, 0.2*inch))
                
                # 1. Numeric distributions
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    viz_path = self._create_numeric_distribution(df, numeric_cols)
                    if viz_path and os.path.exists(viz_path):
                        story.append(Paragraph("Quantitative Metrics Distribution", self.styles['CustomBody']))
                        story.append(Spacer(1, 0.1*inch))
                        story.append(Image(viz_path, width=6.5*inch, height=4*inch))
                        story.append(Spacer(1, 0.2*inch))
                        os.remove(viz_path)
                
                # 2. Categorical analysis
                categorical_cols = df.select_dtypes(include=['object']).columns
                if len(categorical_cols) > 0:
                    viz_path = self._create_categorical_distribution(df, categorical_cols)
                    if viz_path and os.path.exists(viz_path):
                        story.append(Paragraph("Categorical Analysis", self.styles['CustomBody']))
                        story.append(Spacer(1, 0.1*inch))
                        story.append(Image(viz_path, width=6.5*inch, height=3.5*inch))
                        story.append(Spacer(1, 0.2*inch))
                        os.remove(viz_path)
                
                # 3. Correlation heatmap (if multiple numeric columns)
                if len(numeric_cols) >= 2:
                    viz_path = self._create_correlation_heatmap(df, numeric_cols)
                    if viz_path and os.path.exists(viz_path):
                        story.append(PageBreak())
                        story.append(Paragraph("Correlation Analysis", self.styles['CustomBody']))
                        story.append(Spacer(1, 0.1*inch))
                        story.append(Image(viz_path, width=5.5*inch, height=4.5*inch))
                        story.append(Spacer(1, 0.2*inch))
                        os.remove(viz_path)
            
            # AI Insights
            story.append(PageBreak())
            story.append(Paragraph("Strategic Insights & Recommendations", self.styles['CustomHeading']))
            story.append(Spacer(1, 0.2*inch))
            
            if insights and not insights.startswith("AI Analysis skipped"):
                # Parse and format insights professionally
                sections = insights.split('\n\n')
                for section in sections:
                    if not section.strip():
                        continue
                    
                    lines = section.split('\n')
                    for line in lines:
                        if not line.strip():
                            continue
                        
                        line = line.strip()
                        
                        # Section headers
                        if any(header in line for header in ['Key Statistical', 'Notable Patterns', 'Anomalies', 'Business Implications', 'Recommended Actions']):
                            story.append(Spacer(1, 0.15*inch))
                            clean_line = line.replace('#', '').replace('*', '').strip()
                            if clean_line and clean_line[0].isdigit():
                                clean_line = '. '.join(clean_line.split('.')[1:]).strip()
                            story.append(Paragraph(f"<b>{clean_line}</b>", self.styles['CustomBody']))
                        # Bullet points
                        elif line.startswith('-') or line.startswith('•') or line.startswith('*'):
                            story.append(Paragraph(f"• {line[1:].strip()}", self.styles['CustomBody']))
                        # Regular text
                        elif not (line and line[0].isdigit()):
                            story.append(Paragraph(line, self.styles['CustomBody']))
            else:
                story.append(Paragraph(insights, self.styles['CustomBody']))
            
            # Build PDF
            doc.build(story)
            print(f"✓ Main PDF Report generated at {output_path}")
            
            # Generate separate technical data summary
            summary_path = output_path.replace('_report.pdf', '_technical_summary.pdf')
            self._generate_technical_summary(summary_stats, summary_path)
            
            return True
        except Exception as e:
            print(f"❌ Error generating PDF: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _create_numeric_distribution(self, df, numeric_cols):
        """Create professional distribution plots for numeric columns"""
        try:
            cols_to_plot = list(numeric_cols[:6])
            if not cols_to_plot:
                return None
            
            n_cols = min(3, len(cols_to_plot))
            n_rows = (len(cols_to_plot) + n_cols - 1) // n_cols
            
            fig, axes = plt.subplots(n_rows, n_cols, figsize=(12, 3.5*n_rows))
            fig.patch.set_facecolor('white')
            
            if n_rows == 1 and n_cols == 1:
                axes = np.array([axes])
            elif n_rows == 1:
                axes = np.array([axes])
            else:
                axes = axes.flatten()
            
            for idx, col in enumerate(cols_to_plot):
                ax = axes[idx] if len(cols_to_plot) > 1 else axes[0]
                
                # Create histogram with KDE
                data = df[col].dropna()
                ax.hist(data, bins=25, color='#3498db', alpha=0.7, edgecolor='black', density=True)
                
                # Add KDE line
                try:
                    data.plot(kind='kde', ax=ax, color='#e74c3c', linewidth=2, secondary_y=False)
                except:
                    pass
                
                ax.set_title(f'{col}', fontsize=11, fontweight='bold', pad=10)
                ax.set_xlabel('')
                ax.set_ylabel('Density', fontsize=9)
                ax.grid(axis='y', alpha=0.3, linestyle='--')
                
                # Add statistics box
                mean_val = data.mean()
                median_val = data.median()
                ax.axvline(mean_val, color='red', linestyle='--', linewidth=1.5, alpha=0.7)
                ax.axvline(median_val, color='green', linestyle='--', linewidth=1.5, alpha=0.7)
                
                stats_text = f'Mean: {mean_val:.2f}\nMedian: {median_val:.2f}'
                ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, 
                       fontsize=8, verticalalignment='top',
                       bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
            
            # Hide extra subplots
            for idx in range(len(cols_to_plot), len(axes)):
                fig.delaxes(axes[idx])
            
            plt.tight_layout()
            temp_path = 'temp_numeric_dist.png'
            plt.savefig(temp_path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close()
            
            return temp_path
        except Exception as e:
            print(f"Error creating numeric distribution: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _create_categorical_distribution(self, df, categorical_cols):
        """Create professional bar charts for categorical columns"""
        try:
            # Select up to 3 columns with reasonable cardinality
            cols_to_plot = [col for col in categorical_cols if 2 <= df[col].nunique() <= 15][:3]
            if not cols_to_plot:
                return None
            
            n_plots = len(cols_to_plot)
            fig, axes = plt.subplots(1, n_plots, figsize=(5*n_plots, 4))
            fig.patch.set_facecolor('white')
            
            if n_plots == 1:
                axes = [axes]
            
            colors_palette = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6']
            
            for idx, col in enumerate(cols_to_plot):
                value_counts = df[col].value_counts().head(8)
                
                bars = axes[idx].barh(range(len(value_counts)), value_counts.values, 
                                      color=colors_palette[idx % len(colors_palette)],
                                      edgecolor='black', alpha=0.8)
                
                axes[idx].set_yticks(range(len(value_counts)))
                axes[idx].set_yticklabels(value_counts.index, fontsize=9)
                axes[idx].set_title(f'{col}', fontsize=11, fontweight='bold', pad=10)
                axes[idx].set_xlabel('Count', fontsize=9)
                axes[idx].grid(axis='x', alpha=0.3, linestyle='--')
                
                # Add value labels
                for i, v in enumerate(value_counts.values):
                    axes[idx].text(v + max(value_counts.values)*0.01, i, f'{v:,}', 
                                 va='center', fontsize=8, fontweight='bold')
            
            plt.tight_layout()
            temp_path = 'temp_cat_dist.png'
            plt.savefig(temp_path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close()
            
            return temp_path
        except Exception as e:
            print(f"Error creating categorical distribution: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _create_correlation_heatmap(self, df, numeric_cols):
        """Create correlation heatmap for numeric columns"""
        try:
            # Limit to first 10 numeric columns for readability
            cols = list(numeric_cols[:10])
            corr_matrix = df[cols].corr()
            
            fig, ax = plt.subplots(figsize=(8, 6))
            fig.patch.set_facecolor('white')
            
            sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='coolwarm', 
                       center=0, square=True, linewidths=1, 
                       cbar_kws={"shrink": 0.8}, ax=ax, vmin=-1, vmax=1)
            
            ax.set_title('Correlation Matrix', fontsize=14, fontweight='bold', pad=15)
            plt.xticks(rotation=45, ha='right', fontsize=9)
            plt.yticks(rotation=0, fontsize=9)
            
            plt.tight_layout()
            temp_path = 'temp_correlation.png'
            plt.savefig(temp_path, dpi=150, bbox_inches='tight', facecolor='white')
            plt.close()
            
            return temp_path
        except Exception as e:
            print(f"Error creating correlation heatmap: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _generate_technical_summary(self, summary_stats, output_path):
        """Generate separate technical data summary PDF"""
        try:
            doc = SimpleDocTemplate(output_path, pagesize=letter,
                                   rightMargin=72, leftMargin=72,
                                   topMargin=72, bottomMargin=50)
            
            story = []
            
            story.append(Paragraph("Technical Data Summary", self.styles['CustomTitle']))
            story.append(Paragraph("Detailed Schema & Statistics", self.styles['Subtitle']))
            story.append(Spacer(1, 0.5*inch))
            
            # Dataset Overview
            story.append(Paragraph("Dataset Overview", self.styles['CustomHeading']))
            overview = summary_stats.get('dataset_overview', {})
            
            overview_data = [
                ['Metric', 'Value'],
                ['Total Rows', f"{overview.get('total_rows', 'N/A'):,}"],
                ['Total Columns', f"{overview.get('total_columns', 'N/A')}"],
                ['Memory Usage', f"{overview.get('memory_usage_mb', 0):.2f} MB"]
            ]
            
            t = Table(overview_data, colWidths=[3*inch, 3*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0066CC')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f0f0f0')),
                ('GRID', (0, 0), (-1, -1), 1, colors.grey)
            ]))
            
            story.append(t)
            story.append(Spacer(1, 0.3*inch))
            
            # Column Information
            story.append(Paragraph("Column Specifications", self.styles['CustomHeading']))
            story.append(Spacer(1, 0.1*inch))
            
            col_info = summary_stats.get('column_information', {})
            
            for col_name, col_data in col_info.items():
                story.append(Paragraph(f"<b>{col_name}</b>", self.styles['CustomBody']))
                
                details = [
                    f"  • Data Type: {col_data.get('data_type', 'N/A')}",
                    f"  • Non-null Count: {col_data.get('non_null_count', 'N/A'):,}",
                    f"  • Null Count: {col_data.get('null_count', 'N/A'):,} ({col_data.get('null_percentage', 0):.1f}%)",
                    f"  • Unique Values: {col_data.get('unique_values', 'N/A'):,}"
                ]
                
                if 'statistics' in col_data:
                    stats = col_data['statistics']
                    details.extend([
                        f"  • Mean: {stats.get('mean', 'N/A'):.2f}",
                        f"  • Median: {stats.get('median', 'N/A'):.2f}",
                        f"  • Std Dev: {stats.get('std_dev', 'N/A'):.2f}",
                        f"  • Range: [{stats.get('min', 'N/A'):.2f}, {stats.get('max', 'N/A'):.2f}]"
                    ])
                
                if 'top_5_values' in col_data:
                    top_vals = col_data['top_5_values']
                    details.append(f"  • Top Values: {', '.join([f'{k} ({v})' for k, v in list(top_vals.items())[:3]])}")
                
                for detail in details:
                    story.append(Paragraph(detail, self.styles['Normal']))
                
                story.append(Spacer(1, 0.15*inch))
            
            doc.build(story)
            print(f"✓ Technical Summary PDF generated at {output_path}")
            
        except Exception as e:
            print(f"❌ Error generating technical summary: {e}")
            import traceback
            traceback.print_exc()