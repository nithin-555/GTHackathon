import pandas as pd
import json

class DataProcessor:
    """
    Handles data cleaning and summarization.
    """
    def clean_data(self, df):
        """
        Performs basic data cleaning: drops duplicates, handles missing values.
        """
        if df is None or df.empty:
            return df

        # Drop duplicates
        initial_count = len(df)
        df = df.drop_duplicates()
        print(f"Removed {initial_count - len(df)} duplicate rows.")

        # Fill missing numeric values with mean
        numeric_cols = df.select_dtypes(include=['number']).columns
        if not numeric_cols.empty:
            df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())

        # Fill missing string values with 'Unknown'
        string_cols = df.select_dtypes(include=['object']).columns
        if not string_cols.empty:
            df[string_cols] = df[string_cols].fillna('Unknown')
            
        return df

    def summarize_data(self, df):
        """
        Generates a comprehensive, well-formatted summary of the dataset for AI analysis.
        """
        if df is None or df.empty:
            return {}

        # Basic dataset information
        summary = {
            "dataset_overview": {
                "total_rows": int(len(df)),
                "total_columns": int(len(df.columns)),
                "memory_usage_mb": round(df.memory_usage(deep=True).sum() / 1024**2, 2)
            },
            "column_information": {}
        }
        
        # Detailed column analysis
        for col in df.columns:
            col_info = {
                "data_type": str(df[col].dtype),
                "non_null_count": int(df[col].count()),
                "null_count": int(df[col].isnull().sum()),
                "null_percentage": round(df[col].isnull().sum() / len(df) * 100, 2)
            }
            
            # Numeric columns
            if pd.api.types.is_numeric_dtype(df[col]):
                col_info["statistics"] = {
                    "mean": round(df[col].mean(), 2),
                    "median": round(df[col].median(), 2),
                    "std_dev": round(df[col].std(), 2),
                    "min": round(df[col].min(), 2),
                    "max": round(df[col].max(), 2),
                    "q1": round(df[col].quantile(0.25), 2),
                    "q3": round(df[col].quantile(0.75), 2)
                }
                col_info["unique_values"] = int(df[col].nunique())
                
            # Categorical/Object columns
            else:
                col_info["unique_values"] = int(df[col].nunique())
                top_values = df[col].value_counts().head(5).to_dict()
                col_info["top_5_values"] = {str(k): int(v) for k, v in top_values.items()}
            
            summary["column_information"][col] = col_info
        
        # Overall numeric summary
        numeric_df = df.select_dtypes(include=['number'])
        if not numeric_df.empty:
            summary["numeric_summary"] = {
                "total_numeric_columns": len(numeric_df.columns),
                "correlation_insights": "High correlation detected" if (numeric_df.corr().abs() > 0.7).sum().sum() > len(numeric_df.columns) else "Low to moderate correlations"
            }
        
        # Categorical summary
        categorical_df = df.select_dtypes(include=['object'])
        if not categorical_df.empty:
            summary["categorical_summary"] = {
                "total_categorical_columns": len(categorical_df.columns),
                "high_cardinality_columns": [col for col in categorical_df.columns if categorical_df[col].nunique() > 50]
            }
        
        return summary

