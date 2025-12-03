import pandas as pd
import json
import sqlite3
import os

class DataLoader:
    """
    Handles loading data from various sources into a pandas DataFrame.
    """
    def __init__(self):
        pass

    def load_file(self, filepath):
        """
        Determines file type by extension and loads it.
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")

        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.csv':
            return self.load_csv(filepath)
        elif ext == '.json':
            return self.load_json(filepath)
        elif ext in ['.xls', '.xlsx']:
            return self.load_excel(filepath)
        else:
            raise ValueError(f"Unsupported file extension: {ext}")

    def load_csv(self, filepath):
        try:
            print(f"Loading CSV from {filepath}...")
            return pd.read_csv(filepath)
        except Exception as e:
            print(f"Error loading CSV: {e}")
            raise

    def load_json(self, filepath):
        try:
            print(f"Loading JSON from {filepath}...")
            # Try reading as records first, then default
            try:
                return pd.read_json(filepath, orient='records')
            except ValueError:
                return pd.read_json(filepath)
        except Exception as e:
            print(f"Error loading JSON: {e}")
            raise

    def load_excel(self, filepath):
        try:
            print(f"Loading Excel from {filepath}...")
            return pd.read_excel(filepath)
        except Exception as e:
            print(f"Error loading Excel: {e}")
            raise

    def load_sql(self, db_path, query):
        """
        Loads data from a SQLite database.
        """
        try:
            print(f"Loading data from SQL database {db_path} with query: {query}")
            conn = sqlite3.connect(db_path)
            df = pd.read_sql_query(query, conn)
            conn.close()
            return df
        except Exception as e:
            print(f"Error loading SQL: {e}")
            raise
