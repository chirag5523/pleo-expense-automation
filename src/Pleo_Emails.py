from urllib.parse import quote_plus
from sqlalchemy import create_engine
import pandas as pd
from pathlib import Path

# --- Configuration & Credentials ---
username = os.getenv('DB_USER', 'your_username')
password = quote_plus(os.getenv('DB_PASS', 'your_password_here'))
hostname = 'your-server-name.database.windows.net'
database_name = 'YourDatabaseName'
driver = quote_plus('ODBC Driver 17 for SQL Server')

def find_specific_excel_file(base_dir, file_name):
    """Find a specific Excel file in a designated directory."""
    full_path = Path(base_dir)
    excel_files = [file for file in full_path.rglob(file_name)]
    print(f"Looking in: {full_path}")  # Diagnostic print
    print(f"Files found: {excel_files}")  # Diagnostic print
    return excel_files

def export_to_sql(df, table_name, sql_connection_string):
    """Export DataFrame to a SQL table."""
    engine = create_engine(sql_connection_string)
    df.to_sql(table_name, con=engine, if_exists='replace', index=False)
    print("Data has been exported to the SQL server.")

# Generic path placeholder
root_directory = r'C:/Users/Default/Documents/Finance_Project/Expenses/'

# Find the specific Excel file
supplier_file_name = "Pleo Email List.xlsx"
excel_files = find_specific_excel_file(root_directory, supplier_file_name)

if not excel_files:
    print("No Excel files found.")
else:
    try:
        # Load the first found Excel file into a DataFrame, specifically from the 'Data' sheet
        # Note: Your code comment says 'Data' sheet but the code uses 'Sheet1'
        combined_df = pd.read_excel(excel_files[0], sheet_name='Sheet1')

        # Export the DataFrame to SQL
        export_to_sql(combined_df, 'Pleo_Email_List', connection_string)
        print("Exported final merged data to SQL.")

    except Exception as e:
        print(f"An error occurred: {e}")