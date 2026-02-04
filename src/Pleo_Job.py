import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus
from pathlib import Path

# Credentials and connection details
username = 'YOUR_USERNAME'
# URL encode the password to handle special characters
password = quote_plus('YOUR_PASSWORD')
hostname = 'your-server-name.database.windows.net'
database_name = 'YourDatabaseName'
# URL encode the driver name
driver = quote_plus('ODBC Driver 17 for SQL Server')

# Construct the connection string with URL encoding
connection_string = f"mssql+pyodbc://{username}:{password}@{hostname}/{database_name}?driver={driver}"

def find_specific_excel_file(base_dir, file_name):
    """Find a specific Excel file only in the given directory (no subfolders)."""
    full_path = Path(base_dir)
    excel_files = list(full_path.glob(file_name)) # non-recursive
    print(f"Looking in: {full_path}")
    print(f"Files found: {excel_files}")
    return excel_files

def export_to_sql(df, table_name, sql_connection_string):
    """Export DataFrame to a SQL table."""
    engine = create_engine(sql_connection_string)
    df.to_sql(table_name, con=engine, if_exists='replace', index=False)
    print("Data has been exported to the SQL server.")

def format_created_at_column(df, column_name):
    """Format the 'createdAt' column to the desired date format 'DD/MM/YYYY'."""
    df[column_name] = pd.to_datetime(df[column_name], utc=True).dt.strftime('%d/%m/%Y')

# Configuration
# Placeholder for your local/OneDrive directory path
root_directory = r'C:/Users/YourUser/Path/To/Your/Expenses/Python/'

# Find the specific Excel files
supplier_file_name = "df_joined.xlsx"
email_list_file_name = "Pleo Email List.xlsx"

excel_files = find_specific_excel_file(root_directory, supplier_file_name)
email_list_files = find_specific_excel_file(root_directory, email_list_file_name)

if not excel_files:
    print("No Excel files found.")
else:
    try:
        # Load the first found Excel file into a DataFrame, specifically from 'Sheet1'
        combined_df = pd.read_excel(excel_files[0], sheet_name='Sheet1')

        # Check if 'email', 'firstName', or 'lastName' columns have NULL values
        columns_to_check = ['email', 'firstName', 'lastName']
        columns_with_nulls = [col for col in columns_to_check if col in combined_df.columns and combined_df[col].isnull().any()]

        if columns_with_nulls and email_list_files:
            # Load 'Pleo Email List.xlsx' to lookup missing values
            email_df = pd.read_excel(email_list_files[0], sheet_name='Sheet1')

            # Perform a left join to get missing details based on employeeId
            combined_df = pd.merge(
                combined_df,
                email_df[['employeeId', 'email', 'firstName', 'lastName']],
                on='employeeId',
                how='left',
                suffixes=('', '_email_list')
            )

            # Replace NULLs in 'email', 'firstName', and 'lastName' columns with values from lookup
            for column in columns_with_nulls:
                combined_df[column] = combined_df[column].combine_first(combined_df[f"{column}_email_list"])

            # Drop the helper columns
            combined_df.drop(columns=[f"{column}_email_list" for column in columns_with_nulls], inplace=True)

        # Format the 'createdAt' column
        format_created_at_column(combined_df, 'createdAt')

        # Export the DataFrame to SQL
        export_to_sql(combined_df, 'Pleo_dfjoined', connection_string)
        print("Exported final merged data to SQL.")

except Exception as e:
    print(f"An error occurred: {e}")