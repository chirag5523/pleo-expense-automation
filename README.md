Pleo Expense Automation & SQL Integration

This repository contains a suite of Python scripts designed to automate the extraction of company expenses from the Pleo API, join them with employee metadata, and export the cleaned dataset to an Azure SQL Database.

⚙️ Script Breakdown
1. Pleo_Get_Expenses.py

Purpose: The primary data ingestion engine.

    API Connectivity: Uses requests to pull data from three Pleo endpoints: Accounts, Expenses, and Employees.

    Data Flattening: Pleo returns nested JSON (e.g., cardTransaction, amountOriginal). The script uses pd.json_normalize to flatten these into readable table columns.

    Merging: It performs a left join between the Expenses and Employees datasets using the employeeId as the common key.

    Output: Generates df_joined.xlsx, which serves as the input for the final processing job.

2. Pleo_Emails.py

Purpose: Manages the master contact reference.

    Utility: Locates the Pleo Email List.xlsx file in your local directory.

    Database Sync: Exports this master list directly to the SQL table Pleo_Email_List to ensure the database has an up-to-date record of employee roles and emails.

3. Pleo_Job.py

Purpose: The final transformation and "Gap-Filling" script.

    Data Validation: It checks the df_joined.xlsx for any NULL values in the email, firstName, or lastName columns.

    Lookup Logic: If data is missing, it performs a join with the Pleo Email List.xlsx to fill in the gaps based on the employeeId.

    Formatting: Standardizes the createdAt column to a DD/MM/YYYY format for consistent SQL reporting.

    Export: Uploads the final, perfected dataset to the Pleo_dfjoined table in SQL Server.

🚀 How to Run

    Install Dependencies:
    Bash

    pip install pandas requests sqlalchemy pyodbc openpyxl

    Setup Credentials: Update the username, password, and hostname variables in the scripts (or use environment variables).

    Execution Order:

        Run Pleo_Get_Expenses.py to fetch fresh data from the API.

        Run Pleo_Emails.py to sync your master email list.

        Run Pleo_Job.py to clean the data and push it to your SQL Server.

📊 SQL Table Schema

The final table in your database (Pleo_dfjoined) will contain over 40 columns, including:

    Transaction Info: id_x (Expense ID), amountSettled, supplier, status, createdAt.

    Employee Info: firstName, lastName, email, jobTitle, teamId.

    Account Info: Account Name, accountNumber.

🔒 Security Note

This repository uses a .gitignore file to ensure that local Excel files containing sensitive employee information and scripts containing database passwords are not pushed to public version control. Always use environment variables for production credentials.
