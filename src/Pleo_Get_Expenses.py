import requests
import json
import pandas as pd
from datetime import datetime

# API Keys
ST_Key = 'eyJh...' # Truncated for security
API_Key_PB = 'eyJh...' # Truncated for security

# --- Get Company Accounts ---
url = "https://openapi.pleo.io/v1/accounts?pageSize=1000"
headers = {
    "Authorization": f'Bearer {API_Key_PB}',
    "Content-Type": "application/json"
}

response = requests.get(url, headers=headers)
data = response.json()

Accounts = pd.DataFrame(data['accounts'])
Accounts = Accounts.rename(columns={'id': 'accountId', 'name': 'Account Name'})
# Accounts.head()

# --- Get Expenses ---
base_url = 'https://openapi.pleo.io/v1/expenses'
start_date = '2023-10-01T00:00:00.000Z'
end_date = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'

expenses_list = []
next_page_offset = None
first_iteration = True

while True:
    # Build URL with pagination
    if next_page_offset is not None:
        url = f"{base_url}?dateFrom={start_date}&dateTo={end_date}&pageSize=2000&pageOffset={next_page_offset}"
    else:
        url = f"{base_url}?dateFrom={start_date}&dateTo={end_date}&pageSize=2000"

    headers = {
        "Authorization": f'Bearer {API_Key_PB}',
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        raise Exception(f"Failed to fetch data: {response.status_code}, {response.text}")

    data = response.json()
    expenses_list.extend(data['expenses'])

    # Check for next page
    next_page_offset = data['metadata']['pageInfo'].get('nextPageOffset', None)
    
    # Break loop if no more pages
    if next_page_offset is None and not first_iteration:
        break
    
    first_iteration = False

# Convert to DataFrame
df_expenses = pd.DataFrame(expenses_list)

# --- Flatten Nested JSON Columns ---

# 1. Normalize cardTransaction
df_card_transaction = pd.json_normalize(df_expenses['cardTransaction'])
df_expenses = pd.concat([df_expenses, df_card_transaction], axis=1)
df_expenses = df_expenses.drop(columns=['cardTransaction'])

# 2. Normalize amountOriginal
df_amount_original = pd.json_normalize(df_expenses['amountOriginal'])
df_expenses = pd.concat([df_expenses, df_amount_original], axis=1)
df_expenses = df_expenses.drop(columns=['amountOriginal'])

# 3. Normalize lines
df_lines = pd.json_normalize(df_expenses['lines'])
df_expenses = pd.concat([df_expenses, df_lines], axis=1)
df_expenses = df_expenses.drop(columns=['lines'])

df_expenses.to_excel('Expenses.xlsx')

# --- Get Employees ---
base_url_emp = 'https://openapi.pleo.io/v1/employees'
employees_list = []
offset = 0
page_size = 1000

while True:
    url_emp = f"{base_url_emp}?pageOffset={offset}&pageSize={page_size}"
    headers_emp = {
        "accept": "application/json",
        "authorization": f"Bearer {API_Key_PB}"
    }

    response = requests.get(url_emp, headers=headers_emp)

    if response.status_code != 200:
        raise Exception(f"Failed to fetch data: {response.status_code}, {response.text}")

    data = response.json()
    employees_list.extend(data['employees'])
    
    # Update offset and check if we should continue
    offset += len(data['employees'])
    if len(data['employees']) < page_size:
        break

df_employees = pd.DataFrame(employees_list)
df_employees.to_excel('Employees.xlsx')

# --- Final Merge ---
# Merge Expenses with Employees
df_joined = pd.merge(df_expenses, df_employees, how='left', left_on='employeeId', right_on='id')

# Merge with Accounts
df_joined = pd.merge(df_joined, Accounts, how='left', on='accountId')

df_joined.to_excel('df_joined.xlsx')