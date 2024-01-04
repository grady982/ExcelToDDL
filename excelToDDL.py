import sys
sys.path
import pandas as pd

# Read the Excel file into a pandas DataFrame
excel_file_path = r'C:\Users\grady.liu\Desktop\RegIssueAssetsTemplate_20240102174525.xlsx'
df = pd.read_excel(excel_file_path)

# Create MSSQL DDL statements
ddl_statements = []

for index, row in df.iterrows():
    ddl_statement = f"CREATE TABLE {row['schema']}.{row['table name']} (\n"
    ddl_statement += f"    {row['column name']} {row['type']} {'NOT NULL' if row['是否為PK'] == 'Y' else 'NULL'},\n"
    ddl_statement += f"    CONSTRAINT PK_{row['table name']}_{row['column name']} PRIMARY KEY ({row['column name']})"
    ddl_statement += "\n);\n"
    ddl_statements.append(ddl_statement)

# Print the generated MSSQL DDL statements
for statement in ddl_statements:
    print(statement)
