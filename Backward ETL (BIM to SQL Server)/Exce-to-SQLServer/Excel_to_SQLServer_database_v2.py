'''
Load ALL sheets into SQL Server

It includes:

✅ Raw string path handling
✅ Safe dropping/creation of tables
✅ Conversion of all cell values to string (avoids type coercion errors)
✅ Table name cleanup (removing spaces)
'''
import pyodbc
import pandas as pd


# Excel file path (use raw string or double backslashes)
excel_path = r"C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\Data Migration from MS Access to SQL Server\Room Schedule.xlsx"

# Get all sheet names to load all sheets
xls = pd.ExcelFile(excel_path)  # loads the entire file
sheet_names = xls.sheet_names

# Connect to local SQL Server
conn = pyodbc.connect(
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=JESSICA-DESKTOP;"
    "Database=TestDB;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

# Loop through each sheet to define target table names
for sheet in sheet_names:
    df = xls.parse(sheet)  # parsing excel to DataFrame
    df = df.astype(str)    # must convert all values to strings before inserting into a NVARCHAR column (cell value)
    table_name = sheet     # alternative: table_name = sheet.replace(" ", "")

    # Create column definitions 
    cols = ", ".join(f"[{col}] NVARCHAR(MAX)" for col in df.columns)

    # Drop target table if already exists. Otherwise, create it with SQL.
    sql = f"""
    IF OBJECT_ID('[{table_name}]', 'U') IS NOT NULL
    DROP TABLE [{table_name}];

    CREATE TABLE [{table_name}] ({cols});
    """
    cursor.execute(sql)

    # Inserts each row as placeholders using parameterized query technique
    for index, row in df.iterrows():
        placeholders = ", ".join("?" for _ in row)
        cursor.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", *row)  # parameterized query technique


# Commit and close
conn.commit()
cursor.close()
conn.close()
print("[✓] All sheets imported successfully into SQL Server.")


# run this script in the terminal with command: python Excel_to_SQLServer_database_v2.py
