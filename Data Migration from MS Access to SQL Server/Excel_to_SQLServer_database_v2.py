'''
Load ALL sheets into SQL Server

It includes:
* Raw string path handling
* Safe dropping/creation of tables
* Conversion of all cell values to string (avoids type coercion errors)
* Table name cleanup (removing spaces)
'''
import pyodbc
import pandas as pd

excel_path = r"C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\Data Migration from MS Access to SQL Server\Room Schedule.xlsx"

xls = pd.ExcelFile(excel_path)  
sheet_names = xls.sheet_names

conn = pyodbc.connect(
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=JESSICA-DESKTOP;"
    "Database=TestDB;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()

for sheet in sheet_names:
    df = xls.parse(sheet)  
    df = df.astype(str)    
    table_name = sheet     

    cols = ", ".join(f"[{col}] NVARCHAR(MAX)" for col in df.columns)

    sql = f"""
    IF OBJECT_ID('[{table_name}]', 'U') IS NOT NULL
    DROP TABLE [{table_name}];

    CREATE TABLE [{table_name}] ({cols});
    """
    cursor.execute(sql)

    for index, row in df.iterrows():
        placeholders = ", ".join("?" for _ in row)
        cursor.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", *row) 

conn.commit()
cursor.close()
conn.close()
print("[✓] All sheets imported successfully into SQL Server.")
