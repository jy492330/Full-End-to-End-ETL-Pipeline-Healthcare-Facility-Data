'''
Load only ONE specified sheet into SQL Server
'''
import pyodbc
import pandas as pd


excel_path = r"C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\Data Migration from MS Access to SQL Server\Room Schedule.xlsx"  
# excel_path = 'C:\\Users\\jy492\\OneDrive\\Lenovo Desktop (new)\\WORKSPACE\\Data Engineering Projects\\Data Migration from MS Access to SQL Server\\Room Schedule.xlsx'  # Alternative
df = pd.read_excel(excel_path, sheet_name="Room Schedule")
df = df.astype(str) 

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=JESSICA-DESKTOP;"
    "DATABASE=TestDB;"
    "Trusted_Connection=yes;"
)
cursor = conn.cursor()  

table_name = "Room Schedule" 

cols = ", ".join(f'[{col}] NVARCHAR(MAX)' for col in df.columns)

sql = f'''
IF OBJECT_ID('[{table_name}]', 'U') IS NOT NULL  
DROP TABLE [{table_name}];

CREATE TABLE [{table_name}] ({cols});
'''
cursor.execute(sql)

for _, row in df.iterrows():
    placeholders = ", ".join("?" for _ in row)
    cursor.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", *row)  # parameterized query 

conn.commit()
cursor.close()
conn.close()
print("[✓] Excel sheet imported successfully into SQL Server.")
