
'''
✅ Fully Automated MS Access → SQL Server (tables + views) ETL Script:
- Tables go to ABSMC.dbo.[TableName]
- Views recreated as T-SQL views under ABSMC.Views
- No need to use username or password via Windows Authentication.
'''

import pyodbc
import pandas as pd
from sqlalchemy import create_engine, text


ACCESS_DB_PATH = r"C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\ABSMC_Database.mdb"
ACCESS_DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
SQL_DRIVER = "ODBC Driver 17 for SQL Server"
SQL_SERVER = "JESSICA-DESKTOP"
SQL_DB = "ABSMC"

access_conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\Data Migration from MS Access to SQL Server\ABSMC_Database.mdb;'
)
access_conn = pyodbc.connect(access_conn_str)
access_cursor = access_conn.cursor()

tables = [row.table_name for row in access_cursor.tables(tableType="TABLE")]   
views = [row.table_name for row in access_cursor.tables(tableType="VIEW")]    

engine_url = (
    f"mssql+pyodbc://{SQL_SERVER}/{SQL_DB}"  
    f"?driver={SQL_DRIVER.replace(' ', '+')}"  
    "&trusted_connection=yes"                 
)
sql_engine = create_engine(engine_url, fast_executemany=True) 

# --- MIGRATE TABLES ---
for table in tables:
    try:
        df = pd.read_sql(f"SELECT * FROM [{table}]", access_conn)
        df.to_sql(table, sql_engine, if_exists='replace', index=False)
        print(f"[✔] Table '{table}' migrated.")
    except Exception as e:
        print(f"[!] Failed to migrate table '{table}': {e}")


# --- MIGRATE VIEWS ---
# for view in views:
#     try:
#         access_cursor.execute(f"""
#             SELECT MSysObjects.Name, MSysQueries.Sql 
#             FROM MSysObjects 
#             INNER JOIN MSysQueries ON MSysObjects.Id = MSysQueries.ObjectId 
#             WHERE MSysObjects.Name = '{view}'
#         """)
#         row = access_cursor.fetchone()
#         if not row:
#             print(f"[!] View '{view}' definition not found.")
#             continue

#         sql = row.Sql
#         with sql_engine.begin() as conn:
#             conn.execute(text(f"IF OBJECT_ID('{view}', 'V') IS NOT NULL DROP VIEW [{view}];"))
#             conn.execute(text(f"CREATE VIEW [{view}] AS {sql}"))
#         print(f"[✔] View '{view}' recreated.")
#     except Exception as e:
#         print(f"[!] Failed to recreate view '{view}': {e}")


access_cursor.close()
access_conn.close()
print("[✔] ETL process complete.")
