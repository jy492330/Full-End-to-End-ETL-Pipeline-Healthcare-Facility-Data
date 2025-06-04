
'''
Fully Automated MS Access → SQL Server (tables) ETL Script:
* Tables go to ABSMC.dbo.[TableName]
* No need to use username or password via Windows Authentication.
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
        if table == "ROOM":
        # Flatten FloorId by joining to FLOOR and BUILDING tables
            df = pd.read_sql("""
                SELECT
                    ROOM.RoomID,
                    ROOM.RoomNumber,
                    ROOM.CostCenterID,
                    ROOM.Activity,
                    ROOM.RoomFunctionNotes,
                    ROOM.Notes,
                    ROOM.Area,
                    ROOM.RoomSort,
                    ROOM.HeadwallCount,
                    ROOM.BedCount,
                    ROOM.BedStatus,
                    ROOM.AirTested,
                    (BUILDING.BuildingCode & '_' & FLOOR.FloorNumber) AS FloorId
                FROM (ROOM INNER JOIN FLOOR ON ROOM.FloorId = FLOOR.FloorId) 
                    INNER JOIN BUILDING ON FLOOR.Building = BUILDING.BuildingId                            
                """, access_conn)
        else:
            df = pd.read_sql(f"SELECT * FROM [{table}]", access_conn)   
        
        df.to_sql(table, sql_engine, if_exists='replace', index=False)
        print(f"[✔] Table '{table}' migrated.")
    except Exception as e:
        print(f"[!] Failed to migrate table '{table}': {e}")


access_cursor.close()
access_conn.close()
print("[✔] ETL process complete.")
