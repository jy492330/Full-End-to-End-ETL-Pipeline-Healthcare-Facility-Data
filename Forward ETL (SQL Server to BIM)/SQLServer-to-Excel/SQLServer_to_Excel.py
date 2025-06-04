# Full end-to-end headless ETL pipeline

import pyodbc
import pandas as pd

conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=JESSICA-DESKTOP;"
    "DATABASE=ABSMC;"
    "Trusted_Connection=yes;"
)

query_map = {
    "ROOM_ALB.xlsx": """
    SELECT * FROM dbo.ROOM
    WHERE FloorID LIKE 'ALB%'
    """,

    "ROOM_HER.xlsx": """
        SELECT * FROM dbo.ROOM
        WHERE FloorId LIKE 'HER%'
    """,

    "ROOM_MER.xlsx": """
        SELECT * FROM dbo.ROOM
        WHERE FloorId LIKE 'MER%'
    """,

    "ROOM_PRO.xlsx": """
        SELECT * FROM dbo.ROOM
        WHERE FloorId LIKE 'PRO%'
    """,

    "ROOM_PER.xlsx": """
        SELECT * FROM dbo.ROOM
        WHERE FloorId LIKE 'PER%'
    """
}
for filename, sql_query in query_map.items():
    df = pd.read_sql(sql_query, conn)
    df.to_excel(filename, index=False)
    print(f"[✔] Exported {len(df)} rows to {filename}")

conn.close()
print("[✔] All exports completed.")
