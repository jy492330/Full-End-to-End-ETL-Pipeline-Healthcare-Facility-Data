import pyodbc

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\jy492\OneDrive\Lenovo Desktop (new)\WORKSPACE\Data Engineering Projects\Data Migration from MS Access to SQL Server\ABSMC_Database.mdb;'
)

conn = pyodbc.connect(conn_str)
print("[✔] Connected to Access database")
conn.close()
