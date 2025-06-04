'''
✅ Script 1: SQL Server Backup
'''

import os
import pyodbc
from datetime import datetime


# --- CONFIGURATION ---
SQL_Server = "JESSICA-DESKTOP"
SQL_Driver = "{ODBC Driver 17 for SQL Server}"
SQL_DB = "ABSMC"

# --- CREATE BACKUP OUTPUT FOLDER ---
backup_dir = r"C:\SQLBackups\ABSMC_backup"   
os.makedirs(backup_dir, exist_ok=True)

# --- CREATE UNIQUE TIMESTAMPED BACKUP FILE NAME AND FILE PATH ---
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  
backup_file_path = os.path.join(backup_dir, f"{SQL_DB}_backup_{timestamp}.bak")

# --- CONNECT TO SQL SERVER ---
print(f"Connecting to SQL Server at: {SQL_Server}")
conn = pyodbc.connect(
    f"Driver={SQL_Driver};"
    f"Server={SQL_Server};"
    f"Database={SQL_DB};"
    "Trusted_Connection=yes;",
    autocommit=True    
)
cursor = conn.cursor()

# --- BACKUP EXECUTION ---
print(f"Backing up database '{SQL_DB}' to:\n {backup_file_path}")
sql = f"""
BACKUP DATABASE [{SQL_DB}] TO DISK = N'{backup_file_path}'
WITH FORMAT, NAME = N'{SQL_DB} Full Backup';  
"""
cursor.execute(sql)   # implicit transaction


cursor.commit()
cursor.close()
conn.close()

print("[✔] Backup completed successfully.")
print("[✔] File created:", os.path.exists(backup_file_path), "\n", backup_file_path)
