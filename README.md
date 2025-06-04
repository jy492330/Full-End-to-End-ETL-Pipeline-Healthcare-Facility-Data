# Full End-to-End ETL Pipeline for Healthcare Facility Data: Sutter_ABSMC_ETL_Workflow

## --- Summary ---

This is a complete, Python-centric ETL pipeline for automating facility data migration and synchronization workflow between Microsoft Access, SQL Server, Excel, and Autodesk Revit (BIM models). It was developed and optimized during my role as a **Data Migration & Automation Specialist / Project Manager** at CTC – California Technical Contracting, Inc., for five major Sutter Health campuses in the San Francisco Bay Area: Alta Bates, Herrick, Merritt, Peralta, and Summit South Medical Centers.

The solution enables bidirectional data flow:
- **Initial migration between databases** (Access → SQL Server),
- **Forward ETL** (SQL Server → Excel → Revit), and
- **Backward ETL** (Revit → Excel → SQL Server).


This project ensures data accuracy, reduces redundancy, and significantly improves operational reporting and model synchronization for healthcare facilities management.

---

## --- Folder Structure ---

### `Data Migration from MS Access to SQL Server`

Contains scripts and tools to:
- Connect to `.mdb` files via `pyodbc`,
- Extract data and recreate views,
- Apply schema mapping with `SQLAlchemy` ORM,
- Validate, clean and transform records using `Pandas`,
- Back up existing SQL Server databases before insertion.

**Key files and components:**
- `Access_to_SQLServer_ETL(Table+View).py`  
- `SQLServer_BackUp.py`  
- `SQLServer_to_Excel.py`
- `Excel_to_SQLServer_database_v1.py`, `v2.py`  
- `Helper Functions/`  
  - `Convert_Access_WHERE_to_SQLServer.py`  
  - `FindODBCdrivers.py`  
  - `Recreate_Views_Batch_Script.py`  
- `.mdb` + `.ldb` database file: `ABSMC_Database.mdb`

---

### `Forward ETL (SQL Server to BIM)`

This section generates Excel outputs for each medical center (ROOM_ALB, ROOM_HER, etc.) from SQL Server tables using `pyodbc` and `Pandas`, then loads these into BIM using `BIMLink` or `pyRevit`.

**Key components:**
- `SQLServer-to-Excel/`  
  - `SQLServer_to_Excel.py`  
  - Output files for 5 campuses in both `.csv` and `.xlsx` formats
- `Excel-to-BIM/`  
  - Revit models like `Alta Bates_Migration.rvt`, `Herrick_Migration.rvt`, etc.
  - pyRevit buttons like: `ALT_Room.pushbutton/script.py`, `HER_Room.pushbutton/script.py`, etc.
  - Parameter folders + screenshots of successful imports
- `Excel-to-Tableau/`  
  - Foldered by campus for Tableau dashboards (Ashby, Herrick, Merritt, Peralta, and Summit South)

---

### `Backward ETL (BIM to SQL Server)`

This section supports bi-directional sync by extracting parameter values from Revit schedules and pushing them back into the SQL Server database.

**Key tools:**
- Custom `pyRevit` script located in:
  - `pyRevit_Custom_Scripts.extension/ALT.tab/Dev.panel/ALT_Room.pushbutton/script.py`
- Standardized exported schedules from Revit are transformed with `Pandas` and inserted back to SQL using `SQLAlchemy`.

---

## --- Tech Stack ---

| Layer        | Tools & Libraries                                                                 |
|--------------|-------------------------------------------------------------------------------------|
| **Programming**  | Python (3.10+)                                                                  |
| **Database Access** | `pyodbc`, `SQLAlchemy` (Core + ORM), `pandas`                              |
| **Source DB** | Microsoft Access (.mdb), SQL Server                                               |
| **Target Formats** | Excel (`.csv`, `.xlsx`) via `pandas.ExcelWriter`                            |
| **BIM Integration** | Autodesk Revit, `pyRevit`, `BIMLink`                                       |
| **Visualization** | Tableau (organized by campus)                |
| **ETL Automation** | Prefect (for headless scheduling, error logging, monitoring)                |
| **Backup & Versioning** | `SQLServer_BackUp.py`, Git                                              |

---

## --- Results and Impact ---

- Migrated thousands of legacy facility records from Access to SQL Server with schema harmonization.
- Enabled live schedule updates in Revit models with data integrity enforced through custom scripts.
- Allowed facilities teams to sync BIM and database systems effortlessly, reducing manual coordination efforts.
- Set up a robust framework for audit-ready data handling with scheduled Prefect jobs and data backups.
- Enabled visualization and compliance reporting across 5 campuses through structured Excel and Tableau exports.

> **This workflow became a cornerstone of facility operations and digital coordination for Sutter Health campuses.**

---

## --- Notes ---

- This project contains sensitive structure but no PHI or patient data.
- Scripts are modular, reusable with shared parameters across new campuses with minimal config.
- GitHub repo is best viewed with a file/folder tree to track script logic flow.

---

## --- Author ---

Jin Yang (Jin Jessica Yang)  
https://jinjessicayang.com 
https://superb-cucurucho-9279aa.netlify.app
LinkedIn: [@jin-y-30756051](https://www.linkedin.com/in/jin-y-30756051)
