# Full End-to-End ETL Pipeline for Healthcare Facility Data: Sutter_ABSMC_ETL_Workflow

## --- Summary ---

This is a fully automated, Python-centric ETL pipeline that replaces traditional Microsoft tools such as SSIS and SSRS with a more modern, flexible, and headless architecture. Developed during my tenure at CTC – California Technical Contracting, Inc., the system delivers seamless data migration and synchronization between Microsoft Access, SQL Server, Excel, and BIM platform (Revit models). By leveraging Prefect to run Python scripts in automated sequences, the workflow supports fully automated, bidirectional data exchange—enabling SQL-to-BIM and BIM-to-SQL integration across seven major Sutter Health campuses in the San Francisco Bay Area: Alta Bates, Herrick, Merritt, Peralta, Summit South, Peralta MOB, and Summit South MOB.

Bidirectional data flow solution:
- **Initial migration between databases** (Access → SQL Server),
- **Forward ETL** (SQL Server → Excel → Revit), and
- **Backward ETL** (Revit → Excel → SQL Server).


**This ETL framework was purpose-built to outperform traditional SSIS/SSRS-based workflows by leveraging a modern, Python-centric architecture combined with Prefect for fully headless automation, scheduling, deployment and observability.** Unlike SSIS, which is constrained by GUI-based configuration and Microsoft-specific tooling, this solution offers greater flexibility, extensibility, and cloud readiness. It supports seamless bidirectional data integration across MS Access, SQL Server, Excel, and Autodesk Revit (via pyRevit and BIMLink), with robust logging, error handling, version control, and audit readiness. This modern stack enabled agile deployment, modular reuse, and superior control over data flow logic, making it a strategic upgrade for enterprise facility data management.

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
  - Output files for 5+ campuses in both `.csv` and `.xlsx` formats
- `Excel-to-BIM/`  
  - Revit models like `Alta Bates_Migration.rvt`, `Herrick_Migration.rvt`, etc.
  - pyRevit buttons like: `ALT_Room.pushbutton/script.py`, `HER_Room.pushbutton/script.py`, etc.
  - Parameter folders + screenshots of successful imports
- `Excel-to-Tableau/`  
  - Foldered by campus for Tableau dashboards (Ashby, Herrick, Merritt, Peralta, and Summit South)
  - Interactive Tableau dashboards like `Ashby Dashboard_Updated.twb`, `Herrick Dashboard_Updated..twb`, etc. with alternative .html files.

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

- Replaced traditional SSIS packages with scalable Python-based ETL scripts orchestrated via Prefect, enabling fully headless execution, retry logic, and audit-friendly logging.
- Eliminated the need for SSRS by implementing modern reporting solutions using Tableau and Streamlit, delivering dynamic dashboards far beyond SSRS’s static reporting capabilities.
- Migrated thousands of legacy facility records from MS Access to SQL Server with schema harmonization and ORM-based transformations using SQLAlchemy.
- Established a repeatable framework for BIM-data synchronization that eliminated manual intervention and reduced operational redundancy.
- Delivered structured, cross-platform datasets optimized for compliance tracking, executive reporting, and campus-level analytics.

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
