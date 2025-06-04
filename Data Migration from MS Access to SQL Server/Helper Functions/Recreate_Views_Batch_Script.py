
'''
Batch Export/Recreate View SQLs from MS Access to SQL Server
'''

from sqlalchemy import create_engine, text


SQL_DRIVER = "ODBC Driver 17 for SQL Server"
SQL_SERVER = "JESSICA-DESKTOP"
SQL_DB = "ABSMC"

engine_url = (
    f"mssql+pyodbc://{SQL_SERVER}/{SQL_DB}"
    f"?Driver={SQL_DRIVER.replace(' ', '+')}"
    "&Trusted_Connections=yes"
)
sql_engine = create_engine(engine_url, fast_executemany=True)

view_sqls = {
    "q_RESTROOMS_all": """
        CREATE VIEW dbo.q_RESTROOMS_all AS
        SELECT FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, COST_CENTERS.CostCenterName, ROOM.Activity, ROOM.RoomFunctionNotes, ROOM.Area, ROOM.RoomSort, CAMPUS.CampusName, BUILDING.BuildingCode, BUILDING.BuildingName, BUILDING.Addr1, BUILDING.Addr2, BUILDING.City, BUILDING.FacilityNumber
        FROM (CAMPUS INNER JOIN (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) ON CAMPUS.CampusID = BUILDING.CampusId) INNER JOIN (COST_CENTERS INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE (
        ROOM.Activity LIKE '%toil%' Or
        ROOM.Activity LIKE '%restroom%' Or
        ROOM.Activity LIKE 'restrm%' Or 
        ROOM.Activity LIKE '%locker%' Or 
        ROOM.Activity LIKE '%shower%' Or 
        ROOM.Activity LIKE '%tub%' Or 
        ROOM.Activity LIKE '%bath%'
        );
    """,

    "q_ROOM_Ashby": """ 
        CREATE VIEW dbo.q_ROOM_Ashby AS
        SELECT ROOM.RoomId, BUILDING.BuildingCode, FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, ROOM.Activity, ROOM.Area, CAD_COLORS.CadColor, ROOM.RoomSort
        FROM (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) INNER JOIN ((CAD_COLORS INNER JOIN COST_CENTERS ON CAD_COLORS.CadColorId = COST_CENTERS.CADColorId) INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE BUILDING.BuildingCode='ALB';
    """,

    "q_ROOM_Herrick": """ 
        CREATE VIEW dbo.q_ROOM_Herrick AS
        SELECT ROOM.RoomId, BUILDING.BuildingCode, FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, ROOM.Activity, ROOM.Area, COST_CENTERS.CADColorId, ROOM.RoomSort
        FROM (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) INNER JOIN ((CAD_COLORS INNER JOIN COST_CENTERS ON CAD_COLORS.CadColorId = COST_CENTERS.CADColorId) INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE BUILDING.BuildingCode='HER';
    """,

    "q_ROOM_Merritt": """ 
        CREATE VIEW dbo.q_ROOM_Merritt AS
        SELECT ROOM.RoomId, BUILDING.BuildingCode, FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, ROOM.Activity, ROOM.Area, COST_CENTERS.CADColorId, ROOM.RoomSort
        FROM (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) INNER JOIN ((CAD_COLORS INNER JOIN COST_CENTERS ON CAD_COLORS.CadColorId = COST_CENTERS.CADColorId) INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE BUILDING.BuildingCode='MER';
    """,

    "q_ROOM_Peralta": """ 
        CREATE VIEW dbo.q_ROOM_Peralta AS
        SELECT ROOM.RoomId, BUILDING.BuildingCode, FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, ROOM.Activity, ROOM.Area, CAD_COLORS.CadColor, ROOM.RoomSort
        FROM (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) INNER JOIN ((CAD_COLORS INNER JOIN COST_CENTERS ON CAD_COLORS.CadColorId = COST_CENTERS.CADColorId) INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE BUILDING.BuildingCode='PER';
    """,

    "q_ROOM_Prov": """ 
        CREATE VIEW dbo.q_ROOM_Prov AS
        SELECT ROOM.RoomId, BUILDING.BuildingCode, FLOOR.FloorNumber, ROOM.RoomNumber, COST_CENTERS.CostCenterCode, ROOM.Activity, ROOM.Area, CAD_COLORS.CadColor, ROOM.RoomSort
        FROM (BUILDING INNER JOIN FLOOR ON BUILDING.BuildingId = FLOOR.Building) INNER JOIN ((CAD_COLORS INNER JOIN COST_CENTERS ON CAD_COLORS.CadColorId = COST_CENTERS.CADColorId) INNER JOIN ROOM ON COST_CENTERS.CostCenterId = ROOM.CostCenterId) ON FLOOR.FloorId = ROOM.FloorId
        WHERE BUILDING.BuildingCode='PRO';

    """

}  

with sql_engine.begin() as conn:   
    for view_name, sql in view_sqls.items():
        try:
            conn.execute(text(f"DROP VIEW IF EXISTS dbo.{view_name};"))  
            conn.execute(text(sql))
            print(f"[✔] View '{view_name}' recreated.")
        except Exception as e:
            print(f"[!] Failed to create view '{view_name}': {e}")

print("[✔] View recreation process complete.")
