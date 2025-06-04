# === coding: utf-8 ===
import clr
import csv
import os

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import *
from pyrevit import revit, script


if __name__ == '__main__':
    output = script.get_output()
    doc = revit.doc

    if not doc:
        output.print_md("**[ERROR] No active Revit document. Please open a project.**")
        raise Exception("No Revit document")


    csv_path = (r"C:\Users\jy492\OneDrive\Lenovo Desktop (new)\NucampFolder\Backend Python\
                GitHub Projects\Sutter ABSMC ETL Workflow\Forward ETL\SQLServer-to-Excel\ROOM_HER.csv")
    if not os.path.exists(csv_path):
        output.print_md("**[ERROR] CSV file not found at path:** `" + csv_path + "`")
        raise Exception("Missing CSV")


    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)
        data = [row for row in reader]


    level_name = "Level LL"
    level = None
    for lvl in FilteredElementCollector(doc).OfClass(Level):
        if lvl.Name == level_name:
            level = lvl
            break

    if not level:
        with Transaction(doc, "Create Level LL") as t:
            t.Start()
            level = Level.Create(doc, 0)
            level.Name = level_name
            t.Commit()


    with Transaction(doc, "Import Unplaced Rooms from CSV") as t:
        t.Start()
        for row in data:
            room_number = row.get("RoomNumber", "").strip()
            room_name = row.get("Activity", "").strip()
            area_val = row.get("Area", "")

            room = doc.Create.NewRoom(level, UV(0, 0))

            if room.LookupParameter("Number"):
                room.LookupParameter("Number").Set(room_number)
            if room.LookupParameter("Name"):
                room.LookupParameter("Name").Set(room_name)
            if room.LookupParameter("Area") and not room.LookupParameter("Area").IsReadOnly:
                try:
                    room.LookupParameter("Area").Set(float(area_val))
                except:
                    pass

            for key, val in row.items():
                if key in ["RoomNumber", "Activity", "Area"]:
                    continue
                param = room.LookupParameter(key)
                if param and not param.IsReadOnly:
                    try:
                        param.Set(str(val))
                    except:
                        pass
        t.Commit()

    output.print_md("**[OK] ROOM_HER.csv successfully imported as unplaced rooms.**")