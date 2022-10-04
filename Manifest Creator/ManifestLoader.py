from openpyxl import load_workbook
from datetime import datetime
from ManifestGenerator import get_manifests_filepath, create_manifests_filepath
import os

class WorkOrder:
    def __str__(self):
        return f'Builder: {self.builder}, Subdivision: {self.subdivision}, Lot: {self.lot}, Windows: {self.windows}, Doors: {self.doors}, Notes: {self.notes}'

    def __init__(self, builder, subdivision, lot, windows, doors, notes):
        if builder is not None:
            self.builder = builder
        else:
            self.builder = ""
        if subdivision is not None:
            self.subdivision = subdivision
        else:
            self.subdivision = ""
        if lot is not None:
            self.lot = lot
        else:
            self.lot = ""
        if windows is not None:
            self.windows = windows
        else:
            self.windows = ""
        if doors is not None:
            self.doors = doors
        else:
            self.doors = ""
        if notes is not None:
            self.notes = notes
        else:
            self.notes = ""
    
    def get_order(self):
        return {
                    "builder": self.builder,
                    "subdivision": self.subdivision,
                    "lot": self.lot,
                    "windows": self.windows,
                    "doors": self.doors,
                    "notes": self.notes
                }

class Manifest:
    def __init__(self, date, lead, crew):
        self.date = date
        self.lead = lead
        self.crew = crew
        self.workorders = []
    
    def add_workorder(self, workorder):
        self.workorders.append(workorder)

def create_new_manifests(givendate):
    manifests = []

def load_manifests(givendate):
    manifests = []
    manifests_filepath = get_manifests_filepath(givendate)

    if not os.path.exists(manifests_filepath):
        create_manifests_filepath(givendate)
        return

    for f in os.listdir(manifests_filepath):
        for file in os.listdir(manifests_filepath + f + "/"): 
            if(file.endswith(givendate + ".xlsx")):
                print("Found manifest file:", file)
                workbook = load_workbook(get_manifests_filepath(givendate) + f + "/" + file)
                sheet = workbook.active

                lead = f.upper()
                crew = ' '.join(sheet["D7"].value[6:].split(", "))

                manifest = Manifest(givendate, lead, crew)              

                for i in range(6):
                    cellrow = str(i + 11)
                    workorder = WorkOrder(
                        sheet["A" + cellrow].value,
                        sheet["B" + cellrow].value,
                        sheet["C" + cellrow].value,
                        sheet["D" + cellrow].value,
                        sheet["E" + cellrow].value,
                        sheet["F" + cellrow].value)
                    manifest.add_workorder(workorder)
                manifests.append(manifest)
    print()
    filter((lambda x: x.lead != ''), manifests)
    return manifests
        
def insert_new_manifest(givendate, manifests):
    manifests = list(filter((lambda x: x.lead != ''), manifests))
    m = Manifest(givendate, "", "")
    for i in range(6):
        workorder = WorkOrder("", "", "", "", "", "")
        m.add_workorder(workorder)
    manifests.insert(0, m)
    return manifests

if __name__ == "__main__":
    import main
