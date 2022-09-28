from openpyxl import load_workbook
from datetime import datetime
from ManifestGenerator import manifests_filepath
import json
import os

date = None

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

def load_spreadsheets(givendate):
    manifests = []

    for f in os.listdir(manifests_filepath):
        for file in os.listdir(manifests_filepath + f + "/"):
            if(file.endswith(givendate + ".xlsx")):
                print("Found file for", givendate, file)
                workbook = load_workbook(manifests_filepath + f + "/" + file)
                sheet = workbook.active
                manifest = Manifest(str(givendate), f.upper(), sheet["D7"].value[6:].split(", "))              

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
    if manifests:
        save_manifests(manifests)
    return manifests

def load_manifests(givendate = False):
    global manifests
    manifests = []
    dates = []

    for f in os.listdir("Manifest Creator/manifests/"):
        d = f.split()[1][:-5]
        dates.append(d)

    latest = max(dates, key = lambda d: datetime.strptime(d, '%m-%d-%Y'))
    if givendate is not False:
        latest = givendate
        if not os.path.exists("Manifest Creator/manifests/manifests " + latest + ".json"):
            manifests = load_spreadsheets(givendate)
            if not manifests:
                return False
            return manifests

    with open('Manifest Creator/manifests/manifests ' + latest + '.json', "r+") as f:
        data = json.load(f)

        for m in data["manifests"]:
            manifest = Manifest(data["date"], m["lead"], m["crew"])

            for w in m["workorders"]:
                workorder = WorkOrder(w["builder"], w["subdivision"], w["lot"], w["windows"], w["doors"], w["notes"])
                manifest.add_workorder(workorder)
            manifests.append(manifest)
    return manifests

def save_manifests(manifests):
    new_manifests = {}
    manifest_date = manifests[0].date
    new_manifests["date"] = manifest_date
    new_manifests["manifests"] = []
    for m in manifests:
        workorders = []
        for w in m.workorders:
            workorders.append(w.get_order())

        new_manifests["manifests"].append({
            "lead" : m.lead,
            "crew" : m.crew,
            "workorders" : workorders 
        })

    with open('Manifest Creator/manifests/manifests ' + manifest_date + ".json", 'w') as jsonFile:
        json.dump(new_manifests, jsonFile)
        


def print_example():
    for m in manifests:
        for w in m.workorders:
            print(m.lead, w)

    #json.dump(data, open("manifests/manifests 09-20-2022.json", "w"), indent = 4)

if __name__ == "__main__":
    #load_manifests()
    #save_manifests()
    print("Please run main.py")
