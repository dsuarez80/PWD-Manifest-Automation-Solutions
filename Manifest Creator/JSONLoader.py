from datetime import datetime
import json
import os

date = None

class WorkOrder:
    def __str__(self):
        return f'Builder: {self.builder}, Subdivision: {self.subdivision}, Lot: {self.lot}, Windows: {self.windows}, Doors: {self.doors}, Notes: {self.notes}'

    def __init__(self, builder, subdivision, lot, windows, doors, notes):
        self.builder = builder
        self.subdivision = subdivision
        self.lot = lot
        self.windows = windows
        self.doors = doors
        self.notes = notes
    
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

def load_manifests(givendate):
    dates = []

    for f in os.listdir("Manifest Creator/manifests/"):
        d = f.split()[1][:-5]
        dates.append(d)

    latest = max(dates, key = lambda d: datetime.strptime(d, '%m-%d-%Y'))

    if bool(givendate):
        latest = datetime.strptime(givendate, '%m-%d-%Y')
        print(latest)
        if not os.path.exists("Manifest Creator/manifests/manifests " + latest + ".json"):
            return False

    with open('Manifest Creator/manifests/manifests ' + latest + '.json', "r+") as f:
        data = json.load(f)

        global manifests
        manifests = []

        for m in data["manifests"]:
            manifest = Manifest(data["date"], m["lead"], m["crew"])

            for w in m["workorders"]:
                workorder = WorkOrder(w["builder"], w["subdivision"], w["lot"], w["windows"], w["doors"], w["notes"])
                manifest.add_workorder(workorder)
            manifests.append(manifest)
    return manifests

def save_manifests():
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
    load_manifests(False)
    save_manifests()
