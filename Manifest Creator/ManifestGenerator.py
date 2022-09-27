from openpyxl import load_workbook
import os
from datetime import datetime, date, timedelta

weekdays = {0 : "MONDAY", 1 : "TUESDAY", 2 : "WEDNESDAY", 3 : "THURSDAY", 4 : "FRIDAY", 5 : "SATURDAY", 6 : "SUNDAY"}
manifests_filepath = "//dc01/_SHARE/Company/Forms/Daily Job sheets/DAILY JOB SHEETS - SEPTEMBER 2022/Install Manifests/"
active_date = None
weekday = None
crews = []

def execute():
    if os.path.exists(manifests_filepath):
        print("Manifests filepath:", manifests_filepath)
    else:
        print("Manifests filepath does not exist.")
        exit()

    print("How many days ahead of current date?")
    daysn = 0

    while True:
        try:
            choice = int(input())
        except ValueError:
            print("Invalid input.")
            continue

        if choice > 0:
            daysn = choice
            break
        else:
            print("No files have been created.")
            exit()
    
    current_day = date.today()
    future_date = current_day + timedelta(days = daysn)

    weekdayN = future_date.weekday()
    global weekday
    weekday = weekdays[weekdayN]

    formatted_date = date.strftime(future_date, "%m-%d-%Y")
    global active_date
    active_date = formatted_date

    print("Creating files for", weekday, active_date)

    with open('Manifest Creator/resources/lead_installers.txt') as f:
        lines = f.readlines()
        for line in lines:
            global crews
            crews.append(line.split())

    for crew in crews:
        generate_template(crew[0])
    
    print("\nWould you like to print these spreadsheets? [Y]es [N]o - Active date:", active_date)

    while True:
        choice = input()[0].lower()

        if choice == 'y':
            print_manifests()
            break
        elif choice == 'n':
            print("Printing cancelled.")
            break
        else:
            print("Invalid.")

def create_workbook(manifest):
    template_name = "INSTALLATION DAILY MANIFEST NAME DATE"
    workbook_path = "Manifest Creator/resources/" + template_name + ".xlsx"

    if not os.path.exists(workbook_path):
        print("Workbook template does not exist")
        return False
    
    workbook = load_workbook(workbook_path)
    sheet = workbook.active

    global active_date
    active_date = manifest.date
    
    sheet["A2"] = manifest.date
    sheet["A3"] = "DATE: " + weekdays[datetime.strptime(manifest.date, "%m-%d-%Y").weekday()]
    sheet["D6"] = "LEAD INSTALLER: " + manifest.lead
    sheet["D7"] = "CREW: " + ", ".join(manifest.crew.split())

    for i in range(len(manifest.workorders)):
        cellrow = str(i + 11)
        sheet["A" + cellrow] = manifest.workorders[i].builder
        sheet["B" + cellrow] = manifest.workorders[i].subdivision
        sheet["C" + cellrow] = manifest.workorders[i].lot
        sheet["D" + cellrow] = manifest.workorders[i].windows
        sheet["E" + cellrow] = manifest.workorders[i].doors
        sheet["F" + cellrow] = manifest.workorders[i].notes    
        
    template_name = template_name.split(" ")
    template_name = " ".join(template_name[:-2]) + " " + manifest.lead.upper() + " " + manifest.date + ".xlsx"
    file_path = manifests_filepath + manifest.lead + "/"
    final_path = file_path + template_name

    if not os.path.exists(file_path):
        print("\nDirectory does not exist for:", manifest.lead)
        new_dir = os.path.join(manifests_filepath, manifest.lead)
        os.mkdir(new_dir)
        print("Created directory for:", manifest.lead, "under the following filepath:")
        print(new_dir, "\n")
    
    workbook.save(filename = final_path)

def generate_template(lead_name):
    template_name = "INSTALLATION DAILY MANIFEST NAME DATE"
    workbook_path = "Manifest Creator/resources/" + template_name + ".xlsx"
    if not os.path.exists(workbook_path):
        print("Workbook template does not exist.")
        exit()

    workbook = load_workbook(workbook_path)
    sheet = workbook.active

    sheet["A2"] = active_date
    sheet["A3"] = "DATE: " + weekday
    sheet["D6"] = "LEAD INSTALLER: " + lead_name
    teammates = None

    for crew in crews:
        if lead_name == crew[0]:
            if len(crew) > 1:
                teammates = ", ".join(crew[1:])
                sheet["D7"] = "CREW: " + teammates

    template_name = template_name.split(" ")
    template_name = " ".join(template_name[:-2]) + " " + lead_name.upper() + " " + active_date + ".xlsx"
    file_path = manifests_filepath + lead_name + "/"
    final_path = file_path + template_name

    if not os.path.exists(file_path):
        print("\nDirectory does not exist for:", lead_name)
        new_dir = os.path.join(manifests_filepath, lead_name)
        os.mkdir(new_dir)
        print("Created directory for:", lead_name, "under the following filepath:")
        print(new_dir, "\n")
    
    if not os.path.exists(final_path):
        workbook.save(filename = final_path)
        print(template_name)
    else:
        print("File already exists at:\n", os.path.normpath(final_path))

        

def print_manifests():
    print("Printing manifests...")

    for f in os.listdir(manifests_filepath):
        for files in os.listdir(manifests_filepath + f + "/"):
            if(files.endswith(active_date + ".xlsx")):
                workbook = load_workbook(manifests_filepath + f + "/" + files)
                sheet = workbook.active
                A11 = sheet["A11"].value

                if not A11 == None:
                    print("Printing manifest for:", f)
                    os.startfile(os.path.normpath(manifests_filepath + f + "/" + files), "print")
                else:
                    print("Manifest for", f, "is incomplete. Print Aborted.")

if __name__ == "__main__":
    execute()