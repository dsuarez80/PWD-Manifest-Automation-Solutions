from datetime import datetime, date
from tkinter import *
import tkinter as tk
from tkinter import ttk, messagebox
from ManifestGenerator import print_manifests, create_workbook, print_manifest
from ManifestLoader import load_manifests, insert_new_manifest

root = tk.Tk() #creating the main window and storing the window object in 'win'
root.title('Install Manifests') #setting title of the window
root.geometry("1400x900") #setting the size of the window
root.resizable(False, False) #making neither the x nor y resizable for the window
header = Frame(root, pady = 10, padx = 20)
header.pack(fill = BOTH)

manifests = load_manifests(datetime.strftime(date.today(), "%m-%d-%Y"))

Label(header, text = "Date: ").pack(side = LEFT)
date_entry = Entry(header)
date_entry.pack(side = LEFT)
date_entry.insert(0, manifests[0].date)

saved = True

def retrieve_manifests(manifests = manifests):
    try:
        d = date_entry.get()
        date_entry.delete(0, END)
        date_entry.insert(0, date.strftime(datetime.strptime(d, "%m-%d-%Y"), "%m-%d-%Y"))
    except:
        messagebox.showinfo("Information", "Date is invalid.")
        date_entry.insert(0, manifests[0].date)
        return

    global saved
    if not saved:
        if not messagebox.askyesno(title = "Load Manifests", message = "Changes not saved.\n\nAttempt to load new manifests anyway?"):
            return

    manifests = load_manifests(date_entry.get())
    saved = True

    if not manifests:
        create = messagebox.askyesno(title = "Load Manifests", message = "There are no manifest files for " + date_entry.get() + ". Would you like to create them now?")
        if create:
            update_manifests(True)
            return
        else:
            messagebox.showinfo(title = "Load Manifests", message = "Manifests and spreadsheets have not been created for " + date_entry.get() + ".")
            return

    main_frame.destroy()
    initialize(root, manifests)


def populate_manifest(win, manifests):
    #lists of each entry widget
    global leads
    leads = []
    global crews
    crews = []
    global builders
    builders = []
    global subdivisions
    subdivisions = []
    global lots
    lots = []
    global windowsL
    windowsL = []
    global doorsL
    doorsL = []
    global notesL
    notesL = []
    #^these are widgets, not the values stored in manifests object list, though they do contain the values of the data in the entry widget

    global v
    v = []
    manifestno = 1
    i = 2

    def not_saved(*args):
        global saved
        saved = False

    for m in manifests:
        frame = LabelFrame(win, text = 'Manifest No. ' + str(manifestno) + " for " + m.date, padx = 20, pady = 20)
        frame.pack(fill = BOTH, expand = YES)
        
        Label(frame, text = 'Lead').grid(row = i, column = 1) 
        Label(frame, text = 'Crew').grid(row = i, column = 2) 
        i += 1

        lev = tk.StringVar()
        lead = Entry(frame, textvariable = lev)
        lead.grid(row = i, column = 1, ipadx = 60)
        lead.insert(0, m.lead)
        v.append(lev)
        leads.append(lead)

        cv = tk.StringVar()
        crew = Entry(frame, textvariable = cv)
        crew.grid(row = i, column = 2, ipadx = 60)
        crew.insert(0, m.crew)
        v.append(cv)
        crews.append(crew)

        def erase_manifest(x = manifestno):
            for i in range((x-1)*6, (x)*6):
                builders[i].delete(0, END)
                subdivisions[i].delete(0, END)
                lots[i].delete(0, END)
                windowsL[i].delete(0, END)
                doorsL[i].delete(0, END)
                notesL[i].delete(0, END)
            
                builders[i].insert(0, "")
                subdivisions[i].insert(0, "")
                lots[i].insert(0, "")
                windowsL[i].insert(0, "")
                doorsL[i].insert(0, "")
                notesL[i].insert(0, "")
            not_saved()

        def print_manifest_button(x = lead.get()):
            try:
                d = date_entry.get()
                date_entry.delete(0, END)
                date_entry.insert(0, date.strftime(datetime.strptime(d, "%m-%d-%Y"), "%m-%d-%Y"))
                d = date_entry.get()
                print_manifest(x, d)
            except:
                messagebox.showinfo("Information", "Date is invalid.")
                return
        printbtn = Button(frame, text = "Print Manifest", command = print_manifest_button)
        printbtn.grid(row = i, column = 0, padx = 5)
        erasebtn = Button(frame, text = "Erase Entries", command = erase_manifest)
        erasebtn.grid(row = i + 2, column = 0, padx = 5)
        i += 1

        Label(frame, text = 'Builder').grid(row = i, column = 1)
        Label(frame, text = 'Subdivision').grid(row = i, column = 2) 
        Label(frame, text = 'Lot').grid(row = i, column = 3)
        Label(frame, text = 'Windows').grid(row = i, column = 4) 
        Label(frame, text = 'Doors').grid(row = i, column = 5) 
        Label(frame, text = 'Notes').grid(row = i, column = 6)
        i += 1

        for workorder in m.workorders:
            bv = tk.StringVar()
            builder = Entry(frame, textvariable = bv)
            builder.grid(row = i, column = 1, ipadx = 60, ipady = 3)
            builder.insert(0, workorder.builder)
            v.append(bv)
            builders.append(builder)

            sv = tk.StringVar()
            subdivision = Entry(frame, textvariable = sv)
            subdivision.grid(row = i, column = 2, ipadx = 60, ipady = 3)
            subdivision.insert(0, workorder.subdivision)
            v.append(sv)
            subdivisions.append(subdivision)
            
            lv = tk.StringVar()
            lot = Entry(frame, textvariable = lv)
            lot.grid(row = i, column = 3, ipady = 3)
            lot.insert(0, workorder.lot)
            v.append(lv)
            lots.append(lot)

            wv = tk.StringVar()
            windows = Entry(frame, textvariable = wv)
            windows.grid(row = i, column = 4, ipadx = 30, ipady = 3)
            windows.insert(0, workorder.windows)
            v.append(wv)
            windowsL.append(windows)
      
            dv = tk.StringVar()
            doors = Entry(frame, textvariable = dv)
            doors.grid(row = i, column = 5, ipadx = 30, ipady = 3)
            doors.insert(0, workorder.doors)
            v.append(dv)
            doorsL.append(doors)

            nv = tk.StringVar()
            notes = Entry(frame, textvariable = nv)
            notes.grid(row = i, column = 6, ipadx = 60, ipady = 3)
            notes.insert(0, workorder.notes)
            v.append(nv)
            notesL.append(notes)

            i += 1
        manifestno += 1
    for x in v:
        x.trace("w", not_saved)

def update_manifests(new = False, m = manifests):
    # index for looping through work order entries. 
    # example: each manifest has multiple work orders, if manifest 1 has 2 orders, and manifest 2 has 4, 
    # in order to find the 2nd order for the 2nd manifest, that would be index 4 in this case, index 4 for each workorder entry widget
    index = 0

    try:
        d = date_entry.get()
        date_entry.delete(0, END)
        date_entry.insert(0, date.strftime(datetime.strptime(d, "%m-%d-%Y"), "%m-%d-%Y"))
    except:
        messagebox.showinfo("Information", "Date is invalid.")
        return

    for x in range(len(m)):
        m[x].date = date_entry.get()
        m[x].lead = leads[x].get()
        m[x].crew = crews[x].get()

        for y in range(len(m[x].workorders)):
            m[x].workorders[y].builder = builders[index].get()
            m[x].workorders[y].subdivision = subdivisions[index].get()
            m[x].workorders[y].lot = lots[index].get()
            m[x].workorders[y].windows = windowsL[index].get()
            m[x].workorders[y].doors = doorsL[index].get()
            m[x].workorders[y].notes = notesL[index].get()
            index += 1
        
        if m[x].lead:
            create_workbook(m[x])

    if new:
        messagebox.showinfo("Information", "New manifests and spreadsheets have been created for: " + d)
        retrieve_manifests()
    else:
        messagebox.showinfo("Information", "Manifests saved and spreadsheets have been created.")
    
    global saved
    saved = True

def fake_print():
    global saved
    if not saved:
        if not messagebox.askyesno("Information", "Changes not saved.\n\nPrint anyway?"):
            return
    try:
        d = date_entry.get()
        date_entry.delete(0, END)
        date_entry.insert(0, date.strftime(datetime.strptime(d, "%m-%d-%Y"), "%m-%d-%Y"))
    except:
        messagebox.showinfo("Information", "Date is invalid.")
        return
    
    print("printing hehe", date_entry.get())

def print_manifests_button():
    global saved
    if not saved:
        if not messagebox.askyesno("Information", "Changes not saved.\n\nPrint anyway?"):
            return
    try:
        d = date_entry.get()
        date_entry.delete(0, END)
        date_entry.insert(0, date.strftime(datetime.strptime(d, "%m-%d-%Y"), "%m-%d-%Y"))
    except:
        messagebox.showinfo("Information", "Date is invalid.")
        return

    print_manifests(date_entry.get())

def add_manifest():
    m = insert_new_manifest(date_entry.get(), manifests)
    main_frame.destroy()
    initialize(root, m)

#initializes the main frame which houses all manifest entries and scroll bar
def initialize(root, manifests):
    global main_frame
    main_frame = Frame(root)
    main_frame.pack(fill = BOTH, expand = 1)
    my_canvas = Canvas(main_frame)
    my_canvas.pack(side = LEFT, fill = BOTH, expand = 1)
    my_scrollbar = ttk.Scrollbar(main_frame, orient = VERTICAL, command = my_canvas.yview)
    my_scrollbar.pack(side = RIGHT, fill = Y)
    my_canvas.configure(yscrollcommand = my_scrollbar.set)
    my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion = my_canvas.bbox("all")))
    root.bind_all('<MouseWheel>', lambda event: my_canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
    second_frame = Frame(my_canvas, padx = 20)
    my_canvas.create_window((0,0), window = second_frame, anchor = "nw")

    populate_manifest(second_frame, manifests)

#buttons
btn = Button(header, text="Load Manifests", padx = 20, command=retrieve_manifests)
btn.pack(side = LEFT)
date_entry.bind('<Return>', retrieve_manifests)

btn2 = Button(header, text="Print Manifests", padx = 20, command=print_manifests_button)
btn2.pack(side = RIGHT)
btn3 = Button(header, text="Update Manifests", padx = 20, command=update_manifests)
btn3.pack(side = RIGHT)

menu_canvas = Canvas(root, width = 600, height = 60)
menu_canvas.pack(side = TOP, fill = X)
menu_border_frame = Frame(menu_canvas, padx = 20)
menu_canvas.create_window((0,0), window = menu_border_frame, anchor = "nw")
menu_frame = LabelFrame(menu_border_frame, text = "Main Menu", pady = 5, padx = 10)
menu_frame.pack(side = LEFT, fill = BOTH)

btn4 = Button(menu_frame, text="Add New Manifest", padx = 20, command=add_manifest)
btn4.pack(side = LEFT)

initialize(root, manifests)

def ask_quit():
    if saved:
        root.destroy()
    elif messagebox.askyesno("Quit", "Changes not saved.\n\nAre you sure you want to quit?"):
        root.destroy()

root.protocol("WM_DELETE_WINDOW", ask_quit)
root.mainloop() #running the loop that works as a trigger