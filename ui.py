from tkinter import *
from tkinter import ttk

root = Tk()
root.title("GUI")


def prefix_count2(prefix, index):
    if index < 10:
        return prefix + "00" + str(index)
    if 10 <= index < 100:
        return prefix + "0" + str(index)
    if 100 <= index < 1000:
        return prefix + str(index)


def simulate():
    prefix_with_zeros = prefix_count2(prefix.get(), count_min.get())
    preview = prefix_with_zeros + '-' + name.get() + '-' + year.get() \
              + '-' + season.get() + "-PA-RS"
    preview_var.set(preview)


mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# ttk.Button(mainframe, text="Create", command="").grid(column=3, row=3, sticky=W)
# ROW 1 prefix
ttk.Label(mainframe, text="prefix").grid(column=1, row=1)
prefix = StringVar()
prefix_entry = ttk.Entry(mainframe, width=3, textvariable=prefix)
prefix_entry.grid(column=2, row=1)
ttk.Label(mainframe, text="Ex : Af").grid(column=3, row=1)
ttk.Label(mainframe, textvariable=prefix).grid(column=4, row=1)

# ROW 2 COUNT
ttk.Label(mainframe, text="Borne Min Max").grid(column=1, row=2)
count_min = IntVar()
count_min_entry = ttk.Entry(mainframe, width=3, textvariable=count_min)
count_min_entry.grid(column=2, row=2)

count_max = IntVar()
count_max_entry = ttk.Entry(mainframe, width=3, textvariable=count_max)
count_max_entry.grid(column=3, row=2)

ttk.Label(mainframe, text="Ex : >= 0").grid(column=4, row=2)
show_count = StringVar(value=str(count_min.get()) + ' to ' + str(count_max.get()))
ttk.Label(mainframe, textvariable=count_min).grid(column=5, row=2)
ttk.Label(mainframe, textvariable=count_max).grid(column=6, row=2)

# ROW 3 NAME

ttk.Label(mainframe, text="Name").grid(column=1, row=3)
name = StringVar()
name_entry = ttk.Entry(mainframe, width=3, textvariable=name)
name_entry.grid(column=2, row=3)
ttk.Label(mainframe, text="Ex : Bn").grid(column=3, row=3)
ttk.Label(mainframe, textvariable=name).grid(column=4, row=3)

# ROW 4 YEAR

ttk.Label(mainframe, text="Year").grid(column=1, row=4)
year = StringVar()
year_entry = ttk.Entry(mainframe, width=3, textvariable=year)
year_entry.grid(column=2, row=4)
ttk.Label(mainframe, text="Ex : Y21").grid(column=3, row=4)
ttk.Label(mainframe, textvariable=year).grid(column=4, row=4)

# ROW 5 SEASON (Table ?)

ttk.Label(mainframe, text="Season").grid(column=1, row=5)
season = StringVar()
season_entry = ttk.Entry(mainframe, width=3, textvariable=season)
season_entry.grid(column=2, row=5)
ttk.Label(mainframe, text="Ex : Au").grid(column=3, row=5)
ttk.Label(mainframe, textvariable=season).grid(column=4, row=5)

# ROW 6 tab1 (faire attention pour "Culturomique" -> Par défaut en plus? si tjs utilisé?)

# ROW 7 tab2

# Row 8 Preview
ttk.Button(mainframe, text="Simulate", command=simulate).grid(column=1, row=6, sticky=W)

preview_var = StringVar()

ttk.Label(mainframe, textvariable=preview_var).grid(column=3, row=6)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

prefix_entry.focus()
root.bind("<Return>", simulate)

root.mainloop()
