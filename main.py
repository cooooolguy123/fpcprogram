from tkinter import filedialog
from tkinter import *
import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledFrame
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.constants import *
import os, shutil, xlsxwriter, pandas, random, csv
from datetime import datetime
from threading import Thread
from pathlib import Path
from win32com import client


class Run(ttk.Frame):
    def __init__(self, parent):
        self.parent = parent
        self.theme = ttk.Style()
        self.name_text_var = StringVar()

        ttk.Frame.__init__(self, parent)
        self.initialize_ui()

    def add_name(self, event):
        self.selection = self.shift_table.view.selection()
        self.name_window = ttk.Toplevel()
        self.name_window.geometry(f"230x100+{self.parent.winfo_x() + 80}+{self.parent.winfo_y() + 217}")
        self.name_window.resizable(False, False)
        self.name_window.title("Add Name")
        self.r_data = []
        for iid in self.selection:
            print(self.selection)
            r: ttk.tableview.TableRow = self.shift_table.iidmap.get(iid)
            self.r_data.append(r.values)
            print(self.r_data)
            r.delete()
        self.selection = self.r_data[0]
        self.id_label = ttk.Label(self.name_window, text=f"ID: {self.selection[0]}")
        self.id_label.grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.id_label.config(font=("Arial", 12))
        self.shift_label = ttk.Label(self.name_window, text=f"Shift: {self.selection[2]}")
        self.shift_label.grid(row=0, column=0, sticky=E, padx=5, pady=5)
        self.shift_label.config(font=("Arial", 12))
        self.name_frame = ttk.LabelFrame(self.name_window, text=" Change Name: ", width=3)
        self.name_frame.grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.name_entry = ttk.Entry(self.name_frame, width=20, textvariable=self.name_text_var)
        self.name_entry.grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.name_change_btn = ttk.Button(self.name_frame, text="Apply", width=7, command=self.execute_name_change)
        self.name_change_btn.grid(row=0, column=1, sticky=W, padx=5, pady=5)
        self.name_window.mainloop()

    def execute_name_change(self):
        self.shift_table.insert_row("end", (self.selection[0], self.name_text_var.get(), self.selection[2]))
        self.shift_table.load_table_data()
        self.name_entry.delete("0", "end")
        self.row_list = []
        for iid in self.shift_table.view.get_children():
            row: ttk.tableview.TableRow = self.shift_table.iidmap.get(iid)
            self.row_data = []
            self.row_data.append(row.values)
            self.row_data2 = self.row_data[0]
            self.row_info = {"ID": self.row_data2[0], "Name": self.row_data2[1], "Shift": self.row_data2[2]}
            self.row_list.append(self.row_info)

        self.name_df = pandas.DataFrame(self.row_list, columns=["ID", "Name", "Shift"])
        if self.shift_choice.get() == 1:
            if Path(os.path.abspath("fpcs1df.csv")).is_file():
                os.remove(os.path.abspath("fpcs1df.csv"))
                self.name_df.to_csv("fpcs1df.csv", encoding="utf-8", index=False, header=True)
            else:
                self.name_df.to_csv("fpcs1df.csv", encoding="utf-8", index=False, header=True)
        elif self.shift_choice.get() == 2:
            if Path(os.path.abspath("fpcs2df.csv")).is_file():
                os.remove(os.path.abspath("fpcs2df.csv"))
                self.name_df.to_csv("fpcs2df.csv", encoding="utf-8", index=False, header=True)
            else:
                self.name_df.to_csv("fpcs2df.csv", encoding="utf-8", index=False, header=True)
        print(self.name_df)
        self.name_window.destroy()

    def initialize_ui(self):
        self.parent.geometry("390x600+500+200")
        self.parent.title("Fulfillment Performance Calculator")
        self.parent.resizable(False, False)
        self.tabs = ttk.Notebook(self.parent, width=60)
        self.tabs.pack(expand=1, fill=BOTH)
        self.tab1 = ttk.Frame(self.tabs)
        self.tab2 = ttk.Frame(self.tabs)
        self.tab3 = ttk.Frame(self.tabs)
        self.tabs.add(self.tab1, text="Home")
        self.tabs.add(self.tab2, text="Help")
        self.tabs.add(self.tab3, text="About")

        ## Text Variables // StringVar()'s ##
        self.theme_text_var = StringVar()
        self.source_text_var = StringVar()
        self.save_text_var = StringVar()
        self.save_file_text_var = StringVar()
        self.qc_text_var = StringVar()
        self.pick_text_var = StringVar()
        self.replen_text_var = StringVar()
        self.ib_text_var = StringVar()
        self.acc_text_var = StringVar()

        ## OTHER VARIABLES ##
        self.pdf_save_path = f"{Path(self.save_text_var.get())}{self.save_file_text_var.get()}2.pdf"
        self.desktop = os.path.expanduser("~/Desktop")
        self.doc_dir = os.path.expanduser("~/Documents")
        self.color_bool_var = ttk.BooleanVar(value=True)
        self.pdf_bool_var = ttk.BooleanVar(value=False)
        self.shift_choice = IntVar()
        self.shift_choice.set(1)
        self.shift = None
        self.debug_queue = []
        self.settings_file = f"{self.doc_dir}/fpcsettings.xlsx"

        # READ SETTINGS FILE // CREATE SETTINGS FILE
        if Path(self.settings_file).is_file():
            # READ FILE // APPLY SETTINGS
            self.load_settings = pandas.read_excel(self.settings_file)
            self.theme.theme_use(str(self.load_settings.iat[3, 1]))
            self.selected_theme = str(self.load_settings.iat[3, 1])
            self.color_bool_var.set(self.load_settings.iat[0, 1])
            self.pdf_bool_var.set(self.load_settings.iat[1, 1])
            self.shift_choice.set(int(self.load_settings.iat[5, 0]))
        else:
            # CREATE DEFAULT SETTINGS
            self.load_settings = xlsxwriter.Workbook("fpcsettings.xlsx")
            self.settings = self.load_settings.add_worksheet()
            self.settings.write(0, 0, "Goals")
            self.settings.write(1, 0, 40)  # QC GOAL DEFAULT
            self.settings.write(2, 0, 35)  # PICK GOAL DEFAULT
            self.settings.write(3, 0, 35)  # IB GOAL DEFAULT
            self.settings.write(4, 0, 45)  # REPLEN GOAL DEFAULT
            self.settings.write(5, 0, 150)  # ACC GOAL DEFAULT
            self.settings.write(0, 1, "ChkBxs")
            self.settings.write(1, 1, self.color_bool_var.get())
            self.settings.write(2, 1, self.pdf_bool_var.get())
            self.settings.write(3, 1, "Theme")
            self.settings.write(4, 1, "light")
            self.settings.write(0, 2, "SVDIR")
            self.settings.write(1, 2, self.desktop)
            self.settings.write(2, 2, "SVSTR")
            self.settings.write(3, 2, "ExportedUPH")
            self.settings.write(6, 0, int(self.shift_choice.get()))
            self.load_settings.close()
            shutil.move(os.path.abspath("fpcsettings.xlsx"), self.doc_dir)
            self.load_settings = pandas.read_excel(self.settings_file)
            self.selected_theme = str(self.load_settings.iat[3, 1])
            self.theme.theme_use(self.selected_theme)

        # self.tab1 -- Homepage --

        # -- Step 1 -- #
        self.step1 = ttk.LabelFrame(self.tab1, text=" Step 1: Upload Excel spreadsheet ", width=3)
        self.step1.grid(ipadx=4, row=0, column=0, padx=(10, 80), pady=(10, 0), sticky=W)

        self.xlsx_text = ttk.Label(self.step1, text="(.xlsx)", width=7)
        self.xlsx_text.grid(row=0, column=0, sticky=W, padx=(5, 0))
        self.source_entry = ttk.Entry(self.step1, width=37, textvariable=self.source_text_var)
        self.source_entry.grid(row=0, column=1, sticky=W, pady=(5, 10))
        self.source_entry.bind("<Key>", lambda e: "break")
        self.browse_btn = ttk.Button(self.step1, text="Browse", width=7, command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, sticky=W, padx=(5, 0), pady=(0, 5))

        # -- Step 2 -- #
        self.step2 = ttk.LabelFrame(self.tab1, text=" Step 2: Manage Preferences ", width=3)
        self.step2.grid(ipadx=99, row=1, column=0, padx=(10, 80), pady=(10, 0), sticky=W)

        self.step2_tabs = ttk.Notebook(self.step2, width=30)
        self.step2_tabs.pack(expand=1, fill=BOTH)
        self.step2_tab1 = ttk.Frame(self.step2_tabs)
        self.step2_tab2 = ttk.Frame(self.step2_tabs)
        self.step2_tab3 = ttk.Frame(self.step2_tabs)
        self.step2_tabs.add(self.step2_tab1, text="Goals")
        self.step2_tabs.add(self.step2_tab2, text="shift")
        self.step2_tabs.add(self.step2_tab3, text="Misc.")

        ## Goals ##
        self.qc_text = ttk.Label(self.step2_tab1, text="QC", width=4)
        self.qc_text.grid(row=0, column=0, sticky=W, padx=(30, 0), pady=(20, 0))
        self.qc_entry = ttk.Entry(self.step2_tab1, width=5, textvariable=self.qc_text_var)
        self.qc_entry.grid(row=0, column=1, padx=(5, 0), pady=(20, 0))
        self.pick_text = ttk.Label(self.step2_tab1, text="S2//OB", width=8)
        self.pick_text.grid(row=0, column=2, padx=(40, 0), pady=(20, 0))
        self.pick_entry = ttk.Entry(self.step2_tab1, width=5, textvariable=self.pick_text_var)
        self.pick_entry.grid(row=0, column=3, padx=(5, 0), pady=(20, 0))
        self.ib_text = ttk.Label(self.step2_tab1, text="IB", width=4)
        self.ib_text.grid(row=1, column=0, sticky=W, padx=(30, 0), pady=(20, 0))
        self.ib_entry = ttk.Entry(self.step2_tab1, width=5, textvariable=self.ib_text_var)
        self.ib_entry.grid(row=1, column=1, padx=(5, 0), pady=(20, 0))
        self.replen_text = ttk.Label(self.step2_tab1, text="Replen", width=8)
        self.replen_text.grid(row=1, column=2, sticky=W, padx=(40, 0), pady=(20, 0))
        self.replen_entry = ttk.Entry(self.step2_tab1, width=5, textvariable=self.replen_text_var)
        self.replen_entry.grid(row=1, column=3, padx=(5, 0), pady=(20, 0))
        self.acc_text = ttk.Label(self.step2_tab1, text="ACC", width=5)
        self.acc_text.grid(row=2, column=0, sticky=W, padx=(30, 0), pady=(20, 0))
        self.acc_entry = ttk.Entry(self.step2_tab1, width=5, textvariable=self.acc_text_var)
        self.acc_entry.grid(row=2, column=1, padx=(5, 0), pady=(20, 0))
        self.qc_entry.insert(1, self.load_settings.iat[0, 0])
        self.pick_entry.insert(1, self.load_settings.iat[1, 0])
        self.ib_entry.insert(1, self.load_settings.iat[2, 0])
        self.replen_entry.insert(1, self.load_settings.iat[3, 0])
        self.acc_entry.insert(1, self.load_settings.iat[4, 0])
        self.qc_entry.config(state=DISABLED)
        self.pick_entry.config(state=DISABLED)
        self.ib_entry.config(state=DISABLED)
        self.replen_entry.config(state=DISABLED)
        self.acc_entry.config(state=DISABLED)

        ## SHIFT ##
        self.shift_frame = ttk.LabelFrame(self.step2_tab2, text=" Shift Selection: ", width=3)
        self.shift_frame.grid(row=0, column=0, sticky=W, padx=(5, 0), pady=(5, 5))

        self.table_frame = ScrolledFrame(self.shift_frame, autohide=True, width=344, height=140)
        self.table_frame.grid(row=0, column=0, padx=(5, 0), sticky=W)

        # SHIFT RADIOBUTTONS
        self.first_shift_radiobtn = ttk.Radiobutton(self.table_frame, text="First", variable=self.shift_choice, value=1, command=self.shift_data_thread)
        self.first_shift_radiobtn.grid(row=0, column=0, sticky=W, padx=(85, 0), pady=(5, 0))

        self.second_shift_radiobtn = ttk.Radiobutton(self.table_frame, text="Second", variable=self.shift_choice, value=2, command=self.shift_data_thread)
        self.second_shift_radiobtn.grid(row=0, column=0, sticky=W, padx=(140, 0), pady=(5, 0))

        self.both_shifts_radiobtn = ttk.Radiobutton(self.table_frame, text="Both", variable=self.shift_choice, value=3, command=self.shift_data_thread)
        self.both_shifts_radiobtn.grid(row=0, column=0, sticky=W, padx=(210, 0), pady=(5, 0))

        self.first_shift_radiobtn.config(state=DISABLED)
        self.second_shift_radiobtn.config(state=DISABLED)
        self.both_shifts_radiobtn.config(state=DISABLED)

        # SHIFT TABLE
        self.col_data = [
            {"text": "EmployeeID", "stretch": False},
            {"text": "Name", "stretch": False},
            {"text": "Shift", "stretch": False}
        ]
        self.row_data = []
        self.shift_table = Tableview(self.table_frame, coldata=self.col_data, rowdata=self.row_data, paginated=False,
                                     searchable=False, height=3)

        self.shift_table.grid(row=2, column=0, sticky=W, padx=(65, 5), pady=(5, 0))
        self.shift_table.autofit_columns()
        self.shift_table.view.bind("<Double-1>", self.add_name)

        ## MISC. ##

        # SAVE LOCATION
        self.default_save_dir = os.path.expanduser("~/Desktop")
        self.save_location = ttk.LabelFrame(self.step2_tab3, text=" Save Location: ", width=3)
        self.save_location.grid(row=0, column=0, sticky=N, padx=(10, 0), pady=(10, 0), ipadx=5, ipady=3)

        self.save_location_entry = ttk.Entry(self.save_location, width=40, textvariable=self.save_text_var)
        self.save_location_entry.grid(row=0, column=0, sticky=W, padx=(5, 0), pady=(5, 5))
        self.save_location_entry.insert("0", self.load_settings.iat[0, 2])
        self.save_location_entry.bind("<Key>", lambda e: "break")
        self.save_location_btn = ttk.Button(self.save_location, text="Browse", width=7, command=self.save_location_browse)
        self.save_location_btn.grid(row=0, column=1, sticky=W, padx=(5, 0), pady=(5, 5))

        # CHANGE FILE NAME
        self.save_file_frame = ttk.LabelFrame(self.step2_tab3, text=" Change File Name ", width=3)
        self.save_file_frame.grid(row=1, column=0, sticky=W, padx=(10, 0), pady=(10, 0), ipady=2)
        self.save_file_entry = ttk.Entry(self.save_file_frame, width=20, textvariable=self.save_file_text_var)
        self.save_file_entry.grid(row=0, column=0, sticky=W, padx=(5, 0), pady=(5, 5))
        self.save_file_entry.insert("0", self.load_settings.iat[2, 2])
        self.save_file_label = ttk.Label(self.save_file_frame, text=".xlsx", width=5)
        self.save_file_label.grid(row=0, column=1, sticky=W, padx=(5, 0), pady=(5, 5))

        # MISC. CHECKBOXES
        self.add_settings_frame = ttk.LabelFrame(self.step2_tab3, text=" Additional Settings ", width=3)
        self.add_settings_frame.grid(row=1, column=0, padx=(200, 0))

        self.color_checkbox = ttk.Checkbutton(self.add_settings_frame, text="Use Colors for Goal", variable=self.color_bool_var, command=self.save_thread_func)
        self.color_checkbox.grid(row=0, column=0, padx=(10, 5), pady=(5, 5), ipadx=7)

        self.pdf_checkbox = ttk.Checkbutton(self.add_settings_frame, text="Export as PDF", variable=self.pdf_bool_var, command=self.change_export_text)
        self.pdf_checkbox.grid(row=1, column=0, padx=(10, 5), pady=(0, 5), ipadx=5, sticky=W)
        if self.pdf_bool_var.get() is True:
            self.save_file_label.config(text=".pdf")

        # -- Step 3 -- #
        self.step3 = ttk.LabelFrame(self.tab1, text=" Step 3: Watch the magic happen ", width=3)
        self.step3.grid(row=2, column=0, padx=(10, 0), pady=(10, 0), sticky=W, ipadx=5, ipady=10)

        self.debug_log = ttk.ScrolledText(master=self.step3, width=55, height=10)
        self.debug_log.vbar.config(width=0)
        self.debug_log.grid(row=0, column=0, sticky=W, padx=(10, 0), pady=(10, 0))

        self.debug_log.bind("<Key>", lambda e: "break")
        self.load_bar = ttk.Progressbar(self.step3, length=285, orient=HORIZONTAL, mode="determinate")
        self.load_bar.grid(row=1, column=0, sticky=W, pady=(10, 0), padx=(10, 0))
        self.run_btn = ttk.Button(self.step3, width=5, text="Run", command=self.debug_thread)
        self.run_btn.grid(row=1, column=0, sticky=E, pady=(10, 0))
        self.run_btn.config(state=DISABLED)

        # self.tab3 ABOUT

        # ABOUT TEXT
        self.about_text_info = ttk.ScrolledText(master=self.tab3, width=59, height=30)
        self.about_text_info.vbar.config(width=0)
        self.about_text_info.grid(row=0, column=0, sticky=W, padx=(10, 0), pady=(15, 0))
        self.about_text_info.bind("<Key>", lambda e: "break")
        self.about_text_info.insert("1.0", "______________________________________________________________________\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                  .::::::::::.                                                     |\n"
                                    "|                                                .::``::::::::::.                                                   |\n"
                                    "|                                                :::..:::::::::::                                                   |\n"
                                    "|                                                ````````::::::::                                                   |\n"
                                    "|                                        .::::::::::::::::::::::: iiiiiii,                                          |\n"
                                    "|                                     .:::::::::::::::::::::::::: iiiiiiiii.                                        |\n"
                                    "|                                     ::::::::::::::::::::::::::: iiiiiiiiii                                        |\n"
                                    "|                                     ::::::::::::::::::::::::::: iiiiiiiiii                                        |\n"
                                    "|                                     :::::::::: ,,,,,,,,,,,,,,,,,iiiiiiiiii                                        |\n"
                                    "|                                     :::::::::: iiiiiiiiiiiiiiiiiiiiiiiiiii                                        |\n"
                                    "|                                     `::::::::: iiiiiiiiiiiiiiiiiiiiiiiiii`                                        |\n"
                                    "|                                        `:::::: iiiiiiiiiiiiiiiiiiiiiii`                                           |\n"
                                    "|                                                iiiiiiii,,,,,,,,                                                   |\n"
                                    "|                                                iiiiiiiiiii''iii                                                   |\n"
                                    "|                                                `iiiiiiiiii..ii`                                                   |\n"
                                    "|                                                  `iiiiiiiiii`                                                     |\n"
                                    "|                                                                                                                   |\n"
                                    "|                  FULFILLMENT PERFORMANCE CALCULATOR                  |\n"
                                    "|                                                  v1.0.0                                                       |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                          PYTHON VERSION 3.6+ REQUIRED                             |\n"
                                    "|                         EXCEL MUST BE INSTALLED ON PC                             |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    ## "|         THIS IS A PORTABLE EXE, NO INSTALLATION REQUIRED         |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                   ABOUT:                                                  |\n"
                                    "|                                                                                                                   |\n"
                                    "|       THIS TOOL USES THE EXPORTED 'SEARCH TRANSACTIONS'     |\n"
                                    "|       EXCEL SPREADSHEET TO CREATE A PERFORMANCE BASED     |\n"
                                    "|            DOCUMENT WITH ALL ASSIGNMENTS FROM EVERY            |\n"
                                    "|               EMPLOYEE FROM THAT GIVEN DATE WITH ZERO              |\n"
                                    "|                                            REDUNDANCIES.                                         |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                                                                                                                   |\n"
                                    "|                       DEVELOPED ENTIRELY IN PYTHON 3.9 BY:                   |\n"
                                    "|                    NATHAN GARDIKIS | FULFILLMENT TRAINER                 |\n"
                                    "|_____________________________________________________________________|\n")
        self.about_text_info.yview_pickplace("0")
        self.about_frame = ttk.Frame(self.tab3, width=3)
        self.about_frame.grid(row=1, column=0, sticky=W, padx=(10, 0), pady=(10, 0))

        # CHANGE APP THEME
        self.theme_frame = ttk.LabelFrame(self.about_frame, text=" Change App Theme ", width=3)
        self.theme_frame.grid(row=1, column=0, sticky=W, padx=(195, 0), pady=(9, 0), ipady=5)
        self.theme_dropdown = ttk.Combobox(self.theme_frame, width=10, textvariable=self.theme_text_var)
        self.theme_dropdown["values"] = ("light", "dark", "red")
        self.theme_dropdown.insert("0", self.selected_theme)
        self.theme_dropdown.bind("<Key>", lambda e: "break")
        self.theme_dropdown.grid(row=0, column=0, sticky=W, padx=(5, 0), pady=(5, 0))
        self.apply_btn = ttk.Button(self.theme_frame, text="Apply", width=5, command=self.change_theme)
        self.apply_btn.grid(row=0, column=1, sticky=W, padx=(5, 5), pady=(5, 0))

    def change_theme(self):
        self.theme.theme_use(self.theme_text_var.get())
        self.selected_theme = self.theme_text_var.get()
        Thread(target=self.save_settings, daemon=True).start()

    def save_thread_func(self):
        Thread(target=self.save_settings, daemon=True).start()

    def save_settings(self):
        os.remove(self.settings_file)
        self.load_settings = xlsxwriter.Workbook("fpcsettings.xlsx")
        self.settings = self.load_settings.add_worksheet()
        self.settings.write(0, 0, "Goals")
        self.qc_entry.config(state=NORMAL)
        self.settings.write(1, 0, self.qc_text_var.get())
        self.qc_entry.config(state=DISABLED)
        self.pick_entry.config(state=NORMAL)
        self.settings.write(2, 0, self.pick_text_var.get())
        self.pick_entry.config(state=DISABLED)
        self.ib_entry.config(state=NORMAL)
        self.settings.write(3, 0, self.ib_text_var.get())
        self.ib_entry.config(state=DISABLED)
        self.replen_entry.config(state=NORMAL)
        self.settings.write(4, 0, self.replen_text_var.get())
        self.replen_entry.config(state=DISABLED)
        self.acc_entry.config(state=NORMAL)
        self.settings.write(5, 0, self.acc_text_var.get())
        self.acc_entry.config(state=DISABLED)
        self.settings.write(0, 1, "ChkBxs")
        self.settings.write(1, 1, self.color_bool_var.get())
        self.settings.write(2, 1, self.pdf_bool_var.get())
        self.settings.write(3, 1, "Theme")
        self.settings.write(4, 1, str(self.theme_text_var.get()))
        self.settings.write(0, 2, "SVDIR")
        self.settings.write(1, 2, self.save_text_var.get())
        self.settings.write(2, 2, "SVSTR")
        self.settings.write(3, 2, self.save_file_text_var.get())
        self.settings.write(6, 0, int(self.shift_choice.get()))
        self.load_settings.close()
        shutil.move(os.path.abspath("fpcsettings.xlsx"), self.doc_dir)
        self.load_settings = pandas.read_excel(self.settings_file)
        if Path(self.source_text_var.get()).is_file():
            self.qc_entry.config(state=NORMAL)
            self.pick_entry.config(state=NORMAL)
            self.ib_entry.config(state=NORMAL)
            self.replen_entry.config(state=NORMAL)
            self.acc_entry.config(state=NORMAL)
            self.first_shift_radiobtn.config(state=NORMAL)
            self.second_shift_radiobtn.config(state=NORMAL)
            self.both_shifts_radiobtn.config(state=NORMAL)

    def shift_data_thread(self):
        Thread(target=self.compile_shift_data, daemon=True).start()
        Thread(target=self.save_settings, daemon=True).start()

    def compile_shift_data(self):
        self.employee_id_list = []
        self.second_shift_start_time = datetime.strptime("17:00:00", "%H:%M:%S").time()
        self.row_data = []
        if self.shift_choice.get() == 2:
            for row in range(len(self.transactions.index)):
                if row > 0:
                    self.eid = self.transactions.iat[row, 2]
                    self.eid_time = datetime.strptime(self.transactions.iat[row, 4], "%H:%M:%S").time()
                    if int(self.eid_time.hour) > int(self.second_shift_start_time.hour):
                        if len(str(self.eid)) <= 8:
                            if self.eid not in self.employee_id_list:
                                if not str(self.eid) == "AUTO":
                                    self.employee_id_list.append(self.eid)
        elif self.shift_choice.get() == 1:
            for row in range(len(self.transactions.index)):
                if row > 0:
                    self.eid = self.transactions.iat[row, 2]
                    self.eid_time = datetime.strptime(self.transactions.iat[row, 4], "%H:%M:%S").time()
                    if int(self.eid_time.hour) < int(self.second_shift_start_time.hour):
                        if len(str(self.eid)) <= 8:
                            if self.eid not in self.employee_id_list:
                                if not str(self.eid) == "AUTO":
                                    self.employee_id_list.append(self.eid)
            if Path(os.path.abspath("fpcs1df.csv")).is_file():
                self.s1_names = pandas.read_csv(os.path.abspath("fpcs1df.csv"))
                try:
                    for row in range(len(self.employee_id_list)):
                        self.eid = self.s1_names.iat[row, 0]
                        self.name = self.s1_names.iat[row, 1]
                        self.s = self.s1_names.iat[row, 2]
                        self.row_data.append((self.employee_id_list[row], self.name, self.s))
                except IndexError:
                    for eid in range(len(self.employee_id_list)):
                        if self.s1_names.iat[eid, 0] not in self.employee_id_list:
                            self.row_data.append((self.employee_id_list[eid], "TBD", self.shift_choice.get()))
        elif self.shift_choice.get() == 3:
            for row in range(len(self.transactions.index)):
                self.eid = self.transactions.iat[row, 2]
                if len(str(self.eid)) <= 8:
                    if self.eid not in self.employee_id_list:
                        if not str(self.eid) == "AUTO":
                            self.employee_id_list.append(self.eid)
        else:
            for row in range(len(self.transactions.index)):
                self.eid = self.transactions.iat[row, 2]
                if len(str(self.eid)) <= 8:
                    if self.eid not in self.employee_id_list:
                        if not str(self.eid) == "AUTO":
                            self.employee_id_list.append(self.eid)
        if len(self.shift_table.get_rows(visible=True)) > 0:
            self.shift_table.delete_rows()

        self.shift_table.configure(height=len(self.employee_id_list))

        if self.shift_choice.get() == 1:
            if Path(os.path.abspath("fpcs1df.csv")).is_file():
                for row in range(len(self.row_data)):
                    self.shift_table.insert_row("end", list(self.row_data[row]))
                    self.shift_table.autofit_columns()
                    self.shift_table.load_table_data()
        else:
            for eid in range(len(self.employee_id_list)):
                self.row_data.append((self.employee_id_list[eid], "TBD", self.shift_choice.get()))
                self.shift_table.insert_row("end", [self.employee_id_list[eid], "TBD", self.shift_choice.get()])
                self.shift_table.autofit_columns()
                self.shift_table.load_table_data()
        self.shift_table.autofit_columns()
        self.shift_table.load_table_data()

    def change_export_text(self):
        if self.pdf_bool_var.get() is True:
            self.save_file_label.config(text=".pdf")
        elif self.pdf_bool_var.get() is False:
            self.save_file_label.config(text=".xlsx")
        Thread(target=self.save_settings, daemon=True).start()

    def browse_file(self):
        self.initial_file = filedialog.askopenfilename(title="Grab exported transaction spreadsheet",
                                                       filetypes=(("Excel Files  (*.xlsx)", "*.xlsx"),))
        self.source_entry.delete("0", END)
        self.debug_log.delete("0.0", "end")
        self.source_entry.insert("0", self.initial_file)
        self.csv_name = os.path.basename(self.initial_file).replace(".xlsx", ".csv")
        self.csv_temp = os.path.normpath(f"{self.doc_dir}\\{self.csv_name}")
        if Path(self.csv_temp).is_file():
            os.remove(self.csv_temp)
        Thread(target=self.read_tran_spreadsheet, daemon=True).start()
        Thread(target=self.debug_feed, daemon=True).start()
        if Path(self.source_text_var.get()).is_file():
            self.qc_entry.config(state=NORMAL)
            self.pick_entry.config(state=NORMAL)
            self.ib_entry.config(state=NORMAL)
            self.replen_entry.config(state=NORMAL)
            self.acc_entry.config(state=NORMAL)
        if self.load_bar["value"] == 100:
            self.load_bar["value"] = 0
            self.debug_queue = []
            self.debug_log.delete("0.0", "end")
            self.debug_log.update()

        Thread(target=self.save_settings, daemon=True).start()

    def read_tran_spreadsheet(self):
        self.debug_queue.append("Reading transaction spreadsheet... \n")
        try:
            shutil.copy(Path(self.initial_file), self.doc_dir)
            self.xlsx_name = os.path.basename(self.initial_file)
            self.temp_xlsx = os.path.normpath(f"{self.doc_dir}\\{self.xlsx_name}")
            assert os.path.isfile(self.temp_xlsx)
            self.excel = client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.xlsx_file = self.excel.Workbooks.Open(self.temp_xlsx)
            self.csv_name = self.temp_xlsx.replace(".xlsx", ".csv")
            self.xlsx_file.SaveAs(Filename=os.path.normpath(self.csv_name), FileFormat=6)
            self.transactions = pandas.read_csv(self.csv_name)
            Thread(target=self.compile_shift_data, daemon=True).start()
            os.remove(self.temp_xlsx)
            self.debug_queue.append("Action complete; hit the 'run' button below to get started... \n")
            self.run_btn.config(state=NORMAL)
            self.xlsx_file.Close(True)
            self.excel.Quit()
        except shutil.SameFileError:
            self.debug_queue.append("Error: Settings file uploaded as transaction spreadsheet... \n")
            self.debug_queue.append("Upload another transaction spreadsheet... \n")

    def save_location_browse(self):
        self.save_location_entry.delete("0", END)
        self.location = filedialog.askdirectory(initialdir="Desktop")
        self.save_location_entry.insert("0", self.location)
        Thread(target=self.save_settings, daemon=True).start()

    def debug_feed(self):
        try:
            self.debug_log.insert(END, f"\n{str(self.debug_queue[0])}")
            self.debug_log.see(END)
            self.debug_log.update()
            self.debug_queue.remove(self.debug_queue[0])

        except IndexError:
            pass
        finally:
            if self.load_bar["value"] <= 100:
                self.parent.after(10, self.debug_feed)

    def debug_thread(self):
        self.run_btn.config(state=DISABLED)
        self.color_checkbox.config(state=DISABLED)
        self.pdf_checkbox.config(state=DISABLED)
        self.save_file_entry.config(state=DISABLED)
        self.browse_btn.config(state=DISABLED)
        self.save_location_btn.config(state=DISABLED)
        self.debug_queue.append("Reading your transaction spreadsheet: \n")
        self.debug_queue.append(f"{self.source_text_var.get()} \n")
        self.debug_queue.append("Loading......... \n")

        Thread(target=self.save_settings, daemon=True).start()
        Thread(target=self.compile_data, daemon=True).start()

    def compile_data(self):
        Thread(target=self.create_new_spreadsheet, daemon=True).start()
        self.debug_queue.append(f"{len(self.transactions.index)} transactions found... \n")
        self.load_bar["value"] = 5
        # - OB/S2 Picks
        self.pick_eid_dict = {}
        self.pick_total_dict = {}
        self.pick_uph_dict = {}
        self.pick_totals = []
        # Accessory Picks
        self.acc_pick_eid_dict = {}
        self.acc_pick_total_dict = {}
        self.acc_pick_uph_dict = {}
        self.acc_pick_totals = []
        # IB/RTW Putaways
        self.dir_put_eid_dict = {}
        self.dir_put_total_dict = {}
        self.dir_put_uph_dict = {}
        self.dir_put_totals = []
        # Replen Picks
        self.rep_pick_eid_dict = {}
        self.rep_pick_total_dict = {}
        self.rep_pick_uph_dict = {}
        self.rep_pick_totals = []
        # Replen Putaways
        self.rep_put_eid_dict = {}
        self.rep_put_total_dict = {}
        self.rep_put_uph_dict = {}
        self.rep_put_totals = []
        # QC Passes
        self.qc_pass_eid_dict = {}
        self.qc_pass_total_dict = {}
        self.qc_pass_uph_dict = {}
        self.qc_pass_totals = []

        self.stores_list = []
        self.codes_list = []
        self.tracked_codes = [301, 305, 362, 917, 212, 256]
        self.walk_list = ["W01", "W02", "W03", "W04", "W05", "W06", "W07", "W08", "W09", "W10", "W11", "W12", "W13",
                          "W14", "W15", "W16", "W17", "W18", "W19", "W20", "W21", "W22", "W23", "W24", "W25", "W26",
                          "W27", "W28", "W29", "W30", "W31", "W32", "W33", "W34", "W35"]
        self.op_list = ["P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "S07"
                        "R01", "R02", "R03", "R04", "S01", "S02", "S03", "S04", "S05", "S06"]
        self.sitdown_list = ["S01", "S02", "S03", "S04", "S05", "S06"]
        for row in range(len(self.transactions.index)):
            self.store = self.transactions.iat[row, 18]
            self.code = self.transactions.iat[row, 0]
            if len(str(self.store)) <= 4:
                if not str(self.store) == "nan":
                    if self.store not in self.stores_list:
                        self.stores_list.append(self.store)
            if len(str(self.code)) == 3:
                if self.code not in self.codes_list:
                    self.codes_list.append(self.code)
        self.debug_queue.append(f"Employee ID's Located (w/out duplicates) \n"
                                f"{str(self.employee_id_list)} \n")
        self.load_bar["value"] = 10
        self.debug_queue.append(f"Stores picked (w/out duplicates) \n"
                                f"{str(self.stores_list)} \n")
        self.load_bar["value"] = 15
        self.debug_queue.append(f"Transaction types: (w/out duplicates) \n"
                                f"{str(self.codes_list)} \n")
        self.load_bar["value"] = 20
        self.debug_queue.append("Gathering totals.... \n")

        for eid in range(len(self.employee_id_list)):
            self.pick_total = 0
            self.acc_pick_total = 0
            self.dir_put_total = 0
            self.rep_pick_total = 0
            self.rep_put_total = 0
            self.qc_pass_total = 0
            self.pick_times = []
            self.acc_pick_times = []
            self.dir_put_times = []
            self.rep_pick_times = []
            self.rep_put_times = []
            self.qc_pass_times = []
            for row in range(len(self.transactions.index)):
                if self.employee_id_list[eid] in self.transactions.iat[row, 2]:
                    if str(self.tracked_codes[0]) in str(self.transactions.iat[row, 0]):
                        # ACCESSORY PICK COUNT
                        if len(self.transactions.iat[row, 11]) == 7 and str(self.transactions.iat[row, 12]) in self.walk_list:
                            self.acc_pick_total += int(self.transactions.iat[row, 10])
                            self.acc_pick_times.append(self.transactions.iat[row, 4])
                        # OP OPERATOR PICK COUNT
                        if len(self.transactions.iat[row, 11]) == 7 and str(self.transactions.iat[row, 12]) in self.op_list:
                            self.pick_total += int(self.transactions.iat[row, 10])
                            self.pick_times.append(self.transactions.iat[row, 4])
                    if str(self.tracked_codes[2]) in str(self.transactions.iat[row, 0]):
                        # QC PASS COUNT
                        if len(self.transactions.iat[row, 11]) == 5 and len(self.transactions.iat[row, 12]) == 6:
                            self.qc_pass_total += int(self.transactions.iat[row, 10])
                            self.qc_pass_times.append(self.transactions.iat[row, 4])
                    if str(self.tracked_codes[4]) in str(self.transactions.iat[row, 0]):
                        # INBOUND PUT COUNT
                        if str(self.transactions.iat[row, 11]) in self.op_list and len(self.transactions.iat[row, 12]) == 7:
                            self.dir_put_total += int(self.transactions.iat[row, 10])
                            self.dir_put_times.append(self.transactions.iat[row, 4])
                    if str(self.tracked_codes[5]) in str(self.transactions.iat[row, 0]):
                        # REPLEN PICK COUNT
                        if str(self.transactions.iat[row, 11]) in self.sitdown_list and str(self.transactions.iat[row, 12]) == "REP001":
                            self.rep_pick_total += int(self.transactions.iat[row, 10])
                            self.rep_pick_times.append(self.transactions.iat[row, 4])
                    if str(self.tracked_codes[3]) in str(self.transactions.iat[row, 0]) or str(self.tracked_codes[4]) in str(self.transactions.iat[row, 0]):
                        # REPLEN PUTAWAY COUNT
                        if str(self.transactions.iat[row, 11]) in self.op_list and len(self.transactions.iat[row, 12]) == 7:
                            if int(self.transactions.iat[row, 10]) > 1:
                                self.rep_put_total += int(self.transactions.iat[row, 10])
                                self.rep_put_times.append(self.transactions.iat[row, 4])

            # +---|| TOTAL HOUR COUNT ||--+--|| CALCULATE SPEED ||---+ #
            if len(self.pick_times) > 0:
                self.total = 0
                for time in range(len(self.pick_times)):
                    try:
                        self.t1 = datetime.strptime(self.pick_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.pick_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.25:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.pick_start = datetime.strptime(self.pick_times[len(self.pick_times) - 1], "%H:%M:%S").time()
                self.pick_end = datetime.strptime(self.pick_times[0], "%H:%M:%S").time()
                self.pick_time_total = (datetime.combine(datetime.now(), self.pick_end) - datetime.combine(datetime.now(), self.pick_start)).total_seconds() / 3600

                if int(self.pick_start.hour) < 3:
                    self.pick_start2 = datetime.strptime(f"{int(self.pick_start.hour) + 21}:{int(self.pick_start.minute)}:{int(self.pick_start.second)}", "%H:%M:%S").time()
                else:
                    self.pick_start2 = datetime.strptime(f"{int(self.pick_start.hour) - 3}:{int(self.pick_start.minute)}:{int(self.pick_start.second)}", "%H:%M:%S").time()
                if int(self.pick_end.hour) < 3:
                    self.pick_end2 = datetime.strptime(f"{int(self.pick_end.hour) + 21}:{int(self.pick_end.minute)}:{self.pick_end.second}", "%H:%M:%S").time()
                    self.pick_time_total2 = (datetime.combine(datetime.now(), self.pick_end2) - datetime.combine(datetime.now(), self.pick_start2)).total_seconds() / 3600
                else:
                    self.pick_end2 = datetime.strptime(f"{int(self.pick_end.hour) - 3}:{int(self.pick_end.minute)}:{self.pick_end.second}", "%H:%M:%S").time()
                    self.pick_time_total2 = (datetime.combine(datetime.now(), self.pick_end2) - datetime.combine(datetime.now(), self.pick_start2)).total_seconds() / 3600
                self.pick_time_total2 -= self.total
                if self.pick_time_total2 > 0:
                    self.pick_uph = int((self.pick_total / self.pick_time_total2))
                elif self.pick_time_total2 < 0:
                    self.pick_uph = int((self.pick_total / 1))
                else:
                    self.pick_uph = 0

            if len(self.acc_pick_times) > 0:
                self.total = 0
                for time in range(len(self.acc_pick_times)):
                    try:
                        self.t1 = datetime.strptime(self.acc_pick_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.acc_pick_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.25:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.acc_pick_start = datetime.strptime(self.acc_pick_times[len(self.acc_pick_times) - 1], "%H:%M:%S").time()
                self.acc_pick_end = datetime.strptime(self.acc_pick_times[0], "%H:%M:%S").time()
                self.acc_pick_time_total = (datetime.combine(datetime.now(), self.acc_pick_end) - datetime.combine(datetime.now(), self.acc_pick_start)).total_seconds() / 3600
                if int(self.acc_pick_start.hour) < 3:
                    self.acc_pick_start2 = datetime.strptime(f"{int(self.acc_pick_start.hour) + 21}:{int(self.acc_pick_start.minute)}:{int(self.acc_pick_start.second)}", "%H:%M:%S").time()
                else:
                    self.acc_pick_start2 = datetime.strptime(f"{int(self.acc_pick_start.hour) - 3}:{int(self.acc_pick_start.minute)}:{int(self.acc_pick_start.second)}", "%H:%M:%S").time()
                if int(self.acc_pick_end.hour) < 3:
                    self.acc_pick_end2 = datetime.strptime(f"{int(self.acc_pick_end.hour) + 21}:{int(self.acc_pick_end.minute)}:{self.acc_pick_end.second}", "%H:%M:%S").time()
                    self.acc_pick_time_total2 = (datetime.combine(datetime.now(), self.acc_pick_end2) - datetime.combine(datetime.now(), self.acc_pick_start2)).total_seconds() / 3600
                else:
                    self.acc_pick_end2 = datetime.strptime(f"{int(self.acc_pick_end.hour) - 3}:{int(self.acc_pick_end.minute)}:{self.acc_pick_end.second}", "%H:%M:%S").time()
                    self.acc_pick_time_total2 = (datetime.combine(datetime.now(), self.acc_pick_end2) - datetime.combine(datetime.now(), self.acc_pick_start2)).total_seconds() / 3600
                self.acc_pick_time_total2 -= self.total
                if self.acc_pick_time_total2 > 0:
                    self.acc_pick_uph = int((self.acc_pick_total / self.acc_pick_time_total2))
                elif self.acc_pick_time_total2 < 0:
                    self.acc_pick_uph = int((self.acc_pick_total / 1))
                else:
                    self.acc_pick_uph = 0

            if len(self.qc_pass_times) > 0:
                self.total = 0
                for time in range(len(self.qc_pass_times)):
                    try:
                        self.t1 = datetime.strptime(self.qc_pass_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.qc_pass_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.25:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.qc_pass_start = datetime.strptime(self.qc_pass_times[len(self.qc_pass_times) - 1], "%H:%M:%S").time()
                self.qc_pass_end = datetime.strptime(self.qc_pass_times[0], "%H:%M:%S").time()
                self.qc_pass_time_total = (datetime.combine(datetime.now(), self.qc_pass_end) - datetime.combine(datetime.now(), self.qc_pass_start)).total_seconds() / 3600
                if int(self.qc_pass_start.hour) < 3:
                    self.qc_pass_start2 = datetime.strptime(f"{int(self.qc_pass_start.hour) + 21}:{int(self.qc_pass_start.minute)}:{int(self.qc_pass_start.second)}", "%H:%M:%S").time()
                else:
                    self.qc_pass_start2 = datetime.strptime(f"{int(self.qc_pass_start.hour) - 3}:{int(self.qc_pass_start.minute)}:{int(self.qc_pass_start.second)}", "%H:%M:%S").time()
                if int(self.qc_pass_end.hour) < 3:
                    self.qc_pass_end2 = datetime.strptime(f"{int(self.qc_pass_end.hour) + 21}:{int(self.qc_pass_end.minute)}:{self.qc_pass_end.second}", "%H:%M:%S").time()
                    self.qc_pass_time_total2 = (datetime.combine(datetime.now(), self.qc_pass_end2) - datetime.combine(datetime.now(), self.qc_pass_start2)).total_seconds() / 3600
                else:
                    self.qc_pass_end2 = datetime.strptime(f"{int(self.qc_pass_end.hour) - 3}:{int(self.qc_pass_end.minute)}:{self.qc_pass_end.second}", "%H:%M:%S").time()
                    self.qc_pass_time_total2 = (datetime.combine(datetime.now(), self.qc_pass_end2) - datetime.combine(datetime.now(), self.qc_pass_start2)).total_seconds() / 3600
                self.qc_pass_time_total2 -= self.total
                if self.qc_pass_time_total2 > 0:
                    self.qc_pass_uph = int((self.qc_pass_total / self.qc_pass_time_total2))
                elif self.qc_pass_time_total2 < 0:
                    self.qc_pass_uph = int((self.qc_pass_total / 1))
                else:
                    self.qc_pass_uph = 0

            if len(self.dir_put_times) > 0:
                self.total = 0
                for time in range(len(self.dir_put_times)):
                    try:
                        self.t1 = datetime.strptime(self.dir_put_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.dir_put_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.25:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.dir_put_start = datetime.strptime(self.dir_put_times[len(self.dir_put_times) - 1], "%H:%M:%S").time()
                self.dir_put_end = datetime.strptime(self.dir_put_times[0], "%H:%M:%S").time()
                self.dir_put_time_total = (datetime.combine(datetime.now(), self.dir_put_end) - datetime.combine(datetime.now(), self.dir_put_start)).total_seconds() / 3600
                if int(self.dir_put_start.hour) < 3:
                    self.dir_put_start2 = datetime.strptime(f"{int(self.dir_put_start.hour) + 21}:{int(self.dir_put_start.minute)}:{int(self.dir_put_start.second)}", "%H:%M:%S").time()
                else:
                    self.dir_put_start2 = datetime.strptime(f"{int(self.dir_put_start.hour) - 3}:{int(self.dir_put_start.minute)}:{int(self.dir_put_start.second)}", "%H:%M:%S").time()
                if int(self.dir_put_end.hour) < 3:
                    self.dir_put_end2 = datetime.strptime(f"{int(self.dir_put_end.hour) + 21}:{int(self.dir_put_end.minute)}:{self.dir_put_end.second}", "%H:%M:%S").time()
                    self.dir_put_time_total2 = (datetime.combine(datetime.now(), self.dir_put_end2) - datetime.combine(datetime.now(), self.dir_put_start2)).total_seconds() / 3600
                else:
                    self.dir_put_end2 = datetime.strptime(f"{int(self.dir_put_end.hour) - 3}:{int(self.dir_put_end.minute)}:{self.dir_put_end.second}", "%H:%M:%S").time()
                    self.dir_put_time_total2 = (datetime.combine(datetime.now(), self.dir_put_end2) - datetime.combine(datetime.now(), self.dir_put_start2)).total_seconds() / 3600
                self.dir_put_time_total2 -= self.total
                if self.dir_put_time_total2 > 0:
                    self.dir_put_uph = int((self.dir_put_total / self.dir_put_time_total2))
                elif self.dir_put_time_total2 < 0:
                    self.dir_put_uph = int((self.dir_put_total / 1))
                else:
                    self.dir_put_uph = 0

            if len(self.rep_pick_times) > 0:
                self.total = 0
                for time in range(len(self.rep_pick_times)):
                    try:
                        self.t1 = datetime.strptime(self.rep_pick_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.rep_pick_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.40:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.rep_pick_start = datetime.strptime(self.rep_pick_times[len(self.rep_pick_times) - 1], "%H:%M:%S").time()
                self.rep_pick_end = datetime.strptime(self.rep_pick_times[0], "%H:%M:%S").time()
                self.rep_pick_time_total = (datetime.combine(datetime.now(), self.rep_pick_end) - datetime.combine(datetime.now(), self.rep_pick_start)).total_seconds() / 3600
                if int(self.rep_pick_start.hour) < 3:
                    self.rep_pick_start2 = datetime.strptime(f"{int(self.rep_pick_start.hour) + 21}:{int(self.rep_pick_start.minute)}:{int(self.rep_pick_start.second)}", "%H:%M:%S").time()
                else:
                    self.rep_pick_start2 = datetime.strptime(f"{int(self.rep_pick_start.hour) - 3}:{int(self.rep_pick_start.minute)}:{int(self.rep_pick_start.second)}", "%H:%M:%S").time()
                if int(self.rep_pick_end.hour) < 3:
                    self.rep_pick_end2 = datetime.strptime(f"{int(self.rep_pick_end.hour) + 21}:{int(self.rep_pick_end.minute)}:{self.rep_pick_end.second}", "%H:%M:%S").time()
                    self.rep_pick_time_total2 = (datetime.combine(datetime.now(), self.rep_pick_end2) - datetime.combine(datetime.now(), self.rep_pick_start2)).total_seconds() / 3600
                else:
                    self.rep_pick_end2 = datetime.strptime(f"{int(self.rep_pick_end.hour) - 3}:{int(self.rep_pick_end.minute)}:{self.rep_pick_end.second}", "%H:%M:%S").time()
                    self.rep_pick_time_total2 = (datetime.combine(datetime.now(),self.rep_pick_end2) - datetime.combine(datetime.now(), self.rep_pick_start2)).total_seconds() / 3600
                self.rep_pick_time_total2 -= self.total
                if self.rep_pick_time_total2 > 0:
                    self.rep_pick_uph = int((self.rep_pick_total / self.rep_pick_time_total2))
                elif self.rep_pick_time_total2 < 0:
                    self.rep_pick_uph = int((self.rep_pick_total / 1))
                else:
                    self.rep_pick_uph = 0

            if len(self.rep_put_times) > 0:
                self.total = 0
                for time in range(len(self.rep_put_times)):
                    try:
                        self.t1 = datetime.strptime(self.rep_put_times[time], "%H:%M:%S").time()
                        self.t2 = datetime.strptime(self.rep_put_times[time + 1], "%H:%M:%S").time()
                        self.t_total = (datetime.combine(datetime.now(), self.t1) - datetime.combine(datetime.now(), self.t2)).total_seconds() / 3600
                        if self.t_total > 0.40:
                            self.total += self.t_total
                    except IndexError:
                        pass
                self.rep_put_start = datetime.strptime(self.rep_put_times[len(self.rep_put_times) - 1], "%H:%M:%S").time()
                self.rep_put_end = datetime.strptime(self.rep_put_times[0], "%H:%M:%S").time()
                self.rep_put_time_total = (datetime.combine(datetime.now(), self.rep_put_end) - datetime.combine(datetime.now(), self.rep_put_start)).total_seconds() / 3600

                if int(self.rep_put_start.hour) < 3:
                    self.rep_put_start2 = datetime.strptime(f"{int(self.rep_put_start.hour) + 21}:{int(self.rep_put_start.minute)}:{int(self.rep_put_start.second)}", "%H:%M:%S").time()
                else:
                    self.rep_put_start2 = datetime.strptime(f"{int(self.rep_put_start.hour) - 3}:{int(self.rep_put_start.minute)}:{int(self.rep_put_start.second)}", "%H:%M:%S").time()
                if int(self.rep_put_end.hour) < 3:
                    self.rep_put_end2 = datetime.strptime(f"{int(self.rep_put_end.hour) + 21}:{int(self.rep_put_end.minute)}:{self.rep_put_end.second}", "%H:%M:%S").time()
                    self.rep_put_time_total2 = (datetime.combine(datetime.now(), self.rep_put_end2) - datetime.combine(datetime.now(), self.rep_put_start2)).total_seconds() / 3600
                else:
                    self.rep_put_end2 = datetime.strptime(f"{int(self.rep_put_end.hour) - 3}:{int(self.rep_put_end.minute)}:{self.rep_put_end.second}", "%H:%M:%S").time()
                    self.rep_put_time_total2 = (datetime.combine(datetime.now(), self.rep_put_end2) - datetime.combine(datetime.now(), self.rep_put_start2)).total_seconds() / 3600
                self.rep_put_time_total2 -= self.total
                if self.rep_put_time_total2 > 0:
                    self.rep_put_uph = int((self.rep_put_total / self.rep_put_time_total2))
                elif self.rep_put_time_total2 < 0:
                    self.rep_put_uph = int((self.rep_put_total / 1))
                else:
                    self.rep_put_uph = 0

            # +---|| DISPLAY INFORMATION IN LOG // Gather Data ||---+ #
            if self.pick_total > 0 and eid not in self.pick_eid_dict:
                self.pick_totals.append(self.pick_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.pick_total} total  OB/S2 picks \n")
                self.debug_queue.append(f"- {self.pick_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.pick_uph} || Goal: {self.pick_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.pick_eid_dict[eid] = self.employee_id_list[eid]
                self.pick_total_dict[eid] = self.pick_total
                self.pick_uph_dict[eid] = self.pick_uph
            if self.acc_pick_total > 0 and eid not in self.acc_pick_eid_dict:
                self.acc_pick_totals.append(self.acc_pick_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.acc_pick_total} total ACC picks \n")
                self.debug_queue.append(f"- {self.acc_pick_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.acc_pick_uph} || Goal: {self.acc_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.acc_pick_eid_dict[eid] = self.employee_id_list[eid]
                self.acc_pick_total_dict[eid] = self.acc_pick_total
                self.acc_pick_uph_dict[eid] = self.acc_pick_uph
            if self.qc_pass_total > 0 and eid not in self.qc_pass_eid_dict:
                self.qc_pass_totals.append(self.qc_pass_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.qc_pass_total} total QC PASSES \n")
                self.debug_queue.append(f"- {self.qc_pass_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.qc_pass_uph} || Goal: {self.qc_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.qc_pass_eid_dict[eid] = self.employee_id_list[eid]
                self.qc_pass_total_dict[eid] = self.qc_pass_total
                self.qc_pass_uph_dict[eid] = self.qc_pass_uph
            if self.dir_put_total > 0 and eid not in self.dir_put_eid_dict:
                self.dir_put_totals.append(self.dir_put_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.dir_put_total} total IB putaways \n")
                self.debug_queue.append(f"- {self.dir_put_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.dir_put_uph} || Goal: {self.ib_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.dir_put_eid_dict[eid] = self.employee_id_list[eid]
                self.dir_put_total_dict[eid] = self.dir_put_total
                self.dir_put_uph_dict[eid] = self.dir_put_uph
            if self.rep_pick_total > 0 and eid not in self.rep_pick_eid_dict:
                self.rep_pick_totals.append(self.rep_pick_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.rep_pick_total} total Replen picks \n")
                self.debug_queue.append(f"- {self.rep_pick_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.rep_pick_uph} || Goal: {self.replen_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.rep_pick_eid_dict[eid] = self.employee_id_list[eid]
                self.rep_pick_total_dict[eid] = self.rep_pick_total
                self.rep_pick_uph_dict[eid] = self.rep_pick_uph
            if self.rep_put_total > 0 and eid not in self.rep_put_eid_dict:
                self.rep_put_totals.append(self.rep_put_total)
                self.debug_queue.append(f"- ID: {self.employee_id_list[eid]} -"
                                        f"- {self.rep_put_total} total Replens put away \n")
                self.debug_queue.append(f"- {self.rep_put_time_total2} hours \n")
                self.debug_queue.append(f"- UPH: {self.rep_put_uph} || Goal: {self.replen_text_var.get()} \n")
                self.load_bar["value"] += 1
                self.rep_put_eid_dict[eid] = self.employee_id_list[eid]
                self.rep_put_total_dict[eid] = self.rep_put_total
                self.rep_put_uph_dict[eid] = self.rep_put_uph

        self.debug_queue.append("Generating new excel spreadsheet... \n")

        if self.load_bar["value"] < 80:
            self.load_bar["value"] = 80

        Thread(target=self.write_to_spreadsheet, daemon=True).start()
        os.remove(self.csv_name)

    def create_new_spreadsheet(self):
        self.workbook = xlsxwriter.Workbook(f"{self.save_file_text_var.get()}.xlsx")
        self.worksheet = self.workbook.add_worksheet()

    def write_to_spreadsheet(self):
        """                   LAYOUT:
                          :: ACTIVITY ::
        :: ID :: TOTAL :: TOTAL_AVERAGE :: UPH :: UPH_GOAL ::
        """
        self.worksheet.set_column("A:A", 8)
        self.worksheet.set_column("B:B", 5)
        self.worksheet.set_column("C:C", 15)
        self.worksheet.set_column("D:D", 4)
        self.worksheet.set_column("E:E", 5)
        self.worksheet.set_column("G:G", 8)
        self.worksheet.set_column("H:H", 5)
        self.worksheet.set_column("I:I", 15)
        self.worksheet.set_column("J:J", 4)
        self.worksheet.set_column("K:K", 5)
        self.row_position = 0
        self.row_position2 = 0
        self.pick_first_row = 0
        self.acc_pick_first_row = 0
        self.qc_pass_first_row = 0
        self.dir_put_first_row = 0
        self.rep_pick_first_row = 0
        self.rep_put_first_row = 0

        ## STYLES ##

        # TITLE YELLOW #
        self.center_title_style = self.workbook.add_format()
        self.center_title_style.set_bold()
        self.center_title_style.set_align("center")
        self.center_title_style.set_bg_color("yellow")
        self.center_title_style.set_top(2)
        self.center_title_style.set_bottom(2)
        self.center_title_style.set_border_color("black")

        self.title_left_style = self.workbook.add_format()
        self.title_left_style.set_bg_color("yellow")
        self.title_left_style.set_left(2)
        self.title_left_style.set_top(2)
        self.title_left_style.set_bottom(2)
        self.title_left_style.set_border_color("black")

        self.title_right_style = self.workbook.add_format()
        self.title_right_style.set_bg_color("yellow")
        self.title_right_style.set_right(2)
        self.title_right_style.set_top(2)
        self.title_right_style.set_bottom(2)
        self.title_right_style.set_border_color("black")

        # MERGED CELLS #
        self.merged_style = self.workbook.add_format()
        self.merged_style.set_border(2)
        self.merged_style.set_bold()
        self.merged_style.set_border_color("black")
        self.merged_style.set_align("center")
        self.merged_style.set_align("vcenter")
        self.merged_style.set_rotation(30)

        self.merged_false_style = self.workbook.add_format()
        self.merged_false_style.set_border(2)
        self.merged_false_style.set_bold()
        self.merged_false_style.set_border_color("black")
        self.merged_false_style.set_align("center")

        # HEADER INFO #
        self.header_info_style = self.workbook.add_format()
        self.header_info_style.set_border(2)
        self.header_info_style.set_border_color("black")
        self.header_info_style.set_align("center")

        # HEADER CYAN #
        self.header_center_style = self.workbook.add_format()
        self.header_center_style.set_bold()
        self.header_center_style.set_align("center")
        self.header_center_style.set_bg_color("cyan")
        self.header_center_style.set_top(2)
        self.header_center_style.set_bottom(2)
        self.header_center_style.set_border_color("black")

        self.header_left_style = self.workbook.add_format()
        self.header_left_style.set_bg_color("cyan")
        self.header_left_style.set_top(2)
        self.header_left_style.set_bottom(2)
        self.header_left_style.set_left(2)
        self.header_left_style.set_border_color("black")

        self.header_right_style = self.workbook.add_format()
        self.header_right_style.set_bg_color("cyan")
        self.header_right_style.set_top(2)
        self.header_right_style.set_bottom(2)
        self.header_right_style.set_right(2)
        self.header_right_style.set_border_color("black")

        # GOAL STYLES
        self.red_cell_style = self.workbook.add_format()
        self.red_cell_style.set_bg_color("red")

        self.yellow_cell_style = self.workbook.add_format()
        self.yellow_cell_style.set_bg_color("yellow")

        self.green_cell_style = self.workbook.add_format()
        self.green_cell_style.set_bg_color("green")

        self.white_cell_style = self.workbook.add_format()
        self.white_cell_style.set_bg_color("white")

        ## WRITE TO FILE ##

        # CREATE TITLE WITH DATE
        self.worksheet.write(self.row_position, 0, None, self.title_left_style)
        self.worksheet.write(self.row_position, 1, None, self.center_title_style)
        self.worksheet.write(self.row_position, 2, None, self.center_title_style)
        self.worksheet.write(self.row_position, 3, None, self.center_title_style)
        self.worksheet.write(self.row_position, 4, None, self.center_title_style)
        self.worksheet.write(self.row_position, 5, self.transactions.iat[2, 3], self.center_title_style)
        self.worksheet.write(self.row_position, 6, None, self.center_title_style)
        self.worksheet.write(self.row_position, 7, None, self.center_title_style)
        self.worksheet.write(self.row_position, 8, None, self.center_title_style)
        self.worksheet.write(self.row_position, 9, None, self.center_title_style)
        self.worksheet.write(self.row_position, 10, None, self.title_right_style)
        self.row_position += 1
        self.row_position2 += 1
        self.debug_queue.append("Header with date created... \n")
        self.load_bar["value"] = 82

        # CREATE OB/S2 HEADER
        self.worksheet.write(self.row_position, 0, None, self.header_left_style)
        self.worksheet.write(self.row_position, 1, None, self.header_center_style)
        self.worksheet.write(self.row_position, 2, "OB/S2", self.header_center_style)
        self.worksheet.write(self.row_position, 3, None, self.header_center_style)
        self.worksheet.write(self.row_position, 4, None, self.header_right_style)
        self.row_position += 1
        self.worksheet.write(self.row_position, 0, "ID", self.header_info_style)
        self.worksheet.write(self.row_position, 1, "Total", self.header_info_style)
        self.worksheet.write(self.row_position, 2, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position, 3, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position, 4, "Goal", self.header_info_style)
        self.row_position += 1
        self.debug_queue.append(" OB/S2 Header created... \n")
        self.load_bar["value"] = 84

        self.pick_first_row = self.row_position

        # GENERATE OB/S2 DATA
        for key in self.pick_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.pick_uph_dict[key] >= int(self.pick_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.pick_uph_dict[key] >= (int(self.pick_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.pick_uph_dict[key] < (int(self.pick_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style
            self.worksheet.write(self.row_position, 0, self.pick_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 1, self.pick_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 3, self.pick_uph_dict[key], self.goal_style)
            self.row_position += 1
        if (self.row_position - 1) - self.pick_first_row > 1:
            self.worksheet.merge_range(self.pick_first_row, 2, self.row_position - 1, 2, int(sum(self.pick_totals) // len(self.pick_totals)), self.merged_style)
            self.worksheet.merge_range(self.pick_first_row, 4, self.row_position - 1, 4, self.pick_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.pick_first_row, 2, int(sum(self.pick_totals) // len(self.pick_totals)), self.merged_false_style)
            self.worksheet.write(self.pick_first_row, 4, self.pick_text_var.get(), self.merged_false_style)
        self.debug_queue.append(" OB/S2 data writing complete... \n")
        self.load_bar["value"] = 86

        # CREATE QC PASS HEADER
        self.worksheet.write(self.row_position2, 6, None, self.header_left_style)
        self.worksheet.write(self.row_position2, 7, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 8, "QC", self.header_center_style)
        self.worksheet.write(self.row_position2, 9, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 10, None, self.header_right_style)
        self.row_position2 += 1
        self.worksheet.write(self.row_position2, 6, "ID", self.header_info_style)
        self.worksheet.write(self.row_position2, 7, "Total", self.header_info_style)
        self.worksheet.write(self.row_position2, 8, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position2, 9, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position2, 10, "Goal", self.header_info_style)
        self.row_position2 += 1
        self.debug_queue.append(" QC header created... \n")
        self.load_bar["value"] = 88

        self.qc_pass_first_row = self.row_position2

        # GENERATE QC DATA
        for key in self.qc_pass_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.qc_pass_uph_dict[key] >= int(self.qc_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.qc_pass_uph_dict[key] >= (int(self.qc_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.qc_pass_uph_dict[key] < (int(self.qc_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style
            self.worksheet.write(self.row_position2, 6, self.qc_pass_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 7, self.qc_pass_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 9, self.qc_pass_uph_dict[key], self.goal_style)
            self.row_position2 += 1
        if (self.row_position2 - 1) - self.qc_pass_first_row > 1:
            self.worksheet.merge_range(self.qc_pass_first_row, 8, self.row_position2 - 1, 8, int(sum(self.qc_pass_totals) / len(self.qc_pass_totals)), self.merged_style)
            self.worksheet.merge_range(self.qc_pass_first_row, 10, self.row_position2 - 1, 10, self.qc_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.qc_pass_first_row, 8, int(sum(self.qc_pass_totals) / len(self.qc_pass_totals)), self.merged_false_style)
            self.worksheet.write(self.qc_pass_first_row, 10, self.qc_text_var.get(), self.merged_false_style)
        self.debug_queue.append(" QC data writing complete... \n")
        self.load_bar["value"] = 90

        # CREATE IB/RTW HEADER
        self.worksheet.write(self.row_position, 0, None, self.header_left_style)
        self.worksheet.write(self.row_position, 1, None, self.header_center_style)
        self.worksheet.write(self.row_position, 2, "IB/RTW", self.header_center_style)
        self.worksheet.write(self.row_position, 3, None, self.header_center_style)
        self.worksheet.write(self.row_position, 4, None, self.header_right_style)
        self.row_position += 1
        self.worksheet.write(self.row_position, 0, "ID", self.header_info_style)
        self.worksheet.write(self.row_position, 1, "Total", self.header_info_style)
        self.worksheet.write(self.row_position, 2, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position, 3, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position, 4, "Goal", self.header_info_style)
        self.row_position += 1
        self.debug_queue.append(" IB/RTW header created... \n")
        self.load_bar["value"] = 92

        self.dir_put_first_row = self.row_position

        # GENERATE IB/RTW DATA
        for key in self.dir_put_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.dir_put_uph_dict[key] >= int(self.ib_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.dir_put_uph_dict[key] >= (int(self.ib_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.dir_put_uph_dict[key] < (int(self.ib_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style
            self.worksheet.write(self.row_position, 0, self.dir_put_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 1, self.dir_put_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 3, self.dir_put_uph_dict[key], self.goal_style)
            self.row_position += 1
        if (self.row_position - 1) - self.dir_put_first_row > 1:
            self.worksheet.merge_range(self.dir_put_first_row, 2, self.row_position - 1, 2, int(sum(self.dir_put_totals) / len(self.dir_put_totals)), self.merged_style)
            self.worksheet.merge_range(self.dir_put_first_row, 4, self.row_position - 1, 4, self.ib_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.dir_put_first_row, 2, int(sum(self.dir_put_totals) / len(self.dir_put_totals)), self.merged_false_style)
            self.worksheet.write(self.dir_put_first_row, 4, self.ib_text_var.get(), self.merged_false_style)
        self.debug_queue.append(" IB/RTW data writing complete... \n")
        self.load_bar["value"] = 93

        # CREATE ACCESSORIES HEADER
        self.worksheet.write(self.row_position2, 6, None, self.header_left_style)
        self.worksheet.write(self.row_position2, 7, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 8, "ACCESSORIES", self.header_center_style)
        self.worksheet.write(self.row_position2, 9, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 10, None, self.header_right_style)
        self.row_position2 += 1
        self.worksheet.write(self.row_position2, 6, "ID", self.header_info_style)
        self.worksheet.write(self.row_position2, 7, "Total", self.header_info_style)
        self.worksheet.write(self.row_position2, 8, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position2, 9, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position2, 10, "Goal", self.header_info_style)
        self.row_position2 += 1
        self.debug_queue.append(" Accessories header created... \n")
        self.load_bar["value"] = 94

        self.acc_pick_first_row = self.row_position2

        # GENERATE ACCESSORIES DATA
        for key in self.acc_pick_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.acc_pick_uph_dict[key] >= int(self.acc_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.acc_pick_uph_dict[key] >= (int(self.acc_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.acc_pick_uph_dict[key] < (int(self.acc_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style
            self.worksheet.write(self.row_position2, 6, self.acc_pick_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 7, self.acc_pick_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 9, self.acc_pick_uph_dict[key], self.goal_style)
            self.row_position2 += 1
        if (self.row_position2 - 1) - self.acc_pick_first_row > 1:
            self.worksheet.merge_range(self.acc_pick_first_row, 8, self.row_position2 - 1, 8, int(sum(self.acc_pick_totals) / len(self.acc_pick_totals)), self.merged_style)
            self.worksheet.merge_range(self.acc_pick_first_row, 10, self.row_position2 - 1, 10, self.acc_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.acc_pick_first_row, 8, int(sum(self.acc_pick_totals) / len(self.acc_pick_totals)), self.merged_false_style)
            self.worksheet.write(self.acc_pick_first_row, 10, self.acc_text_var.get(), self.merged_false_style)
        self.debug_queue.append("Accessories data writing complete... \n")
        self.load_bar["value"] = 95

        # CREATE REPLEN PICKS HEADER
        self.worksheet.write(self.row_position, 0, None, self.header_left_style)
        self.worksheet.write(self.row_position, 1, None, self.header_center_style)
        self.worksheet.write(self.row_position, 2, "REPLEN PICKING", self.header_center_style)
        self.worksheet.write(self.row_position, 3, None, self.header_center_style)
        self.worksheet.write(self.row_position, 4, None, self.header_right_style)
        self.row_position += 1
        self.worksheet.write(self.row_position, 0, "ID", self.header_info_style)
        self.worksheet.write(self.row_position, 1, "Total", self.header_info_style)
        self.worksheet.write(self.row_position, 2, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position, 3, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position, 4, "Goal", self.header_info_style)
        self.row_position += 1
        self.debug_queue.append("Replen pick header created... \n")
        self.load_bar["value"] = 96

        self.rep_pick_first_row = self.row_position

        # GENERATE REPLEN PICKS DATA
        for key in self.rep_pick_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.rep_pick_uph_dict[key] >= int(self.replen_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.rep_pick_uph_dict[key] >= (int(self.replen_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.rep_pick_uph_dict[key] < (int(self.replen_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style

            self.worksheet.write(self.row_position, 0, self.rep_pick_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 1, self.rep_pick_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position, 3, self.rep_pick_uph_dict[key], self.goal_style)
            self.row_position += 1
        if (self.row_position - 1) - self.rep_pick_first_row > 1:
            self.worksheet.merge_range(self.rep_pick_first_row, 2, self.row_position - 1, 2, int(sum(self.rep_pick_totals) / len(self.rep_pick_totals)), self.merged_style)
            self.worksheet.merge_range(self.rep_pick_first_row, 4, self.row_position - 1, 4, self.replen_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.rep_pick_first_row, 2, int(sum(self.rep_pick_totals) / len(self.rep_pick_totals)), self.merged_false_style)
            self.worksheet.write(self.rep_pick_first_row, 4, self.replen_text_var.get(), self.merged_false_style)
        self.debug_queue.append("Replen pick data writing complete \n")
        self.load_bar["value"] = 97

        # CREATE REPLEN PUTAWAY HEADER
        self.worksheet.write(self.row_position2, 6, None, self.header_left_style)
        self.worksheet.write(self.row_position2, 7, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 8, "REPLEN PUTAWAY", self.header_center_style)
        self.worksheet.write(self.row_position2, 9, None, self.header_center_style)
        self.worksheet.write(self.row_position2, 10, None, self.header_right_style)
        self.row_position2 += 1
        self.worksheet.write(self.row_position2, 6, "ID", self.header_info_style)
        self.worksheet.write(self.row_position2, 7, "Total", self.header_info_style)
        self.worksheet.write(self.row_position2, 8, "Total Average", self.header_info_style)
        self.worksheet.write(self.row_position2, 9, "UPH", self.header_info_style)
        self.worksheet.write(self.row_position2, 10, "Goal", self.header_info_style)
        self.row_position2 += 1
        self.debug_queue.append("replen putaway header created... \n")
        self.load_bar["value"] = 98

        self.rep_put_first_row = self.row_position2

        # GENERATE REPLEN PUTAWAY DATA
        for key in self.rep_put_eid_dict.keys():
            if self.color_bool_var.get() is True:
                if self.rep_put_uph_dict[key] >= int(self.replen_text_var.get()):
                    self.goal_style = self.green_cell_style
                elif self.rep_put_uph_dict[key] >= (int(self.replen_text_var.get()) - 10):
                    self.goal_style = self.yellow_cell_style
                elif self.rep_put_uph_dict[key] < (int(self.replen_text_var.get()) - 10):
                    self.goal_style = self.red_cell_style
            elif self.color_bool_var.get() is False:
                self.goal_style = self.white_cell_style
            else:
                self.goal_style = self.white_cell_style
            self.worksheet.write(self.row_position2, 6, self.rep_put_eid_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 7, self.rep_put_total_dict[key], self.goal_style)
            self.worksheet.write(self.row_position2, 9, self.rep_put_uph_dict[key], self.goal_style)
            self.row_position2 += 1
        if (self.row_position2 - 1) - self.rep_put_first_row > 1:
            self.worksheet.merge_range(self.rep_put_first_row, 8, self.row_position2 - 1, 8, int(sum(self.rep_put_totals) / len(self.rep_put_totals)), self.merged_style)
            self.worksheet.merge_range(self.rep_put_first_row, 10, self.row_position2 - 1, 10, self.replen_text_var.get(), self.merged_style)
        else:
            self.worksheet.write(self.rep_put_first_row, 8, int(sum(self.rep_put_totals) / len(self.rep_put_totals)), self.merged_false_style)
            self.worksheet.write(self.rep_put_first_row, 10, self.replen_text_var.get(), self.merged_false_style)
        self.debug_queue.append("Replen putaway data writing complete \n")
        self.load_bar["value"] = 99

        self.workbook.close()
        self.save_title = self.save_file_text_var.get()
        try:
            shutil.move(os.path.abspath(f"{self.save_title}.xlsx"), Path(self.save_text_var.get()))
        except shutil.Error:
            self.debug_queue.append(f"Error: {self.save_title}.xlsx already exists...\n")
            self.debug_queue.append("Creating new name for file...\n")
            self.save_title = random.randint(100000, 999999)
            os.rename(f"{self.save_file_text_var.get()}.xlsx", f"{self.save_title}.xlsx")
            shutil.move(os.path.abspath(f"{self.save_title}.xlsx"), Path(self.save_text_var.get()))

        self.pdf_checkbox.config(state=NORMAL)
        if self.pdf_bool_var.get() is True:
            # EXPORT AS PDF FILE
            self.debug_queue.append("Converting to PDF format... \n")
            self.path_to_workbook = os.path.normpath(f"{self.save_text_var.get()}\\{self.save_title}.xlsx")
            self.excel = client.Dispatch("Excel.Application")
            self.open_workbook = self.excel.Workbooks.Open(self.path_to_workbook)
            self.open_workbook.Worksheets[0].ExportAsFixedFormat(0, os.path.normpath(f"{self.save_text_var.get()}\\{self.save_title}.pdf"))
            self.open_workbook.Close(True)
            os.remove(os.path.normpath(f"{self.save_text_var.get()}\\{self.save_title}.xlsx"))
            self.debug_queue.append(f"File: {self.save_title}.pdf ---- \n"
                                    f"---- Saved to: {self.save_text_var.get()} \n")
            self.load_bar["value"] = 100
            self.pdf_checkbox.config(state=DISABLED)
            self.excel.Quit()
        elif self.pdf_bool_var.get() is False:
            # KEEP AS XLSX FILE
            self.debug_queue.append(f"File: {self.save_title}.xlsx ---- \n"
                                    f"---- Saved to: {self.save_text_var.get()} \n")
            self.load_bar["value"] = 100
            self.pdf_checkbox.config(state=DISABLED)
        self.run_btn.config(state=DISABLED)
        self.color_checkbox.config(state=NORMAL)
        self.pdf_checkbox.config(state=NORMAL)
        self.save_file_entry.config(state=NORMAL)
        self.browse_btn.config(state=NORMAL)
        self.save_location_btn.config(state=NORMAL)
        self.source_entry.delete("0", END)


if __name__ == "__main__":
    root = ttk.Window()
    run = Run(root)
    root.mainloop()
