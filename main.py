import customtkinter, csv, time, threading, sys
from tkinter import messagebox, ttk, END, StringVar
from printermanager import PrinterManager
from variablecontrollers import Version, SNMPVar, PrinterVar, CSVFieldnames, SPClass, CSVCheck, Paths
from csvmanagement import GenerateCSV, CSVChecker
from sharepointmanagement import SharePointHandler
from os import system
from difflib import Differ
from win32.win32api import GetSystemMetrics
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from webbrowser import open_new_tab

# from PIL import Image (Might be needed in the future, will be removed if it isn't needed)

'''
    Information:
    This the main script
    
    
    
    Improvements possible:
    I have scripts already made that only load into memory and then skips having any local files/folders
    other than the installer exe for updating the software automatically.
'''

# Application Configurator manages version checking and has a function for automatic change logs (very experimental)
class AppHandler():
        
    def close_and_open_window(self, close, open):
        close.withdraw()
        open.deiconify()

    def generate_change_log(self, context, sp_rel_path, old_file_path, new_file_path, out_file_path):
        with open(old_file_path, 'r', encoding='utf-8') as old_file, open(new_file_path, 'r', encoding='utf-8') as new_file, open(out_file_path, 'w', encoding='utf-8') as out_file:
            old_lines = old_file.readlines()
            new_lines = new_file.readlines()

            differ = Differ()
            diff = list(differ.compare(old_lines, new_lines))

            old_vars = {}
            new_vars = {}

            for line in diff:
                line = line.strip()
                if line.startswith('- '):
                    if ' = ' in line:
                        var, val = line[2:].split(' = ', 1)
                        old_vars[var] = val
                    else:
                        var = line[2:]
                        old_vars[var] = None
                elif line.startswith('+ '):
                    if ' = ' in line:
                        var, val = line[2:].split(' = ', 1)
                        new_vars[var] = val
                    else:
                        var = line[2:]
                        new_vars[var] = None
            for var in set(old_vars.keys()) | set(new_vars.keys()):
                old_val = old_vars.get(var)
                new_val = new_vars.get(var)
                if old_val is None:
                    out_file.write(f'Added {var} {new_val}\n')
                elif new_val is None:
                    out_file.write(f'Removed {var} {old_val}\n')
                elif old_val != new_val:
                    out_file.write(f'Changed {var} from {old_val} to {new_val}\n')
        
        SharePointHandler().upload_item(context, out_file_path, sp_rel_path)


    def version_check(self, function_context, SharePoint_file_path, local_file_path, version, version_fieldnames):
        SharePointHandler().download_item(function_context, SharePoint_file_path, local_file_path)
        with open(local_file_path) as version_list:
            f1_reader = csv.DictReader(version_list)
            for row in f1_reader:
                if row[version_fieldnames[0]] != version_fieldnames[0]:
                    if float(row[version_fieldnames[0]]) > version:
                        return [False, row[version_fieldnames[0]], row[version_fieldnames[1]]]
                    elif float(row[version_fieldnames[0]]) > (version + 0.5):
                        return [False, row[version_fieldnames[0]], "True"]
                    else:
                        return [True, row[version_fieldnames[0]], row[version_fieldnames[1]]]
    
    def update_application(self, context, installer_rel_path, installer_path):
        SharePointHandler().download_item(context, installer_rel_path, installer_path)
        exit_timer = threading.Thread(target=self.wait_for_exit)
        exit_timer.start()
        system(installer_path)
    
    def wait_for_exit(self):
        timer = 5
        while timer > 0:
            time.sleep(1)
            timer = timer -1
        sys.exit(1)
# END

# Admin login TopLevel window
class AdminLogin():
    def __init__(self):
        self.admin = customtkinter.CTkToplevel()

        self.admin.geometry("500x300")
        self.admin.resizable(0, 0)
        self.admin.title("PTM Login")

        self.main_frame = customtkinter.CTkFrame(self.admin, height=300, width=500, fg_color="transparent")
        self.main_frame.pack_propagate(0) # Pack_propagate prevents the window resizing to match the widgets
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.pack(fill="both", expand="true")

        self.admin_title = customtkinter.CTkLabel(self.main_frame, text="Printer Monitor login", font=customtkinter.CTkFont(size=40, weight="bold"))
        self.admin_title.grid(row=0, column=0, padx=5, pady=25, sticky="ew")

        self.login_frame = customtkinter.CTkFrame(self.main_frame)
        self.login_frame.grid(row=1, column=0, padx=5, pady=2, sticky="ew")

        self.label_user = customtkinter.CTkLabel(self.login_frame, text="Outlook Email:", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_user.grid(row=0, column=0, padx=5, pady=5)

        self.entry_user = customtkinter.CTkEntry(self.login_frame, width=335, font=customtkinter.CTkFont(size=20, weight="bold"))
        self.entry_user.grid(row=0, column=1, padx=1, pady=5)

        self.user_check_var = StringVar(self.login_frame,"off")
        self.user_checkbox = customtkinter.CTkCheckBox(self.login_frame, text="Remember Outlook",
                                        variable=self.user_check_var, onvalue="on", offvalue="off",
                                        font=customtkinter.CTkFont(size=20, weight="bold"))
        self.user_checkbox.grid(row=1, column=1, padx=2, pady=5)

        self.label_pw = customtkinter.CTkLabel(self.login_frame, text="Password:", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.label_pw.grid(row=2, column=0, padx=5, pady=5)

        self.entry_pw = customtkinter.CTkEntry(self.login_frame, width=335, font=customtkinter.CTkFont(size=20, weight="bold"))
        self.entry_pw.configure(show="*")
        self.entry_pw.bind("<Return>", lambda event : threading.Thread(target=getlogin).start())
        self.entry_pw.grid(row=2, column=1, padx=1, pady=5)

        self.pw_check_var = StringVar(self.login_frame,"off")
        self.pw_checkbox = customtkinter.CTkCheckBox(self.login_frame, text="Show password", 
                                        command=lambda: threading.Thread(target=toggle_password).start(), variable=self.pw_check_var,
                                        onvalue="on", offvalue="off", font=customtkinter.CTkFont(size=20,
                                                                                        weight="bold"))
        self.pw_checkbox.grid(row=3, column=1, padx=2, pady=5)

        self.label_error = customtkinter.CTkLabel(self.login_frame, text="Incorrect credentials", font=customtkinter.CTkFont(size=12, weight="bold"), text_color="#FF9494")

        self.login_button = customtkinter.CTkButton(self.login_frame, width=20, text="Login", command=lambda: threading.Thread(target=getlogin).start(), font=customtkinter.CTkFont(size=20, weight="bold"))
        self.login_button.grid(row=4, column=1, padx=2, pady=5)

        def getlogin():
            self.login_button.configure(state="disabled", text="Logging in..")
            self.label_error.grid_forget()
            validation = validate(self.entry_user.get(), self.entry_pw.get())
            if validation == True:
                threading.Thread(target=check_for_csv).start()
                threading.Thread(target=remember_mail).start()
                AppHandler().close_and_open_window(self.admin, root)     
            else:
                self.login_button.configure(state="enabled", text="Login")
                self.label_error.grid(row=4, column=0, padx=2, pady=5)
                time.sleep(2)
                self.label_error.grid_forget()
        
        def validate(username, password):
            # Checks SharePoint for a matching username to email
            try:
                SPClass.context = ClientContext(SPClass.team_site_url).with_credentials(UserCredential(username, password))
                SPClass.context.web.get().execute_query()
                return True
            except:
                SPClass.context = None
                return False

        def toggle_password():
            pwe = self.entry_pw
            if  pwe.cget('show') == '':
                pwe.configure(show='*')
            else:
                pwe.configure(show='')

        def show_rem_mail(): # Inputs rememberd mail (If it exists) into user entry
            try:
                with open (Paths.local_config_path, 'r') as local_config_list:
                    outlook_reader = csv.DictReader(local_config_list, fieldnames = CSVFieldnames.local_config_fieldnames)
                    for row in outlook_reader:
                        if row['Outlook'] != 'Outlook' and row['Outlook'] != '':
                            self.entry_user.insert(0, row[CSVFieldnames.local_config_fieldnames[0]])
                            onvar = StringVar(self.login_frame,"on")
                            self.user_checkbox.configure(variable=onvar)
            except:
                print("No email has been saved")

        def remember_mail():
            if str(self.user_checkbox.get()) == "on":
                with open (Paths.local_config_path, 'w') as local_config_list:
                    outlook_writer = csv.DictWriter(local_config_list, fieldnames = CSVFieldnames.local_config_fieldnames)            
                    outlook_writer.writeheader()
                    row = {CSVFieldnames.local_config_fieldnames[0]: self.entry_user.get()}
                    outlook_writer.writerow(row)       
            self.entry_user.delete(0, END)
            self.entry_pw.delete(0, END)

         # Checks if the CSVs have been created
        def check_for_csv():
            threading.Thread(target=CSVChecker().ptmcsvcheck).start()
            threading.Thread(target=CSVChecker().ptmscsvcheck).start()
            threading.Thread(target=CSVChecker().versioncsvchecker).start()

            GenerateCSV()
            threading.Thread(target=check_for_updates).start()

        def check_for_updates():
            updated = AppHandler().version_check(SPClass.context, SPClass.ptm_verison_file_rel_path, Paths.version_path, Version.version, CSVFieldnames.version_fieldnames)
            if updated[0] == False:
                update_box = messagebox.askquestion('Version out of date!','Old version: {}\nNew version: {}\nDo you want to update? (Recommended)'.format(Version.version, updated[1]),
                                            icon='warning')
                if update_box == 'yes':
                    AppHandler().update_application(SPClass.context, SPClass.ptm_installer_rel_path, Paths.installer_path)
                elif update_box == 'no' and updated[2] == "True":
                    messagebox.showerror("Critical update", "The program is in critical need of an update.\nIt will need to update now.")
                    AppHandler().update_application(SPClass.context, SPClass.ptm_installer_rel_path, Paths.installer_path)
                else:
                    messagebox.showinfo("Update canceled", "The update is non-critical it will now run without the update.\nUpdate software when possible.")
                    Version.version_checked = True
            else:
                Version.version_checked = True

        threading.Thread(target=show_rem_mail).start()
        self.admin.protocol("WM_DELETE_WINDOW", self.quit_app)

    def quit_app(self):
        sys.exit(1)
# END

# PTM dashboard
class PTM(customtkinter.CTk):
    def __init__(self, *args, **kwargs):
        customtkinter.CTk.__init__(self, *args, **kwargs)
        self.my_timer = 120
        self.initial_load = False
        
        # Setting set resolution for scaling to be correct
        self.width = GetSystemMetrics(0)
        self.height = GetSystemMetrics(1)

        self.geometry(str(self.width)+"x"+str(self.height))  # Sets window size to monitor pixel width and height
        self.title("Printer Monitor")

        # Attributes. e.g. set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Create navigation frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(6, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="Printer\nMonitor",
                                                            compound="left", font=customtkinter.CTkFont(size=60, weight="bold"))
        self.navigation_frame_label.grid(row=1, column=0, padx=50, pady=(40,40))

        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="🏠 Home",
                                                   fg_color="transparent", text_color=("gray90", "gray90"), hover_color=("gray30", "gray30"),
                                                   command=self.home_button_event, font=customtkinter.CTkFont(size=25, weight="bold"))
        self.home_button.grid(row=2, column=0, sticky="ew")

        self.ptm_frame_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="🖶 Printers",
                                                   fg_color="transparent", text_color=("gray90", "gray90"), hover_color=("gray30", "gray30"),
                                                   command=self.ptm_button_event, font=customtkinter.CTkFont(size=25, weight="bold"))
        self.ptm_frame_button.grid(row=3, column=0, sticky="ew")

        self.search_frame_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="🔎 Search",
                                                   fg_color="transparent", text_color=("gray90", "gray90"), hover_color=("gray30", "gray30"),
                                                   command=self.search_button_event, font=customtkinter.CTkFont(size=25, weight="bold"))
        self.search_frame_button.grid(row=4, column=0, sticky="ew")

        self.configure_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="⚙️ Settings",
                                                      fg_color="transparent",
                                                      command=self.configure_button_event, font=customtkinter.CTkFont(size=25, weight="bold"))
        self.configure_button.grid(row=5, column=0, sticky="ew")

        self.update_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="⇣ Update",
                                                      fg_color="transparent", text_color=("#ADD8E6","#ADD8E6"), hover_color=("gray30", "gray30"),
                                                      command=self.update_button_event, font=customtkinter.CTkFont(size=15, weight="bold"))
        self.update_button.grid(row=7, column=0, sticky="ew")

        self.update_button.grid_forget()

        self.scaling_label = customtkinter.CTkLabel(self.navigation_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=8, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                                    variable=StringVar(self,"100%"), command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=9, column=0, padx=20, pady=(10, 20))

        # Creates home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(2, weight=1)

        self.home_info_frame = customtkinter.CTkFrame(self.home_frame, corner_radius=0, fg_color="transparent")
        self.home_info_frame.pack()

        self.home_label = customtkinter.CTkLabel(self.home_info_frame, text="Printer Monitor",
                                                    font=customtkinter.CTkFont(size=65, weight="bold"))
        self.home_label.grid(row=0, column=1, padx=50, pady=120)

        self.home_description_label = customtkinter.CTkLabel(self.home_info_frame, text="The definitive tool to keep track of toner levels, toner types and printer locations!\nTool-Tip: You can click the printer names to open their dashboard in your default browser (Firefox is recommended)",
                                                    font=customtkinter.CTkFont(size=25, weight="bold"))
        self.home_description_label.grid(row=1, column=1, padx=50)
        
        # Creates printer monitor frame
        self.ptm_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.ptm_frame.grid_columnconfigure(0, weight=1)
        self.ptm_frame.grid_rowconfigure(0, weight=1)

        # Frame for printer treeview
        self.treeview_frame = customtkinter.CTkFrame(self.ptm_frame, corner_radius=0, fg_color="transparent")
        self.treeview_frame.grid(row=0, column=0, padx=5, pady=10, sticky="ewns")

        self.stw = ttk.Style()
        self.stw.theme_use('default')
        self.stw.configure("mystyle.Treeview", highlightthickness=0, bd=0, font=('Roboto', 17, 'bold'), rowheight=40, background="#4b4b4b", foreground="white") # Modify the font of the body
        self.stw.map("mystyle.Treeview.Item", background=[('selected', '#ADD8E6')], foreground=[('selected', '#383838')])
        self.stw.configure("mystyle.Treeview.Heading", font=('Roboto', 17, 'bold'), background="#595959", foreground="white")
        self.stw.map("mystyle.Treeview.Heading", background=[('active', '#636363')])
        self.stw.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})]) # Remove the borders
        
        # This is the printer treeview.
        self.tv1 = ttk.Treeview(self.treeview_frame, style="mystyle.Treeview")
        self.column_list = [CSVFieldnames.fieldnames[0], CSVFieldnames.fieldnames[1], CSVFieldnames.fieldnames[2]]
        self.tv1['columns'] = self.column_list
        self.tv1["show"] = "headings"  # Removes empty column and tree opener (+/- symbol)
        for column in self.column_list:
            self.tv1.heading(column, text=column)
            self.tv1.column(column, width=50)
        self.tv1.bind("<<TreeviewSelect>>", self.open_treeview_item)
        self.tv1.place(relheight=1, relwidth=0.995)
        self.treescroll = customtkinter.CTkScrollbar(self.treeview_frame)
        self.treescroll.configure(command=self.tv1.yview)
        self.tv1.configure(yscrollcommand=self.treescroll.set)
        self.treescroll.pack(side="right", fill="y")

        # Creates search frame
        self.search_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.search_frame.grid_columnconfigure(0, weight=1)
        self.search_frame.grid_rowconfigure(0, weight=1)
        self.search_frame.grid_rowconfigure(1, weight=2)

        # Creates the search config frame
        self.search_configuration_frame = customtkinter.CTkFrame(self.search_frame, fg_color="transparent")
        self.search_configuration_frame.grid(row=0, column=0)

        # Creates search entry frame
        self.search_entry_frame = customtkinter.CTkFrame(self.search_configuration_frame, corner_radius=5)
        self.search_entry_frame.grid(row=0, column=0, pady=10)

        self.search_label = customtkinter.CTkLabel(self.search_entry_frame, text="Printer name:", 
                                                            compound="left", font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_label.grid(row=0, column=0, padx=5, pady=10)

        self.search_entry = customtkinter.CTkEntry(self.search_entry_frame, width=250, font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_entry.bind("<Return>", lambda event : threading.Thread(target=search_printer(self.search_entry.get())).start())
        self.search_entry.grid(row=0, column=1, padx=5, pady=10)

        self.search_button = customtkinter.CTkButton(self.search_entry_frame, text="Search printer", command=lambda: threading.Thread(target=search_printer(self.search_entry.get())).start(),
                                                text_color=("gray90", "gray90"), hover_color=("gray30", "gray30"), font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_button.grid(row=1, column=1, padx=5, pady=10)

        # Creates search result frame
        self.search_result_frame = customtkinter.CTkFrame(self.search_configuration_frame, corner_radius=5)
        self.search_result_frame.grid(row=1, column=0, pady=10)

        self.search_result_label = customtkinter.CTkLabel(self.search_result_frame, text="Printer name: ", 
                                                            compound="left", font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_result_label.bind("<ButtonRelease-1>", self.open_search_result)
        self.search_result_label.grid(row=0, column=0, padx=5, pady=10)

        self.search_mres_label = customtkinter.CTkLabel(self.search_result_frame, text="Model: ", 
                                                            compound="left", font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_mres_label.grid(row=1, column=0, padx=5, pady=10)

        self.search_lres_label = customtkinter.CTkLabel(self.search_result_frame, text="Location: ", 
                                                            compound="left", font=customtkinter.CTkFont(size=35, weight="bold"))
        self.search_lres_label.grid(row=2, column=0, padx=5, pady=10)

        # Frame for search treeview
        self.search_treeview_frame = customtkinter.CTkFrame(self.search_frame, corner_radius=0, fg_color="transparent")
        self.search_treeview_frame.grid(row=1, column=0, padx=5, pady=2, sticky="ewns")
        
        # This is the search treeview.
        self.tv2 = ttk.Treeview(self.search_treeview_frame, style="mystyle.Treeview")
        self.column_lists = ["Color", "Type", "%"]
        self.tv2['columns'] = self.column_lists
        self.tv2["show"] = "headings"  # Removes empty column
        for column in self.column_lists:
            self.tv2.heading(column, text=column)
            self.tv2.column(column, width=50)
        self.tv2.place(relheight=1, relwidth=0.995)
        self.treescroll2 = customtkinter.CTkScrollbar(self.search_treeview_frame)
        self.treescroll2.configure(command=self.tv2.yview)
        self.tv2.configure(yscrollcommand=self.treescroll2.set)
        self.treescroll2.pack(side="right", fill="y")

        # Configure frame
        self.configure_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.configure_frame.grid_columnconfigure(0, weight=1)

        self.configure_tabview = customtkinter.CTkTabview(self.configure_frame,
                                                            segmented_button_selected_hover_color=("gray30"),
                                                            segmented_button_unselected_hover_color=("gray30"),
                                                            segmented_button_fg_color=("gray10"),
                                                            segmented_button_selected_color=("gray20"),
                                                            segmented_button_unselected_color=("gray10"))
        self.configure_tabview.pack()
        
        self.tab_info = self.configure_tabview.add("Info")

        # Information tabview
        self.tab_info_label = customtkinter.CTkLabel(self.tab_info, text="Printer Monitor", 
                                                            compound="center", font=customtkinter.CTkFont(size=35, weight="bold"))
        self.tab_info_label.grid(row=0, column=0, pady=(5,15), padx=5)

        self.tab_info_frame = customtkinter.CTkFrame(self.tab_info, corner_radius=0, fg_color="transparent")
        self.tab_info_frame.grid(row=1, column=0, pady=5, padx=0)
        
        self.version_num = customtkinter.CTkLabel(self.tab_info_frame, text="Version: ", 
                                                            compound="left", font=customtkinter.CTkFont(size=25, weight="bold"))
        self.version_num.grid(row=1, column=0, pady=5, padx=10)

        self.version_chk = customtkinter.CTkButton(self.tab_info_frame, text="Check for update", command=lambda: threading.Thread(target=check_for_updates).start(),
                                                        font=customtkinter.CTkFont(size=25, weight="bold"))
        self.version_chk.grid(row=1, column=1, pady=5, padx=10)
        
        self.app_creator = customtkinter.CTkLabel(self.tab_info_frame, text="Buster Landstrom",
                                                            compound="left", font=customtkinter.CTkFont(size=25, weight="bold"))

        # Select default frame
        self.select_frame_by_name("home")

        def load_data(): # Loads printer data into treeview
            pid = 0 # Printer element ID
            printerid = 0 # Real printer ID
            subpid = 0 # Sub printer ID
            with open(Paths.ptm_path, 'r') as ptm_list:
                ptm_reader = csv.DictReader(ptm_list, fieldnames = CSVFieldnames.fieldnames)
                for row in ptm_reader:
                    if row[CSVFieldnames.fieldnames[0]] != CSVFieldnames.fieldnames[0]:
                        try:
                            self.tv1.insert(parent="", index="end", values=(row[CSVFieldnames.fieldnames[0]], row[CSVFieldnames.fieldnames[1]], row[CSVFieldnames.fieldnames[2]]), iid=pid, open=True)
                            printerid = pid
                            pid = pid + 1
                            ton1per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner1[1], row[CSVFieldnames.fieldnames[5]])
                            self.tv1.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[4]], row[CSVFieldnames.fieldnames[3]], str(ton1per) + "%"), iid=pid)
                            self.tv1.move(pid, printerid, subpid)
                            pid = pid + 1
                            subpid = subpid + 1
                            if row[CSVFieldnames.fieldnames[3]] != "black":
                                ton2per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner2[1], row[CSVFieldnames.fieldnames[8]])
                                ton3per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner3[1], row[CSVFieldnames.fieldnames[11]])
                                ton4per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner4[1], row[CSVFieldnames.fieldnames[14]])
                                self.tv1.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[7]], row[CSVFieldnames.fieldnames[6]], str(ton2per) + "%"), iid=pid)
                                self.tv1.move(pid, printerid, subpid)
                                pid = pid + 1
                                subpid = subpid + 1
                                self.tv1.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[10]], row[CSVFieldnames.fieldnames[9]], str(ton3per) + "%"), iid=pid)
                                self.tv1.move(pid, printerid, subpid)
                                pid = pid + 1
                                subpid = subpid + 1
                                self.tv1.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[13]], row[CSVFieldnames.fieldnames[12]], str(ton4per) + "%"), iid=pid)
                                self.tv1.move(pid, printerid, subpid)
                                pid = pid + 1
                                if ton1per > 1 and ton2per > 1 and ton3per > 1 and ton4per > 1:
                                    self.tv1.delete(printerid)
                                subpid = 0
                            else:
                                if ton1per > 5:
                                    self.tv1.delete(printerid)
                                subpid = 0
                        except:
                            self.tv1.delete(printerid)
                            print("Did not find printer " + row[CSVFieldnames.fieldnames[0]])

        def search_printer(printername):
            SharePointHandler().download_item(SPClass.context, SPClass.ptm_file_rel_path, Paths.ptm_path)
            pname = printername.upper()
            self.tv2.delete(*self.tv2.get_children())  # *=splat operator
            printername = pname + PrinterVar.domain_name
            with open(Paths.ptm_path, 'r') as ptm_list:
                ptm_reader = csv.DictReader(ptm_list)
                for row in ptm_reader:
                    if row[CSVFieldnames.fieldnames[0]] == printername:
                        self.search_result_label.configure(text="Printer name: " + pname, text_color=("#79ADDC", "#79ADDC"), font=customtkinter.CTkFont(size=35, weight="bold", underline=True), compound="left")
                        self.search_mres_label.configure(text="Model: " + row[CSVFieldnames.fieldnames[1]], compound="left")
                        self.search_lres_label.configure(text="Location: " + row[CSVFieldnames.fieldnames[2]], compound="left")
                        ton1per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner1[1], row[CSVFieldnames.fieldnames[5]])
                        self.tv2.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[3]], row[CSVFieldnames.fieldnames[4]], str(ton1per) + "%"))
                        if row[CSVFieldnames.fieldnames[3]] != "black":
                            ton2per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner2[1], row[CSVFieldnames.fieldnames[8]])
                            ton3per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner3[1], row[CSVFieldnames.fieldnames[11]])
                            ton4per = PrinterManager().get_toner_percentage(PrinterVar.community, row[CSVFieldnames.fieldnames[0]], SNMPVar.toner4[1], row[CSVFieldnames.fieldnames[14]])
                            self.tv2.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[6]], row[CSVFieldnames.fieldnames[7]], str(ton2per) + "%"))
                            self.tv2.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[9]], row[CSVFieldnames.fieldnames[10]], str(ton3per) + "%"))
                            self.tv2.insert(parent="", index="end", value=(row[CSVFieldnames.fieldnames[12]], row[CSVFieldnames.fieldnames[13]], str(ton4per) + "%"))

        def refresh_data():
            SharePointHandler().download_item(SPClass.context, SPClass.ptm_file_rel_path, Paths.ptm_path)
            SharePointHandler().download_item(SPClass.context, SPClass.ptm_shaked_file_rel_path, Paths.ptm_shaked_path)
            self.tv1.delete(*self.tv1.get_children())  # *=splat operator
            self.version_num.configure(text="Version: {}".format(Version.version), compound="left")
            threading.Thread(target=load_data).start()
            if self.initial_load == False:
                self.initial_load = True
                check_for_ptm()

        def countdown():
            self.my_timer = 120
            while self.my_timer > 0 and self.ptm_frame.grid_info() != {}:
                time.sleep(1)
                self.my_timer -= 1
            if self.ptm_frame.grid_info() != {}:
                threading.Thread(target=refresh_data).start()
            if self.winfo_viewable() == True:
                threading.Thread(target=check_for_updates).start()
            check_for_ptm()


        def check_for_ptm():
            if self.ptm_frame.grid_info() != {}:
                threading.Thread(target=countdown).start()
            else:
                self.after(100, check_for_ptm) # run itself again after 1000 ms

        def check_for_updates():
            version_check = AppHandler().version_check(SPClass.context, SPClass.ptm_verison_file_rel_path, Paths.version_path, Version.version, CSVFieldnames.version_fieldnames)
            if version_check[0] == False and version_check[2] != "True":
                self.update_button.grid(row=6, column=0, sticky="ew")
            elif version_check[0] == False and version_check[2] == "True":
                messagebox.showerror('Version out of date!',
                            'Critical Update needed\nOld version: {}\nNew version: {}'.format(Version.version, version_check[1]))
                AppHandler().update_application()
            else:
                self.update_button.grid_forget()
                
        # Checks if the CSVs have been created
        def check_for_refresh():
            if Version.version_checked == True:
                threading.Thread(target=refresh_data).start()
            else:
                self.after(1000, check_for_refresh) # run itself again after 1000 ms

        threading.Thread(target=check_for_refresh).start()
        self.open_toplevel(AdminLogin())
        self.protocol("WM_DELETE_WINDOW", self.quit_app)

    # Frame toggle handler
    def select_frame_by_name(self, name):
        # set button color for selected button
        self.home_button.configure(fg_color=("gray25") if name == "home" else "transparent")
        self.ptm_frame_button.configure(fg_color=("gray25") if name == "ptm" else "transparent")
        self.search_frame_button.configure(fg_color=("gray25") if name == "search" else "transparent")
        self.configure_button.configure(fg_color=("gray25") if name == "configure" else "transparent")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "ptm":
            self.ptm_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.ptm_frame.grid_forget()
        if name == "search":
            self.search_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.search_frame.grid_forget()
        if name == "configure":
            self.configure_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.configure_frame.grid_forget()
    # END

    # Frame navigator events (Can be used to toggle CTkFrame element)
    def home_button_event(self):
        self.select_frame_by_name("home")

    def ptm_button_event(self):
        self.select_frame_by_name("ptm")

    def search_button_event(self):
        self.select_frame_by_name("search")
    
    def configure_button_event(self):
        self.select_frame_by_name("configure")
    # END

    def update_button_event(self):
        AppHandler().update_application(SPClass.context, SPClass.ptm_installer_rel_path, Paths.installer_path)
    
    # Event trigger to change scaling of UI elements
    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)
    # END  
        
    # Opens toplevel window
    def open_toplevel(self, function):
        self.withdraw()
        tl = function
    
    def open_search_result(self, a):
        s = self.search_result_label.cget("text")
        printername = s.replace('Printer name: ', '')
        try:
            Paths.firefox.open_new_tab(printername+PrinterVar.domain_name)
        except: 
            open_new_tab(printername+PrinterVar.domain_name)

    def open_treeview_item(self, a):
        curr_item = self.tv1.focus()
        items = self.tv1.item(curr_item)
        values = items['values']
        printername = values[0]
        try:
            Paths.firefox.open_new_tab(printername)
        except: 
            open_new_tab(printername)

    def quit_app(self):
        sys.exit(1)
# END

# Main application loop
if __name__ == "__main__":

    SPClass = SPClass()
    Version = Version()
    CSVCheck = CSVCheck()
    CSVFieldnames = CSVFieldnames()
    Paths = Paths()
    SNMPVar = SNMPVar()
    PrinterVar = PrinterVar()

    customtkinter.deactivate_automatic_dpi_awareness()
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")
    
    # Main window
    root=PTM()
    root.mainloop()
# END OF FUNCTION