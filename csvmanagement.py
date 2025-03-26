from printermanager import PrinterManager
from shutil import move
from os import path, mkdir, remove
from tempfile import NamedTemporaryFile
from variablecontrollers import *

# CSV Generation manager
class GenerateCSV():
    def __init__(self):

        def GenerateTonerCSV():
            if CSVCheck.ptm_csv_created == False: # Creates the csv data file for ONLY the printer names
                with open (Paths.ptm_path, 'w') as ptm_list:
                    ptm_writer = csv.DictWriter(ptm_list, fieldnames = CSVFieldnames.fieldnames)
                    ptm_writer.writeheader()
                    while PrinterVar.pnum < 8000:
                        printername = PrinterVar.printer_prefix + str(PrinterVar.pnum) + PrinterVar.domain_name
                        try:
                            modelname = str(PrinterManager().get_printer_var(PrinterVar.community, printername, SNMPVar.modeloid))
                            location = str(PrinterManager().get_printer_var(PrinterVar.community, printername, SNMPVar.locationoid))
                            if modelname != "":
                                slot1 = PrinterManager().get_slot_info(PrinterVar.community, modelname, printername, SNMPVar.toner1[0], SNMPVar.toner1[2], SNMPVar.toner1[3])
                                if slot1[0] != "black":
                                    slot2 = PrinterManager().get_slot_info(PrinterVar.community, modelname, printername, SNMPVar.toner2[0], SNMPVar.toner2[2], SNMPVar.toner2[3])
                                    slot3 = PrinterManager().get_slot_info(PrinterVar.community, modelname, printername, SNMPVar.toner3[0], SNMPVar.toner3[2], SNMPVar.toner3[3])
                                    slot4 = PrinterManager().get_slot_info(PrinterVar.community, modelname, printername, SNMPVar.toner4[0], SNMPVar.toner4[2], SNMPVar.toner4[3])
                                    row = {CSVFieldnames.fieldnames[0]: printername, CSVFieldnames.fieldnames[1]: modelname, CSVFieldnames.fieldnames[2]: location, CSVFieldnames.fieldnames[3]: slot1[0], CSVFieldnames.fieldnames[4]: slot1[1], CSVFieldnames.fieldnames[5]: slot1[2], CSVFieldnames.fieldnames[6]: slot2[0], CSVFieldnames.fieldnames[7]: slot2[1], CSVFieldnames.fieldnames[8]: slot2[2], CSVFieldnames.fieldnames[9]: slot3[0], CSVFieldnames.fieldnames[10]: slot3[1], CSVFieldnames.fieldnames[11]: slot3[2], CSVFieldnames.fieldnames[12]: slot4[0], CSVFieldnames.fieldnames[13]: slot4[1], CSVFieldnames.fieldnames[14]: slot4[2]}
                                else:
                                    row = {CSVFieldnames.fieldnames[0]: printername, CSVFieldnames.fieldnames[1]: modelname, CSVFieldnames.fieldnames[2]: location, CSVFieldnames.fieldnames[3]: slot1[0], CSVFieldnames.fieldnames[4]: slot1[1], CSVFieldnames.fieldnames[5]: slot1[2]}
                                ptm_writer.writerow(row)
                                print(row)
                        except:
                            print("Printer " + printername + " not found")
                        PrinterVar.pnum = PrinterVar.pnum + 1
                SharePointHandler().upload_item(SPClass.context, Paths.ptm_path, SPClass.ptm_rel_path)
            else:
                SharePointHandler().download_item(SPClass.context, SPClass.ptm_file_rel_path, Paths.ptm_path)

        def GenerateVersionCSV():
            if CSVCheck.version_csv_created == False: # Creates the csv file for version control
                with open (Paths.version_path, 'w') as version_list:
                    version_writer = csv.DictWriter(version_list, fieldnames = CSVFieldnames.version_fieldnames)
                    version_writer.writeheader()
                    row = {CSVFieldnames.version_fieldnames[0]: Version.version, CSVFieldnames.version_fieldnames[1]: Version.critical_version}
                    version_writer.writerow(row)
                SharePointHandler().upload_item(SPClass.context, Paths.version_path, SPClass.ptm_rel_path)
            else:
                SharePointHandler().download_item(SPClass.context, SPClass.ptm_verison_file_rel_path, Paths.version_path)
                update_csv = False
                tempfile = NamedTemporaryFile(mode="w", delete=False)
                with open (Paths.version_path, 'r') as version_list, tempfile:
                    version_reader = csv.DictReader(version_list, fieldnames = CSVFieldnames.version_fieldnames)
                    version_writer = csv.DictWriter(tempfile, fieldnames = CSVFieldnames.version_fieldnames)
                    version_writer.writeheader()
                    for row in version_reader:
                        if row[CSVFieldnames.version_fieldnames[0]] != CSVFieldnames.version_fieldnames[0]:
                            oversion = float(row[CSVFieldnames.version_fieldnames[0]])
                            if oversion < Version.version and Version.test_file == False:
                                row = {CSVFieldnames.version_fieldnames[0]: Version.version, CSVFieldnames.version_fieldnames[1]: Version.critical_version}
                                version_writer.writerow(row)
                                update_csv = True
                            if Version.test_file != False:
                                print("generated")
                                #SharePointHandler().download_item(SPClass.context, SPClass.ptm_py_file_path, Paths.old_py_path)
                                #AppHandler().generate_change_log(SPClass.context, SPClass.ptm_rel_path, Paths.old_py_path, Paths.new_py_path, Paths.changes_path) Disabled since the function is only experimental
                if update_csv == True:
                    move(tempfile.name, Paths.version_path)
                    SharePointHandler().upload_item(SPClass.context, Paths.version_path, SPClass.ptm_rel_path)
                else:
                    try:
                        remove(tempfile.name)
                    except OSError as error:
                        with open (Paths.error_path, 'w') as error_log:
                            error_log.write(error)

        def GenerateShakedCSV():
            if CSVCheck.ptm_shaked_csv_created == False: # Creates the csv file (Shaked database)
                with open (Paths.ptm_shaked_path, 'w') as shaked_list:
                    shaked_writer = csv.DictWriter(shaked_list, fieldnames = CSVFieldnames.ammount_shaken_fieldnames)
                    shaked_writer.writeheader()
                SharePointHandler().upload_item(SPClass.context, Paths.ptm_shaked_path, SPClass.ptm_rel_path)
            else:
                SharePointHandler().download_item(SPClass.context, SPClass.ptm_shaked_file_rel_path, Paths.ptm_shaked_path)

        if path.exists(Paths.global_path) == False:
            try:
                mkdir(Paths.global_path)
            except OSError as error:
                print(error)

        threading.Thread(target=GenerateVersionCSV()).start()
        threading.Thread(target=GenerateTonerCSV()).start()
        threading.Thread(target=GenerateShakedCSV()).start()
        #SharePointHandler().download_item(SPClass.context, SPClass.json_config_rel_path, Paths.config_json) This downloads a config json which just automatically loads important variables into memory but it is experimental and has not been configured yet
# END


class CSVChecker: # Kinda redundant but I really want to keep track of the csv creations so no errors are thrown

    def ptmcsvcheck(self):
        ptm_file = SharePointHandler().fop_exist(SPClass.context, SPClass.ptm_file_rel_path, Paths.error_path)
        print(ptm_file)  # This is used for debugging purposes
        if ptm_file:
            CSVCheck.ptm_csv_created = True
        else:
            CSVCheck.ptm_csv_created = False

    def ptmscsvcheck(self):
        ptms_file = SharePointHandler().fop_exist(SPClass.context, SPClass.ptm_shaked_file_rel_path, Paths.error_path)
        if ptms_file:
            CSVCheck.ptm_shaked_csv_created = True
        else:
            CSVCheck.ptm_shaked_csv_created = False

    def versioncsvchecker(self):
        version_file = SharePointHandler().fop_exist(SPClass.context, SPClass.ptm_verison_file_rel_path, Paths.error_path)
        if version_file:
            CSVCheck.version_csv_created = True
        else:
            CSVCheck.version_csv_created = False

