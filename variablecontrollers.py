from webbrowser import Mozilla
from os import path

# Local file/folder paths
class Paths:
    def __init__(self):
        self._global_path = path.expandvars(r'%LOCALAPPDATA%/PrinterMonitor')  # Local folder where application data is
        self._ptm_path = path.join(self._global_path, 'ptm.csv')
        self._ptm_shaked_path = path.join(self._global_path, 'ptm_shaked.csv')
        self._local_config_path = path.join(self._global_path, 'local_config.csv')
        self._version_path = path.join(self._global_path, 'version.csv')
        self._error_path = path.join(self._global_path, 'error_log.txt')
        self._installer_path = path.join(self._global_path, 'PrinterMonitor.exe')
        self._config_json = path.join(self._global_path, 'config.json')
        self._changes_path = path.join(self._global_path, 'changes.txt')
        self._old_py_path = path.join(self._global_path, 'main.py')
        self._new_py_path = 'main.py'
        self._firefox = Mozilla("C:\\Program Files\\Mozilla Firefox\\firefox.exe")

    @property
    def global_path(self):
        return self._global_path

    @property
    def ptm_path(self):
        return self._ptm_path

    @property
    def ptm_shaked_path(self):
        return self._ptm_shaked_path

    @property
    def local_config_path(self):
        return self._local_config_path

    @property
    def version_path(self):
        return self._version_path

    @property
    def error_path(self):
        return self._error_path

    @property
    def installer_path(self):
        return self._installer_path

    @property
    def config_json(self):
        return self._config_json

    @property
    def changes_path(self):
        return self._changes_path

    @property
    def old_py_path(self):
        return self._old_py_path

    @property
    def new_py_path(self):
        return self._new_py_path

    @property
    def firefox(self):
        return self._firefox


# END

# Variables that keep track of if the CSVs have been created or not
class CSVCheck:
    def __init__(self):
        self._ptm_csv_created = None
        self._ptm_shaked_csv_created = None
        self._version_csv_created = None

    @property
    def ptm_csv_created(self):
        return self._ptm_csv_created

    @ptm_csv_created.setter
    def ptm_csv_created(self, new_value):
        self._ptm_csv_created = new_value

    @property
    def ptm_shaked_csv_created(self):
        return self._ptm_shaked_csv_created

    @ptm_shaked_csv_created.setter
    def ptm_shaked_csv_created(self, new_value):
        self._ptm_shaked_csv_created = new_value

    @property
    def version_csv_created(self):
        return self._version_csv_created

    @version_csv_created.setter
    def version_csv_created(self, new_value):
        self._version_csv_created = new_value


# END

# CSV fieldnames
class CSVFieldnames:
    def __init__(self):
        self._fieldnames = ["Printer", "Model", "Location", "S1Color", "S1Type", "S1Max", "S2Color", "S2Type", "S2Max",
                            "S3Color", "S3Type", "S3Max", "S4Color", "S4Type", "S4Max"]
        self._ammount_shaken_fieldnames = ["Amount of toners shaked"]
        self._version_fieldnames = ["Version", "Critical"]
        self._local_config_fieldnames = [
            "Outlook"]  # Add future "local only" variables here to store them locally into csv file

    @property
    def fieldnames(self):
        return self._fieldnames

    @property
    def ammount_shaken_fieldnames(self):
        return self._ammount_shaken_fieldnames

    @property
    def version_fieldnames(self):
        return self._version_fieldnames

    @property
    def local_config_fieldnames(self):
        return self._local_config_fieldnames


# END

# SharePoint URL paths
class SPClass:
    def __init__(self):
        self._ptm_rel_path = '/sites/unit-support-toolkit/Shared Documents/data/ptm'  # Relative path to the Shared Documents folder that is the root to the Printer Toner Monitor application
        self._ptm_file_rel_path = path.join(self._ptm_rel_path, 'ptm.csv')
        self._ptm_shaked_file_rel_path = path.join(self._ptm_rel_path, 'ptm_shaked.csv')
        self._ptm_verison_file_rel_path = path.join(self._ptm_rel_path, 'version.csv')
        self._ptm_changes_file_path = path.join(self._ptm_rel_path, 'changes.txt')
        self._ptm_py_file_path = path.join(self._ptm_rel_path, 'main.py')
        self._json_config_rel_path = path.join(self._ptm_rel_path, 'config.json')
        self._ptm_installer_rel_path = '/sites/unit-support-toolkit/Shared Documents/Programs/PrinterMonitor/Installer/PrinterMonitor.exe'  # Relative path for the exe installer in the SharePoint documents
        self._team_site_url = 'https://example.sharepoint.com/sites/unit-support-toolkit/'  # This is the url that leads to the teamsite (Change this to lead to the right url for your team)
        self._context = None

    @property
    def ptm_rel_path(self):
        return self._ptm_rel_path

    @ptm_rel_path.setter
    def ptm_rel_path(self, new_value):
        self._ptm_rel_path = new_value

    @property
    def ptm_file_rel_path(self):
        return self._ptm_file_rel_path

    @ptm_file_rel_path.setter
    def ptm_file_rel_path(self, new_value):
        self._ptm_file_rel_path = new_value

    @property
    def ptm_shaked_file_rel_path(self):
        return self._ptm_shaked_file_rel_path

    @ptm_shaked_file_rel_path.setter
    def ptm_shaked_file_rel_path(self, new_value):
        self._ptm_shaked_file_rel_path = new_value

    @property
    def json_config_rel_path(self):
        return self._json_config_rel_path

    @json_config_rel_path.setter
    def json_config_rel_path(self, new_value):
        self._json_config_rel_path = new_value

    @property
    def ptm_verison_file_rel_path(self):
        return self._ptm_verison_file_rel_path

    @ptm_verison_file_rel_path.setter
    def ptm_verison_file_rel_path(self, new_value):
        self._ptm_verison_file_rel_path = new_value

    @property
    def ptm_changes_file_path(self):
        return self._ptm_changes_file_path

    @ptm_changes_file_path.setter
    def ptm_changes_file_path(self, new_value):
        self._ptm_changes_file_path = new_value

    @property
    def ptm_py_file_path(self):
        return self._ptm_py_file_path

    @ptm_py_file_path.setter
    def ptm_py_file_path(self, new_value):
        self._ptm_py_file_path = new_value

    @property
    def ptm_installer_rel_path(self):
        return self._ptm_installer_rel_path

    @ptm_installer_rel_path.setter
    def ptm_installer_rel_path(self, new_value):
        self._ptm_installer_rel_path = new_value

    @property
    def team_site_url(self):
        return self._team_site_url

    @team_site_url.setter
    def team_site_url(self, new_value):
        self._team_site_url = new_value

    @property
    def context(self):
        return self._context

    @context.setter
    def context(self, new_value):
        self._context = new_value


# END

# SNMP printer variables
class PrinterVar:
    def __init__(self):
        self._pnum = 1000  # This is both used as the starting number for specific printers but also can be changed to fit your needs
        self._domain_name = ".example.site.net"  # Site domain name for printers (This can be changed or even removed if it is needed
        self._printer_prefix = "QSE0"  # Prefix for printers (This should be changed since everyone has a different base name for their printers)
        self._community = "public"  # SNMP Community name

    @property
    def pnum(self):
        return self._pnum

    @pnum.setter
    def pnum(self, new_value):
        self._pnum = new_value

    @property
    def domain_name(self):
        return self._domain_name

    @property
    def printer_prefix(self):
        return self._printer_prefix

    @property
    def community(self):
        return self._community


# END

# SNMP OID variables (Object Identifier) THIS IS FOR KYOCERA ONLY RIGHT NOW BUT YOU CAN CHANGE THESE VARIABLES TO WORK WITH ANY MODEL
class SNMPVar:
    def __init__(self):  # Kyocera oid request start is this 1.3.6.1.4.1.1347
        self._hostnameoid = "1.3.6.1.4.1.1347.40.10.1.1.5.1"  # Printer hostname
        self._descoid = "1.3.6.1.2.1.1.1.0"  # Printer description
        self._objid = "1.3.6.1.2.1.1.2.0"  # Printer object ID
        self._uptimeoid = "1.3.6.1.2.1.1.3.0"  # Printer up time
        self._locationoid = "1.3.6.1.2.1.1.6.0"  # Printer location
        self._modeloid = "1.3.6.1.4.1.1347.43.5.1.1.36.1"  # Printer modelname
        # .1.3.6.1.4.1.1347.40.10.1.1.4.1 = Ip Address
        # .1.3.6.1.4.1.1347.40.10.1.1.6.1 = subnetmask
        # .1.3.6.1.4.1.1347.40.10.1.1.7.1 = default gateway
        # .1.3.6.1.4.1.1347.40.10.1.1.8.1 = primary dns server
        # .1.3.6.1.4.1.1347.40.10.1.1.9.1 = secondary dns server

        self._ps_oid = ".1.3.6.1.4.1.1347.43.8.1.1.11.1.1"  # Paper size for mp tray (bw), values are 8=A4, 13=A5-R (These might not match for all models)
        self._pq_oid = "1.3.6.1.4.1.1347.43.10.1.1.11"  # Print quality, values are 1=Auto, 2=Text, 3=Text/Photo, 4=Photo, 5=Graphics. (These might not match for all models)
        self._pd_oid = "1.3.6.1.2.1.2.2.1.10.1"  # Duplex printing, values are 1=On, 2=Off, 3=Booklet. (These might not match for all models)
        self._pc_oid = "1.3.6.1.4.1.1347.42.1.1.1.1.5"  # Color mode, values are 1=Auto Color, 2=Full Color, 3=Grayscale. (These might not match for all models)
        self._po_oid = ".1.3.6.1.2.1.43.15.1.1.7.1.1"  # Orientation, values are 3=Portrait, 4=Landscape. (These might not match for all models), .1 for 1st .2 for 2st etc.
        self._pn_oid = ".1.3.6.1.4.1.1347.43.5.1.1.26.1"  # Number of copies, values are integers between 1 and 999. (These might not match for all models) PS. remember not to print to many
        self._pt_oid = ".1.3.6.1.2.1.43.8.2.1.15.1.2"  # Paper tray (Default paper source), values for color printers are 1=Cassette 1, 2=Cassette 2, 3=Cassette 3, 4=Cassette 4, 5=MP Tray and for bw printers are 1=Cassette 1, 2=MP Tray. (These might not match for all models)
        # .1.3.6.1.4.1.1347.40.10.1.1.10.1 = domainname
        # .1.3.6.1.2.1.43.8.2.1.12.1.2 = mediatype cassette 1 (bw)
        # .1.3.6.1.2.1.43.8.2.1.12.1.1 = mediatype MP Tray (bw)
        # .1.3.6.1.4.1.1347.43.8.1.1.11.1.2 = paper size c1 (bw)

        self._ton1coloroid = "1.3.6.1.2.1.43.12.1.1.4.1.1"  # Slot 1 toner color
        self._ton1maxoid = "1.3.6.1.2.1.43.11.1.1.8.1.1"  # Slot 1 max toner value
        self._ton1curroid = "1.3.6.1.2.1.43.11.1.1.9.1.1"  # Slot 1 current toner ammount
        self._ton1typeoid = "1.3.6.1.2.1.43.11.1.1.6.1.1"  # Slot 1 toner type

        self._ton2coloroid = "1.3.6.1.2.1.43.12.1.1.4.1.2"  # Slot 2 toner color
        self._ton2maxoid = "1.3.6.1.2.1.43.11.1.1.8.1.2"  # Slot 2 max toner value
        self._ton2curroid = "1.3.6.1.2.1.43.11.1.1.9.1.2"  # Slot 2 current toner ammount
        self._ton2typeoid = "1.3.6.1.2.1.43.11.1.1.6.1.2"  # Slot 2 toner type

        self._ton3coloroid = "1.3.6.1.2.1.43.12.1.1.4.1.3"  # Slot 3 toner color
        self._ton3maxoid = "1.3.6.1.2.1.43.11.1.1.8.1.3"  # Slot 3 max toner value
        self._ton3curroid = "1.3.6.1.2.1.43.11.1.1.9.1.3"  # Slot 3 current toner ammount
        self._ton3typeoid = "1.3.6.1.2.1.43.11.1.1.6.1.3"  # Slot 3 toner type

        self._ton4coloroid = "1.3.6.1.2.1.43.12.1.1.4.1.4"  # Slot 4 toner color
        self._ton4maxoid = "1.3.6.1.2.1.43.11.1.1.8.1.4"  # Slot 4 max toner value
        self._ton4curroid = "1.3.6.1.2.1.43.11.1.1.9.1.4"  # Slot 4 current toner ammount
        self._ton4typeoid = "1.3.6.1.2.1.43.11.1.1.6.1.4"  # Slot 4 toner type

        self._toner1 = [self._ton1coloroid, self._ton1curroid, self._ton1maxoid,
                        self._ton1typeoid]  # Toner one container
        self._toner2 = [self._ton2coloroid, self._ton2curroid, self._ton2maxoid,
                        self._ton2typeoid]  # Toner two container
        self._toner3 = [self._ton3coloroid, self._ton3curroid, self._ton3maxoid,
                        self._ton3typeoid]  # Toner three container
        self._toner4 = [self._ton4coloroid, self._ton4curroid, self._ton4maxoid,
                        self._ton4typeoid]  # Toner four container

        self._p_config = [self._ps_oid, self._pq_oid, self._pd_oid, self._pc_oid, self._po_oid, self._pn_oid,
                          self._pt_oid]  # Printer configuration variables

    @property
    def hostnameoid(self):
        return self._hostnameoid

    @property
    def descoid(self):
        return self._descoid

    @property
    def objid(self):
        return self._objid

    @property
    def uptimeoid(self):
        return self._uptimeoid

    @property
    def locationoid(self):
        return self._locationoid

    @property
    def modeloid(self):
        return self._modeloid

    @property
    def ps_oid(self):
        return self._ps_oid

    @property
    def pq_oid(self):
        return self._pq_oid

    @property
    def pd_oid(self):
        return self._pd_oid

    @property
    def pc_oid(self):
        return self._pc_oid

    @property
    def po_oid(self):
        return self._po_oid

    @property
    def pn_oid(self):
        return self._pn_oid

    @property
    def pt_oid(self):
        return self._pt_oid

    @property
    def toner1(self):
        return self._toner1

    @property
    def toner2(self):
        return self._toner2

    @property
    def toner3(self):
        return self._toner3

    @property
    def toner4(self):
        return self._toner4

    @property
    def p_config(self):
        return self._p_config


# END

# Gets version variables
class Version:
    def __init__(self):
        self._version = 2.1  # This is the current version of this Printer Toner Monitor, change this to whatever you want since version control is done automatically when the program has been setup
        self._critical_version = False  # This value notes if the version can be skipped or not (if it has fixes that resolve critical bugs it should be enabled)
        self._test_file = True  # Only updates the version control file if this setting is set to false (false = live version ready for use, true = test version not done yet)
        self._version_checked = False  # Keeps track of if the version has been checked or not during runtime

    @property
    def version(self):
        return self._version

    @property
    def critical_version(self):
        return self._critical_version

    @property
    def test_file(self):
        return self._test_file

    @property
    def version_checked(self):
        return self._version_checked

    @version_checked.setter
    def version_checked(self, new_value):
        self._version_checked = new_value
# END
