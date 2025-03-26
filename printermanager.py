from pysnmp.entity.rfc3413.oneliner import cmdgen
from pysnmp.hlapi import *


# SNMP Requests for printer preferences
class PrinterManager():
    '''
        These functions also dont have explanations because
        the function names and variables should be easy to understand.
    '''

    def get_printer_var(self, community, printername, oid):
        g = getCmd(
            SnmpEngine(), CommunityData(community), UdpTransportTarget(
                (printername, 161)), ContextData(), ObjectType(
                ObjectIdentity(oid)))

        errorIndication, errorStatus, errorIndex, varBinds = next(g)
        for varBind in varBinds:
            for x in varBind:
                if type(x) != ObjectIdentity:
                    var = x
        return var

    def get_slot_info(self, community, model, printername, coloroid, maxoid, typeoid):

        if model != "FS-4200DN" and model != "FS-3920DN":
            toner_color = str(self.get_printer_var(community, printername, coloroid))
            toner_type = str(self.get_printer_var(community, printername, typeoid))
            toner_max = int(self.get_printer_var(community, printername, maxoid))
        else:
            toner_color = "black"
            toner_type = str(self.get_printer_var(community, printername, typeoid))
            toner_max = int(self.get_printer_var(community, printername, maxoid))

        slot = [toner_color, toner_type, toner_max]
        return slot

    def get_toner_percentage(self, community, printername, curroid, toner_max):

        toner_curr = int(self.get_printer_var(community, printername, curroid))

        tp = 100 * toner_curr / int(toner_max)

        tp = round(tp)

        return tp

    def configure_printer_var(self, community, printername, varoid, val):

        # Initializing command generator
        cmdGen = cmdgen.CommandGenerator()

        # Build config request message
        errorIndication, errorStatus, errorIndex, varBinds = cmdGen.setCmd(
            cmdgen.CommunityData(community),
            cmdgen.UdpTransportTarget((printername, 161)),
            ((varoid, val),)
        )

        # Check for result and raise exception for error management
        if errorIndication:
            raise Exception(errorIndication)
        else:
            return f"Successfully configured printer variable with oid {varoid}"

    def get_printer_config(self, community, printername, pvariables, **kwargs):

        ps_oid = pvariables[0]
        pq_oid = pvariables[1]
        pd_oid = pvariables[2]
        pc_oid = pvariables[3]
        po_oid = pvariables[4]
        pn_oid = pvariables[5]
        pt_oid = pvariables[6]

        pconfig = list()

        pconfig.append(str(PrinterManager().get_printer_var(community, printername, ps_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, pq_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, pd_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, pc_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, po_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, pn_oid)))
        pconfig.append(str(PrinterManager().get_printer_var(community, printername, pt_oid)))
        print(pconfig)

        return pconfig
# END