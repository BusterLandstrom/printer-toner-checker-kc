# Printer Toner Checker KC

A modular and configurable SNMP-based tool to monitor toner level and toner model name built in Python, currently optimized for Kyocera printers (But is easily modified to work with any SNMP enabled printer device).  
It supports both enterprise environments with SharePoint integration and smaller office/home setups with multiple printers.

---

## Features

- **SNMP Toner Level Retrieval** – Fetches toner and status data via SNMP.
- **Enterprise Authentication** – Uses secure SharePoint authentication to control user access.
- **Auto-Updating** – Place a new installer in your SharePoint `Shared Documents` folder and the app will automatically download, replace, and run the latest version.
- **Experimental Features** – Hidden utilities and tools for users to discover.
- **Modular Design** – Easily extend functionality by adding new OIDs and visual components.
- **Dynamic Device Scanning** – Automatically scans devices based on configurable numeric identifiers, which can easily be swapped out for a static IP list if preferred.

---

## Setup & Configuration

This application is primarily built with `customtkinter` (but works with `tkinter`), and can be packaged as a single `.exe` for easy deployment using `NSIS`.

To configure it for your environment:

### Lines to Update for Custom Use:

- **Device Range / Identifier Logic**  
  By default, the app scans devices using a numeric range and identifier pattern. To use a static list of IPs instead, you can easily just update the logic and scan it instead.

- **Authentication Settings**  
  Replace company-specific authentication headers or tokens where indicated.

- **Update URL (SharePoint)**  
  If you are going to be using auto-updates (This is only intended if the app will be continually developed in-house), set the path to your internal SharePoint `.exe` installer file:
  ```python
  "https://yourcompany.sharepoint.com/sites/YourSite/Shared%20Documents/ptminstaller.exe"
  ```

- **GUI Components**  
  To add more SNMP values to the interface:
  1. Add the corresponding OID in the SNMP class.
  2. Add a new line in the GUI setup where you want it displayed.

  That’s it — the app handles the rest.

---

## Use Cases

- **For Companies:**  
  Deploy this internally to monitor multiple network printers, enforce access via SharePoint, and allow seamless updates without distributing new versions manually.

- **For Power Users/Home Office:**  
  Great for home setups with multiple printers. Run the app locally, update identifiers or IPs as needed, and manage your printer fleet with ease.