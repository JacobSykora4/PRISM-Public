PRISM - Prospect Research & Intelligence System for Major Gifts
====================================================
Author: Jacob V. Sykora  
Organization: Archdiocese of Omaha  
File: PRISM.xlam  
Version: 2.7 (October 2025)

----------------------------------------------------
DESCRIPTION
----------------------------------------------------
PRISM is an Excel-based toolkit for automating 
donor research and data gathering. It integrates 
public data sources (like parcel search sites, 
FEC campaign contributions, and Whitepages) into 
a streamlined, repeatable process.

The system is packaged as a .xlam add-in and runs 
entirely within Excel using VBA and Power Query.

----------------------------------------------------
KEY FEATURES
----------------------------------------------------
- Whitepages Lookup Automation
  > Generates search URLs and parses results via Power Query

- FEC Contributions Search
  > Queries the FEC API by name and address, auto-filters
    duplicates and ambiguous matches

- Parcel Data Querying
  > Supports:
    • Beacon (Schneider Geospatial)
    • Sarpy County GIS
    • Dickinson County, IA
    • gWorks counties in Nebraska

- Address Normalization
  > Built-in Power Query functions:
    • Normalize
    • Unabreviate
    • normalizeColumns

- Geographic Filtering
  > Supports NE and IA counties with mapping for ZIP and parcel
    search URLs

- Link Generation
  > Auto-generates URLs for LinkedIn, Whitepages, and county 
    assessor records

----------------------------------------------------
INSTALLATION
----------------------------------------------------
1. Save PRISM.xlam to a trusted folder
2. Follow instructions in the "Start-up and Usage guide.pdf"

----------------------------------------------------
USAGE NOTES
----------------------------------------------------
- PRISM is designed for internal use by the Archdiocese 
  Development Office.
- It avoids external dependencies like RibbonX or Selenium.
- Sheets in the add-in are hidden or removed. All logic is 
  modular and centralized in VBA modules.
- Designed for extensibility: county mappings, name filters, 
  and data extraction logic are modular.

----------------------------------------------------
SUPPORT & CONTACT
----------------------------------------------------
Author: Jacob V. Sykora  
Arch O Email: jvsykora@archomaha.org
Personal Email: sykojv@gmail.com

For support, updates, or maintenance documentation, 
contact the Office of Stewardship and Development.

----------------------------------------------------
DISCLAIMER
----------------------------------------------------
This tool is provided "as-is" with no warranties. 
Unauthorized redistribution or modification is prohibited.
Users are responsible for complying with data privacy 
and acceptable use policies.

----------------------------------------------------
LICENSE
----------------------------------------------------

PRISM (Prospect Research and Intelligence System for Major Gifts) is the intellectual property of Jacob Sykora and the Archdiocese of Omaha.

Permission to use PRISM is granted exclusively to:

- The Archdiocese of Omaha
- Other Catholic dioceses or affiliated institutions with prior written approval from Jacob Sykora or the Archdiocese of Omaha.

Redistribution, modification, or use outside the approved scope is prohibited without express permission.

All rights reserved.
