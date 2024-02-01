Attribute VB_Name = "TextESRD"
'3.2 File format
'All metadata files must be stored in .csv format.
'-   The values/fields separator shall be a dollar sign (ASCII code=&#36), shown as: $
'-   The metadata must end with one line containing the text: EOF.


'8.1.1   E-mail header - "Subject:"
'To be able to automatic recognize the file type, the "Subject" of the e-mail must consist of exact one of the following subjects:
'File Category   Integration Type    "Subject:" in e-mail header
'Integration files   TIR-Supplies    TIR-Supplies
'    TIR-Tools   TIR-Tools
'    TIR-Enterprise  TIR-Enterprise
'    TIR-Zones   TIR-Zones
'    TIR-Circuit-Breakers    TIR-Circuit-Breakers
'    TIR-Access-Points   TIR-Access-Points
'    Illustrations Illustrations
'    IPC-Spares  IPC-Spares
'    Wiring Equipment List
'
'd
'
'    Wiring-Equipment List
'    Wiring Wire List
'    Wiring -WireList
'    Wiring Hook-up List,
'Plug and receptacle
'    Wiring-Plug&ReceptacleList
'    Wiring Hook-up List,
'Terminal
'    Wiring -TerminalList
'    Wiring Hook-up List,
'Splice
'    Wiring -SpliceList
'    Wiring Hook-up List, Earth Point    Wiring-EarthPointList



'3.3 File name
'The naming convention for the Meta data file itself shall consist of the following 4-6 different parts depending on the file category of engineering source. See matrix below:
'
'    File Category   Metadata    Integration Type (Applicable for Integration file only) Date    Time    Extension
'Illustration files  Integration _Metadata   _DB _Illustration   _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_Metadata_DB_Illustration_20110524_0854.csv
'
'Engineering source data Author  _Metadata   _YYYYMMDD   _HHMM   .csv
'
'Example: Author_Metadata_20110521_1540.csv
'
'Engineering source data ConvertedDM _Metadata   _YYYYMMDD   _HHMM   .csv
'
'Example: ConvertedDM_Metadata_20130404_1540.csv
'

'
'4.2 File name
'The naming convention for the TIR-data file itself shall consist of the following 6 different parts. See matrix below:
'    File Category   Integration Type (Applicable for Integration file only) Date    Time    Extension
'TIR-data files  Integration _DB _TIR_Supplies
'_TIR_Tools
'_TIR_Enterprise
'_TIR_Circuit_Breakers
'_TIR_Zones
'_TIR_Access_Points  _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_DB_TIR_Tools_20110523_1153.csv


'5.2 File name
'The naming convention for the IPC integration files itself shall consist of the following 6 different parts. See matrix below:
'    File Category   Integration Type (Applicable for Integration file only) Date    Time    Extension
'IPC integration files   Integration _DB _ IPC-Spares
'    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_DB_IPC -Spares_20110523_1153.csv

'6.3 File name
'The naming convention for the wiring integration files itself shall consist of the following 6 different parts. See the matrix below.
'    File Category   Integration Type (Applicable for Integration file only) DMC Date    Time    Extension
'Wiring integration files - Equipment List   Integration     _XML    _Wiring_ EL
'    _DMC    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring_EL_MRJ-A-91-00-00-00A-056A-A_20120523_1153.csv
'
'Wiring
'integration files - Wire List   Integration _XML    _Wiring_WL
'    _DMC    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring_WL_MRJ-A-91-00-AB-00A-057A-A_20120523_1154.csv
'
'Wiring integration files - Hook-up List, Plug and receptacle    Integration _ XML   _Wiring_HL-PR   _DMC
'    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring_HL-PR_MRJ-A-91-00-00-00A-057D-A_20120523_1155.csv
'
'Wiring integration files - Hook-up List, Terminal   Integration _ XML   _Wiring_HL-T
'    _DMC    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring_HL-T_MRJ-A-91-00-00-00A-057E-A_20120523_1156.csv
'
'Wiring integration files - Hook-up List, Splice Integration _XML    _Wiring_HL-S
'    _DMC    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring_HL-S_MRJ-A-91-00-00-00A-057F-A_20120523_1157.csv
'
'Wiring integration files - Hook-up List, Earth Point    Integration _XML    _Wiring_HL-EP
'    _DMC    _YYYYMMDD   _HHMM   .csv
'
'Example: Integration_XML_Wiring _HL-EP_MRJ-A-91-00-00-00A-057G-A _20120523_1158.csv





'3.4.1   Meta data for engineering source data files
'Name of Data Filed / Requirement    Description / Allowable value   Remarks
'File Category (Mandatory)   File Category of Engineering Source.
'Only applicable with the value Author
'File Name (Mandatory)   Document/ID number of Engineering Source:
'E.g. NA212234   The file name for the content file must not include the issue number of the file. The issue number must be handled as metadata.
'
'Note: For amendment specifications, the file issue for its corresponding base specification is included in the file name.
'File Issue (Mandatory)  The version number for the engineering source data file:
'E.g. A  File issue numbering is either natural numbers or alpha-based.
'
'File issues based on natural numbers adheres to the following sequence: 1, 2, 3, …, 9, 10, 11, …, 99, 100, 101, … Note: Dates that adhere to the ISO 8601 YYYYMMDD format are allowed since they are a subset of the natural numbers
'
'File issues that are alpha-based adheres to the following sequence: NC, A, B, C, …, Y, , AA, AB, …, AY,  BA, BB, … The letter I, O, X, Z must not be used in this series (e.g. AI, AO, AX, AZ etc.). When the version handling is coming to NB, NC is not allowed to used, since this is the first issue.
'File Title (Mandatory)  The title of the Engineering Source:
'MRJ-200 System Configuration Index-Electrical Power System (ATA24)
'File Format (Mandatory) The format of the engineering source data file:
'E.g.pdf , cgm, tif, XML, xls, jpg, mpeg, doc
'Engineering Source Category (Mandatory) The category of engineering source:
'E.g. DWG, Specification etc.    For a complete list of allowed values, refer to column 1 "Engineering Source Category" in the table in chapter 8 "File Category and Format".
'Responsible Department (Mandatory)  The responsible organization (MITAC or Department/Group at MITAC) for engineering source data
'Aircraft Model (Mandatory)  Applicable Aircraft Model for the engineering source data file:
'E.g. MRJ-200, MRJ-300,
'Both
'VridgeR Reference (Optional)    -   Field not used
'Zone (Optional) Applicable zone number for engineering Source:
'E.g. 120
'Access Point (Optional) Applicable access point for engineering source:
'E.g. 517
'Part Number (Optional)  Applicable part number for engineering source
'E.g. 5771331-1101
'Export Control (Conditional)    Applicable Export Control  Classification Number for engineering source data file:
'E.g. ECCN 9E991
'E.g. ECCN 7E004.b2, ECCN 7E994  Mandatory when the engineering source data require the re-export control  regulated by the U.S. EAR.
'
'When Applicable Export Control Number has two or more values, each value must be separated by a comma "," and a space.
'Comments (Optional) Additional information
'Active (Mandatory)  Active used to make a document not valid. Only applicable with the value Yes or No  Criteria for Yes or No:
'Yes = For use (normally)
'No = Document must not be used anymore
'Change Number (Mandatory)   Written notice number:
'E.g.WNN -XXXXXX
'Change order number:
'E.g.CON -XXXXXX
'N/A:
'E.g. N/A    Written notice number: No effect on the Basic design concept of MRJ according change management process
'Change order number: Effect on the Basic design concept of MRJ according change management process
'N/A: No change order connected to the document
'Original Engineering Source ID [Extracted From] (Optional)  ID for the original engineering source. This meta data will be required, if the contents of engineering source was extracted from other engineering document, and ID was changed from original one.
'
