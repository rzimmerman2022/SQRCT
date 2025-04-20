SQRCT System Architecture (Updated Apr 20, 2025)
1. Introduction
This document describes the architecture of the SQRCT (Strategic Quote Recovery & Conversion Tracker) system in its current Excel-based implementation. It details the major components, data flow, data models, error handling, and security considerations based on the Power Query M code and VBA modules developed for the User/Master workbooks and the separate SyncTool workbook.
The primary purpose of this system is to track sales quotes, automate follow-up stage calculation, allow user input for engagement tracking via a standardized interface, provide distinct filtered views for Active and Archived quotes, and synchronize user edits into a master dataset, all within the Microsoft Excel environment.
2. High-Level Overview
The SQRCT system operates across multiple interconnected Excel workbooks: individual user workbooks (e.g., for Ryan "RZ", Ally "AF") and a central "Automated Master" workbook. A separate "SyncTool" workbook orchestrates the merging of user edit data between these files.
Core Components:
•	Data Sources: Network folders (daily CSV quote exports), local Excel tables (historical data).
•	Power Query Engine (within User/Master Workbooks): Ingests (CSVQuotes, ExistingQuotes), merges (MasterQuotes_Raw), transforms, calculates status (AutoStage, AutoNote), and prepares the final quote dataset (MasterQuotes_Final).
•	Excel Data Target (within User/Master Workbooks): The output of the MasterQuotes_Final query, loaded into the workbook for VBA access.
•	VBA Core Logic (Module_Dashboard): Orchestrates the main dashboard refresh (RefreshDashboard), builds the core data array (BuildDashboardDataArray), merges Power Query data with user edits, handles initial sheet setup (SetupDashboard), UserEdits backup/restore, and contains various helper functions (e.g., CleanDocumentNumber, ModernButton).
•	VBA View Generation (modArchival): Creates and manages the "SQRCT Active" and "SQRCT Archive" sheets (RefreshActiveView/RefreshArchiveView), including filtering data (CopyFilteredRows using IsPhaseArchived), applying consistent formatting (ApplyViewFormatting), creating UI elements (AddNavigationButtons), and managing view-specific record counts via Properties (ActiveRecords, ArchiveRecords). Includes a helper (FormatControlRow) for standard Row 2 styling.
•	VBA Utilities (modUtilities): Contains shared helper functions like Engagement Phase prefix lookup (GetPhaseFromPrefix), applying Data Validation (ApplyPhaseValidationToListColumn), and displaying/styling dynamic counts (UpdateAllViewCounts).
•	VBA User Edit Capture (Worksheet_Change on Dashboard Sheet): Automatically captures changes made by users in specific columns (L-N: Phase, Last Contact, Comments) and writes them to the hidden "UserEdits" sheet within the same workbook, tagging the edit with user identity and timestamp.
•	VBA Phase Validation (Workbook_SheetChange in ThisWorkbook): Intercepts changes in Phase columns (L on Dashboard, B on UserEdits), uses GetPhaseFromPrefix for auto-complete/validation against PHASE_LIST, and triggers prompts for "Other" phases.
•	Hidden "UserEdits" Sheet (within User/Master Workbooks): Persistent log of user modifications (Columns A-F: DocNum, Phase, LastContact, Comments, ChangeSource, Timestamp).
•	SyncTool Workbook (VBA Application): Separate Excel file containing VBA modules for manually synchronizing "UserEdits" data between User/Master workbooks. Reads all sources, resolves conflicts (timestamp priority), writes merged data back only to the Master "UserEdits" sheet, and logs its actions.
•	UI Elements: Standardized Row 2 controls including refresh buttons, view navigation buttons, dynamic count display (J2:L2), and timestamp (N2).
Diagram (Conceptual Flow):
Code snippet
flowchart TD
    subgraph UserMasterWb [User/Master Workbook (RZ/AF/Master)]
        direction LR
        subgraph PQ_Engine [Power Query Engine]
            direction TB
            CSVSource[("CSV Files")] --> CSVQuotes
            ExcelSource[("ExistingQuotes Table")] --> ExistingQuotes
            CSVQuotes --> MasterQuotes_Raw
            ExistingQuotes --> MasterQuotes_Raw
            MasterQuotes_Raw --> MasterQuotes_Final[MasterQuotes_Final (Calculates AutoStage/Note)]
        end

        subgraph VBA_Engine [VBA Engine]
             direction TB
             MasterQuotes_Final -- Reads Data --> ModDash(Module_Dashboard)
             UserEditsSheet[(UserEdits Sheet)] -- Reads/Writes --> ModDash
             ModDash -- Builds Array --> RefreshDash[RefreshDashboard]

             RefreshDash -- Calls --> ModArch(modArchival)
             RefreshDash -- Calls --> ModUtil(modUtilities)
             RefreshDash -- Writes Data & Formats --> DashboardSheet{SQRCT Dashboard}

             ModArch -- Creates/Updates --> ActiveSheet{SQRCT Active}
             ModArch -- Creates/Updates --> ArchiveSheet{SQRCT Archive}
             ModArch -- Reads Data --> DashboardSheet
             ModArch -- Calls --> ModUtil

             DashboardSheet -- User Edit L-N --> WsCode{Worksheet_Change}
             WsCode -- Writes --> UserEditsSheet

             ThisWb{ThisWorkbook} -- SheetChange --> ModUtil
             ModUtil -- Phase Lookup --> ListsSheet[(Lists Sheet / PHASE_LIST)]

        end

        PQ_Engine --> VBA_Engine

    end

    subgraph SyncToolWb [SyncTool Workbook (Manual)]
        direction TB
        SyncUI{SyncTool UI} --> SyncCode[Sync VBA Logic]
        SyncCode -- Reads --> RZ_UserEdits[(RZ UserEdits)]
        SyncCode -- Reads --> AF_UserEdits[(AF UserEdits)]
        SyncCode -- Reads --> Master_UserEdits[(Master UserEdits)]
        SyncCode -- Merges/Resolves --> MergedData
        MergedData -- Writes --> Master_UserEdits
        SyncCode -- Triggers Refresh --> MasterWorkbookRef[Master Workbook Refresh]
    end

    UserMasterWb -- Contains --> RZ_UserEdits
    UserMasterWb -- Contains --> AF_UserEdits
    UserMasterWb -- Is Target For --> MergedData

    style UserEditsSheet fill:#f9f,stroke:#333,stroke-width:2px
    style RZ_UserEdits fill:#f9f,stroke:#333,stroke-width:2px
    style AF_UserEdits fill:#f9f,stroke:#333,stroke-width:2px
    style Master_UserEdits fill:#f9f,stroke:#333,stroke-width:2px
3. Component Breakdown
•	Power Query Queries: (Descriptions remain largely the same as your previous version, assuming CLIENT QUOTES is still separate/adjunct) 
o	CLIENT QUOTES: Identifies quote document files/metadata. Potentially for ad-hoc use.
o	CSVQuotes: Ingests daily CSVs.
o	ExistingQuotes: Loads historical data from local table.
o	MasterQuotes_Raw: Combines CSV and Existing data.
o	MasterQuotes_Final: Core processing - filters, calculates age/status (AutoStage, AutoNote), deduplicates. Loads output to Excel for VBA.
•	VBA Modules (User/Master Workbooks): 
o	Module_Dashboard: Orchestrates RefreshDashboard. Calls BuildDashboardDataArray. Manages UserEdits backup/restore (CreateUserEditsBackup, RestoreUserEditsFromBackup). Contains initial setup (SetupDashboard), button helper (ModernButton), data processing helpers (CleanDocumentNumber, ResolvePhase, etc.), sheet protection (ProtectUserColumns), and CF application helpers (ApplyColorFormatting, etc.). Interacts heavily with other modules.
o	modArchival: Handles generation of SQRCT Active / SQRCT Archive sheets (RefreshActiveView, RefreshArchiveView). Contains filtering logic (CopyFilteredRows, IsPhaseArchived). Manages view formatting (ApplyViewFormatting) and consistent Row 2 UI (AddNavigationButtons, FormatControlRow). Stores/provides view counts (ActiveRecords, ArchiveRecords Properties).
o	modUtilities: Contains shared helper functions: GetPhaseFromPrefix (for auto-complete), ApplyPhaseValidationToListColumn (for dropdown setup), UpdateAllViewCounts (for displaying counts in J2:L2).
o	Module_Identity: Defines workbook owner ("RZ", "AF", "MASTER").
o	ThisWorkbook Code: Contains Workbook_SheetChange event handler to manage Engagement Phase input validation, auto-complete, and "Other" phase prompting by calling modUtilities.GetPhaseFromPrefix.
o	Worksheet Code (SQRCT Dashboard Sheet): Contains Worksheet_Change event handler to capture user edits in columns L, M, N and save them to the hidden UserEdits sheet.
•	VBA Modules (SyncTool Workbook): (Based on previous README - verify if accurate) 
o	Module_SyncTool_Manager: Orchestrates synchronization workflow.
o	Module_File_Processor: Handles reading/writing external User/Master workbooks' UserEdits sheets.
o	Module_Conflict_Handler: Implements conflict resolution logic (timestamps, comments).
o	Supporting Modules: For Logging, UI, Constants, Utilities within the SyncTool.
4. Data Model
•	MasterQuotes_Final (Power Query Output / Excel Table): 
o	Structure: Processed quote data (Doc Num, Dates, Customer Info, Salesperson, AutoStage, AutoNote, DataSource, etc. - Columns A:K approx).
o	Storage: Loaded into Excel Table/Connection. Read by VBA.
•	UserEdits (Hidden Sheet in RZ/AF/Master): 
o	Structure: User overrides. Columns: A: DocNumber (Key), B: Engagement Phase, C: Last Contact Date, D: User Comments, E: ChangeSource ("RZ", "AF", "MASTER"), F: Timestamp.
o	Storage: Hidden Excel sheet. Acts as local change log. Master sheet is target for SyncTool.
•	SyncTool Log Sheets: SyncLog, ErrorLog, DocChangeHistory within SyncTool workbook.
5. Data Flow
1.	Data Ingestion (Power Query): CSVs + Local Table -> MasterQuotes_Raw -> MasterQuotes_Final -> Loaded to Excel.
2.	Dashboard Refresh (Module_Dashboard.RefreshDashboard): 
o	(If SaveAndRestore Mode): Read L-N from Dashboard -> Update UserEdits sheet (SaveUserEditsFromDashboard).
o	Build data array (BuildDashboardDataArray) by merging MasterQuotes_Final output with current UserEdits data (using ResolvePhase logic).
o	Write merged array A:N to Dashboard sheet.
o	Sort Dashboard.
o	Apply column widths, number formats, data validation (ApplyPhaseValidationToListColumn).
o	Apply Conditional Formatting (ApplyColorFormatting, etc.).
o	Protect Dashboard, unlocking L:N (ProtectUserColumns); Apply Freeze Panes (FreezeDashboard).
o	Call modArchival.RefreshAllViews to regenerate Active/Archive sheets.
o	Call modArchival.FormatControlRow(ws) to set base grey A2:N2 format on Dashboard.
o	Apply A2 blue override styling to Dashboard A2.
o	Call modArchival.AddNavigationButtons(ws) to add buttons and timestamp to Dashboard N2.
o	Call modUtilities.UpdateAllViewCounts(ws) (using unprotect/reprotect wrapper) to display counts in Dashboard J2:L2.
o	Create/Update Text-Only Sheet.
o	Display completion message.
3.	User Editing (Worksheet_Change on Dashboard): 
o	User changes cell in L, M, or N.
o	Event triggers -> Get DocNum (A) -> Find/Create row in hidden UserEdits -> Write L, M, N values to UserEdits B, C, D -> Write User ID (E) & Timestamp (F).
4.	Phase Input (Workbook_SheetChange in ThisWorkbook): 
o	User changes cell in L (Dashboard) or B (UserEdits).
o	Event triggers -> Call modUtilities.GetPhaseFromPrefix.
o	If unique match -> Auto-complete/correct case.
o	If "Other..." match -> Show prompt, select Comments column.
o	If no/ambiguous match -> Show error, undo change.
5.	Synchronization (SyncTool - Manual): 
o	User selects RZ, AF, Master file paths in SyncTool UI -> Clicks "Sync".
o	SyncTool VBA reads UserEdits from all 3 files.
o	Resolves conflicts (timestamp priority, comment merge).
o	Writes resolved data only to UserEdits sheet in Master workbook.
o	Triggers refresh of Master workbook.
o	Logs actions in SyncTool log sheets.
6. Error Handling Strategy
•	Power Query: Relies mainly on default behavior; some try...otherwise. Limited logging.
•	VBA (User/Master): Uses On Error GoTo [Label] in main routines (RefreshDashboard, ApplyViewFormatting, etc.) for controlled cleanup. Uses On Error Resume Next more sparingly in specific formatting/object access blocks where failure is non-critical or handled immediately. Errors logged via DebugLog to Immediate Window and potentially UserEditsLog sheet.
•	VBA (SyncTool): Structured error handling with centralized logging to dedicated sheets within the SyncTool.
7. Security Considerations
•	Password: Sheet/Workbook protection uses a blank password defined in Module_Dashboard.PW_WORKBOOK constant.
•	Network Path Access: Power Query and SyncTool require user/tool access to specified file paths. Parameterization or configuration sheets recommended over hardcoding.
•	SyncTool File Access: Requires Read/Write access to User/Master workbooks.
•	Macro Security: Relies on users enabling macros. Standard Trust Center settings apply.
•	Data Exposure: Access control relies on filesystem/SharePoint permissions.
8. Deployment & Execution
•	System Operation: Runs entirely within Microsoft Excel (.xlsm files).
•	User Interaction: Users work within their individual .xlsm files, interacting with the "SQRCT Dashboard".
•	Data Refresh: Power Query refreshes likely triggered manually or via VBA refresh buttons. VBA refreshes triggered by buttons.
•	Synchronization: Manual process using the separate SyncTool workbook. Requires user intervention to select files and initiate.
•	Code Updates: Require manual export/import between the Git repository and the operational Excel files' VBA Editors.


