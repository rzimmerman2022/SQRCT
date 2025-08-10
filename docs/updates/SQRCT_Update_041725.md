# SQRCT Update - April 17, 2025

**Last Updated:** 2025-04-17  
**Version:** Major Feature Release  
**Type:** Refactoring, Performance, New Features

## Summary
feat: Refactor data loading, fix layout, add Active/Archive views

This commit implements major updates to the SQRCT Dashboard VBA code, addressing performance, layout issues, error handling, and adding new Active/Archive view functionality.

**`Module_Dashboard` Refactoring & Fixes:**

* **Refactored Data Loading (Performance & Robustness):**
    * Replaced the older, formula-based `PopulateMasterQuotesData`, `PopulateWorkflowLocation`, and `RestoreUserEditsToDashboard` subroutines with a new, faster array-based approach.
    * Added `BuildDashboardDataArray` function: This is the new core function that loads all source data (MasterQuotes_Final, DocNum_LatestLocation, UserEdits) into memory arrays/dictionaries, merges them according to defined logic (including Phase resolution), and returns a single complete data array ready to be written to the dashboard sheet.
    * Added `GetMQF_HeaderMap` and `MQFIdx` helper functions: Enables `BuildDashboardDataArray` to look up required columns in the `MasterQuotes_Final` source by header *name* (e.g., "Document Number", "AutoStage", "DataSource", "Historic Stage") instead of relying on fixed column index numbers (like 1, 10, 11). This makes the code significantly more robust if columns are reordered in the Power Query source.
    * Added/Verified `GetTableOrRangeData` and `BuildDictionaryFromArray` general helper functions for reading source tables/ranges.
    * Modified `LoadUserEditsToDictionary` function: Changed it to return a dictionary mapping `CleanedDocNum -> Array(Phase, Contact, Comments)` for efficient use by `BuildDashboardDataArray`.
    * Replaced the `RefreshDashboard` subroutine: Implemented a new version that orchestrates the refactored flow: Call `BuildDashboardDataArray` -> Write the returned array directly to the sheet as values -> Sort -> Apply Formatting -> Call `modArchival.RefreshAllViews` -> etc.
    * Deleted obsolete subroutines: `PopulateMasterQuotesData`, `PopulateWorkflowLocation`, `RestoreUserEditsToDashboard`.

* **Layout Fixes:**
    * Standardized column constants for J (`DB_COL_WORKFLOW_LOCATION`) and K (`DB_COL_MISSING_QUOTE`) and ensured correct usage throughout.
    * Corrected the header array definition in `InitializeDashboardLayout` to match the J=Workflow, K=Missing Quote order.
    * Fixed Column F width calculation in `RefreshDashboard`'s formatting step (Step 7) to use a fixed width (`10.5`) after initial AutoFit, preventing the date column from becoming too wide.
    * Implemented safe Row 3 height calculation in `RefreshDashboard`'s formatting step (Step 7) using `.AutoFit` followed by a minimum height check (`If .RowHeight < 15 Then .RowHeight = 15`), preventing header truncation.

* **Error Fixes & Logic Updates:**
    * Added `BuildRowIndexDict` helper function: Creates a `CleanedDocNum -> Sheet Row Number` dictionary specifically for the `Worksheet_Change` event.
    * Modified `Worksheet_Change` (on SQRCT Dashboard sheet): Now uses `BuildRowIndexDict` to find the target row on the `UserEdits` sheet, resolving the Error 424 ("Object required" / Type Mismatch) caused by the refactored `LoadUserEditsToDictionary`.
    * Added `ResolvePhase` helper function: Implements the logic to determine the correct "Engagement Phase" to display, correctly handling the "Legacy Process" placeholder and prioritizing user edits over source data (`Historic Stage` for Legacy source, `AutoStage` for non-Legacy).
    * Modified `BuildDashboardDataArray`: Calls `ResolvePhase` to populate the Phase column (L / index 12) in the output array. Requires reading `DataSource` and `Historic Stage` from the source.
    * Modified `SaveUserEditsFromDashboard`: Added a check to prevent saving the literal text "Legacy Process" to the `UserEdits` sheet, ensuring the `ResolvePhase` logic works correctly on subsequent refreshes.
    * Added `ToggleWorkbookStructure` helper function (Public): Manages workbook structure protection, called by `modArchival` when adding sheets. Ensured calls were added to `RefreshDashboard`'s main execution path and cleanup.
    * Ensured necessary Constants and Subs/Functions used by `modArchival` or sheet modules (e.g., `DASHBOARD_SHEET_NAME`, `DB_COL_*`, `PW_WORKBOOK`, `ModernButton`, `ApplyColorFormatting`, `ApplyWorkflowLocationFormatting`, `LogUserEditsOperation`, `ToggleWorkbookStructure`, `BuildRowIndexDict`, `RowIndexDictAdd`, `ResolvePhase`) are declared `Public`.
    * Corrected module name references (`modDashboard` -> `Module_Dashboard`) in calls from `modArchival`.

* **Other Updates:**
    * Added an explicit Number Formatting step in the refactored `RefreshDashboard` after sorting to ensure correct display of currency, dates, and integers.
    * Ensured `DEBUG_WorkflowLocation` constant is set to `False`.
    * Added the call to `modArchival.RefreshAllViews` near the end of `RefreshDashboard`.

**New Module `modArchival`:**

* Created new module `modArchival` to encapsulate Active/Archive view logic.
* Added code to `modArchival` to:
    * Define constants for Active phases (`ACTIVE_PHASES`), sheet names (`SH_ACTIVE`, `SH_ARCHIVE`), and titles.
    * Implement `RefreshAllViews` as the main entry point called by `RefreshDashboard`.
    * Implement `RefreshActiveView` and `RefreshArchiveView` to manage sheet creation/clearing.
    * Implement `GetOrCreateSheet` helper (includes title formatting and structure protection toggle).
    * Implement `CopyFilteredRows` helper:
        * Reads data from main dashboard sheet into an array using the robust `.Resize` method to prevent errors caused by Excel omitting blank trailing columns (Fixes Error 9 - "Subscript out of range").
        * Calculates the actual column count (`numCols`) from the array read via `.Resize`.
        * Filters rows based on `IsPhaseActive`/`IsPhaseArchived` helper functions (checking against `ACTIVE_PHASES`).
        * Writes the filtered data array to the target sheet (`SH_ACTIVE` or `SH_ARCHIVE`).
        * Adds a row count display to the sheet.
    * Implement `ApplyViewFormatting` helper: Copies column widths from main dashboard, applies number formats, applies conditional formatting (calling helpers in `Module_Dashboard`), sets freeze panes (Rows 1-3), protects sheets as read-only (allowing selection only). Corrected unlock range to start at Row 4.
    * Implement `AddNavigationButtons` helper: Deletes old buttons, calls `Module_Dashboard.ModernButton` to add "All Items", "Active", "Archive" buttons to Row 2 (anchored G2, I2, K2), renames buttons for consistency. Includes sheet unprotect/reprotect logic.
    * Implement `btnViewActive`, `btnViewArchive`, `btnViewAll` public subs as button click handlers (these call `RefreshAndActivate`).
    * Implement `Log` helper shim to call `Module_Dashboard.LogUserEditsOperation`.
