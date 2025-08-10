# SQRCT Update - April 18, 2025

**Last Updated:** 2025-04-18  
**Version:** v4.0.0  
**Type:** Features, Bug Fixes, UI Standardization

## Summary
Feat/Fix: Standardize UI Layout, Fix Button Creation & Phase Filter Bugs (v4.0.0)

Description: Addresses UI inconsistencies, button errors, and filter bugs. Implements standardized Row 2 layout and resolves compile errors. Establishes v4.0.0 baseline.

Detailed Changes:

1.  Feature: Standardized Row 2 UI Layout & Control (modArchival::AddNavigationButtons, Module_Dashboard::SetupDashboard)
    * Implemented new consistent Row 2 layout: Control Panel (A2), Blank (B2), Std Refresh (C2, auto-width), Preserve Edits (D2, auto-width), Blank (E2), All Items (F2, slim), Active (G2, slim), Archive (H2, slim), Blank (I2), Placeholder Counts (J2, K2, L2), Blank (M2), Timestamp (N2).
    * Centralized Row 2 UI element creation logic within modArchival::AddNavigationButtons using a btnDefs array definition for easier maintenance.
    * Updated AddNavigationButtons to calculate button positions based on target cell anchors (C2, D2, F2, G2, H2) and explicitly center buttons.
    * Implemented auto-fit width logic for C2/D2 buttons and fixed slimmer width (65px) for F2/G2/H2 buttons within AddNavigationButtons.
    * Added placeholder text ("All: TBC", "Active: TBC", "Archive: TBC") with Font Size 9 to J2, K2, L2 via AddNavigationButtons.
    * Added left-aligned timestamp ("Refreshed: ...") with Font Size 9 to N2 via AddNavigationButtons.
    * Removed all cell merges from Row 2 via SetupDashboard.
    * Added code to clear Row 2 shapes, content, and merges within AddNavigationButtons before creating new elements.
    * (Optional) Added standardized column width settings to the end of SetupDashboard to support new Row 2 layout.

2.  Feature: Enhanced & Robust Button Creation (Module_Dashboard::ModernButton)
    * Refactored ModernButton from a Sub to a Public Function returning the created Shape object.
    * Modified ModernButton signature to accept targetCell As Range and ByVal buttonWidth As Double. Added ByVal to string parameters.
    * Added explicit code within ModernButton to force Fill/Font visibility and ZOrder.
    * Updated AddNavigationButtons to correctly call ModernButton Function, pass arguments, capture shape, and modify it directly.

3.  Fix: UI Consistency Across Sheets (Module_Dashboard::RefreshDashboard, modArchival::ApplyViewFormatting)
    * Removed obsolete Module_Dashboard::SetupDashboardUI_EndRefresh subroutine and calls.
    * Added new Step 9 block to Module_Dashboard::RefreshDashboard to explicitly set standard row heights and call modArchival.AddNavigationButtons for main dashboard.
    * Confirmed modArchival::ApplyViewFormatting also sets standard row heights.

4.  Fix: Blank Phase Filtering (modArchival::IsPhaseArchived, ::IsPhaseActive)
    * Confirmed IsPhaseArchived logic correctly returns False for blank phases (treating them as Active). (Note: Verification for other phases pending).
    * Refined IsPhaseActive (pending final logic decision).

5.  Fix: Date Formatting on Dashboard (Module_Dashboard::RefreshDashboard)
    * Added explicit .NumberFormat = "mm/dd/yyyy" for date columns (E, F, M) in Step 7.

6.  Fix: Multiple Compile Errors
    * Resolved various compile errors by correcting signatures, calls, removing obsolete code, and fixing array syntax (incl. btnDefs definition).

7.  Refactoring/Code Quality:
    * Introduced btnDefs array in AddNavigationButtons.
    * Improved shape handling in AddNavigationButtons.
    * Added enhanced logging to CopyFilteredRows filter loop.

Pending:
* Final verification that non-blank, non-active phases (e.g., Second F/U) are correctly filtered to Archive.
* Implementation of logic to calculate and pass real Total/Active/Archive counts to AddNavigationButtons.
* Implementation of Phase Auto-Complete/Validation system.
* Potential future refactoring (modConstants, etc.).