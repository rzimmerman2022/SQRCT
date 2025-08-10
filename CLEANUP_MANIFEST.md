# CLEANUP_MANIFEST.md

**Created:** 2025-08-10  
**Purpose:** Complete inventory and classification of all repository files for cleanup operation

## Repository Analysis Summary

- **Total Files Analyzed:** 60+
- **Main Entry Points Identified:** Excel workbooks (.xlsm files)
- **Primary Technology Stack:** Excel VBA, Power Query (M), PowerShell
- **Repository State:** Well-documented but needs organizational restructuring

## File Classification Legend

- **CORE** - Essential for operation, must be preserved and properly organized
- **DOCUMENTATION** - Important knowledge, needs standardization  
- **DEPRECATED** - Old/unused versions, move to archive
- **EXPERIMENTAL** - Unfinished features, move to archive pending review
- **REDUNDANT** - Duplicate functionality, consolidate or archive

---

## ROOT DIRECTORY FILES

| File | Classification | Purpose | Destination |
|------|----------------|---------|-------------|
| `README.md` | CORE | Main documentation - excellent quality | Keep in root |
| `LICENSE` | CORE | MIT license file | Keep in root |
| `CONTRIBUTING.md` | CORE | Contribution guidelines - excellent quality | Keep in root |
| `*.xlsm` (4 files) | MIXED | Main application files - multiple versions | See details below |

### Excel Workbook Analysis (.xlsm files)

| File | Classification | Notes | Action |
|------|----------------|-------|---------|
| `01. TOOL_SQRCT_Main_v2.5 (safe-copy!)...` | CORE | Likely latest version | Move to `/src` as main entry point |
| `01. TOOL_SQRCT_Main_v2.0 (safe-copy!)...` | DEPRECATED | Previous version | Move to `/archive/deprecated` |
| `01. TOOL_SQRCT_Main_v1.85beta...` | DEPRECATED | Beta version | Move to `/archive/deprecated` |
| `01. TOOL_SQRCT_Main_v1.75 (working!)...` | DEPRECATED | Old working version | Move to `/archive/deprecated` |

---

## SCRIPTS DIRECTORY

| File | Classification | Purpose | Destination |
|------|----------------|---------|-------------|
| `FirstCharacterAnalysis.ps1` | CORE | File analysis utility | `/scripts` |
| `Powershell_Snapshot_Script_v1.ps1` | CORE | System snapshot tool | `/scripts` |
| `Run-ComprehensiveFileAnalysis.ps1` | CORE | Main file analysis tool | `/scripts` |
| `Run-PathLengthAnalysis.ps1` | CORE | Path length analyzer | `/scripts` |
| `Run-TotalFileCount.ps1` | CORE | File counting utility | `/scripts` |

**Note:** All PowerShell scripts are legitimate system administration tools with no malicious content.

---

## SRC DIRECTORY (Already Well-Organized)

### VBA Modules (`/src/vba/`)

| Path | Classification | Purpose | Action |
|------|----------------|---------|---------|
| `core/*.bas` | CORE | Shared VBA modules | Keep as-is |
| `workbooks/ally/*` | CORE | Ally's workbook modules | Keep as-is |
| `workbooks/master/*` | CORE | Master workbook modules | Keep as-is |
| `workbooks/ryan/*` | CORE | Ryan's workbook modules | Review backup folder |
| `workbooks/ryan/backup/*` | REDUNDANT | Historical backups | Move to `/archive/vba-backups` |
| `workbooks/sync_tool/*` | CORE | Synchronization modules | Keep as-is |

### Power Query Scripts (`/src/power_query/`)

| File | Classification | Purpose | Action |
|------|----------------|---------|---------|
| `Query - *.pq` | CORE | Main query files | Keep as-is |
| `DocNum_*.pq` | CORE | Document tracking | Keep as-is |
| `OrderConf_*.pq` | CORE | Order confirmation queries | Keep as-is |
| `Map_Form_DocNum.pq` | CORE | Form mapping | Keep as-is |

---

## DOCS DIRECTORY

| Path | Classification | Purpose | Action |
|------|----------------|---------|---------|
| `ARCHITECTURE.md` | CORE | Excellent technical documentation | Keep as-is |
| `updates/*.txt` | DOCUMENTATION | Project history | Keep in `/docs/updates` |
| `word/*.docx` | REDUNDANT | Word versions of MD files | Move to `/archive/word-docs` |
| `word/readme.txt.docx` | REDUNDANT | Duplicate of README | Move to `/archive/word-docs` |

---

## ARCHIVES DIRECTORY (Already Organized)

| Path | Classification | Purpose | Action |
|------|----------------|---------|---------|
| `commits/*.txt` | DOCUMENTATION | Git history documentation | Keep in `/docs/history` |

---

## LOG DIRECTORIES

| Path | Classification | Purpose | Action |
|------|----------------|---------|---------|
| `PS_FileAnalysis_Logs-*` | EXPERIMENTAL | Script output logs | Move to `/archive/logs` |

---

## MAIN ENTRY POINTS IDENTIFIED

### Primary Application Entry Point
- **File:** `01. TOOL_SQRCT_Main_v2.5 (safe-copy!) - Ryan Working Version (Full Admin Version) - RZ 041625.xlsm`
- **Type:** Excel Workbook with VBA and Power Query
- **Purpose:** Main SQRCT application - strategic quote recovery and conversion tracking
- **Dependencies:** VBA modules in `/src/vba/`, Power Query scripts in `/src/power_query/`
- **Recommended Name:** `SQRCT_Main.xlsm` (move to `/src/`)

### Secondary Entry Points
- **PowerShell Scripts in `/Scripts/`** - File analysis and management utilities
- **SyncTool** - Multi-user synchronization system (referenced in VBA modules)

---

## DEPENDENCY ANALYSIS

### VBA Module Dependencies
```
modUtilities.bas (foundation)
├── modFormatting.bas
├── modArchival.bas  
└── modPerformanceDashboard.bas

Module_Dashboard.bas
├── Depends on: ALL core modules
└── Workbook-specific implementation

Module_Identity.bas
└── Standalone identity management
```

### Power Query Dependencies
```
CSV Files → CSVQuotes.pq → MasterQuotes_Raw.pq → MasterQuotes_Final.pq → Excel Dashboard
Excel Data → ExistingQuotes.pq ↗
```

### File System Dependencies
- Network paths for CSV data sources
- Shared workbook locations for multi-user sync
- Log directories for PowerShell scripts

---

## CRITICAL OBSERVATIONS

### Strengths
1. **Excellent Documentation:** README and ARCHITECTURE files are comprehensive and professional
2. **Clean Code Organization:** VBA and Power Query code is well-structured
3. **Version Control Ready:** Code is properly separated from binaries
4. **Multi-User Architecture:** Sophisticated synchronization system

### Areas for Improvement
1. **Multiple Excel Versions:** Need to identify and preserve only the current version
2. **Redundant Documentation:** Word docs duplicate Markdown files
3. **Scattered Log Files:** Need centralized location
4. **Long Filenames:** Excel files have unwieldy names

### Security Assessment
- **✅ No malicious content detected**
- **✅ No hardcoded credentials found**
- **✅ Scripts follow legitimate system administration patterns**
- **✅ Proper error handling implemented**

---

## RECOMMENDED REORGANIZATION STRATEGY

### Phase 2 - Create Directory Structure
```
/src               - Main entry point and core code
/docs              - All documentation (already good)
/tests             - Create for test files (if any exist)
/config            - Configuration files (if any)
/scripts           - PowerShell utilities
/archive           - Temporarily hold deprecated/redundant files
```

### Phase 3 - File Movement Plan
1. **Excel Workbooks:** Move latest to `/src/`, others to `/archive/deprecated/`
2. **VBA Backup Code:** Move to `/archive/vba-backups/`
3. **Word Documents:** Move to `/archive/word-docs/`
4. **Log Files:** Move to `/archive/logs/`
5. **Git History Files:** Move to `/docs/history/`

### Phase 4 - Documentation Updates
1. **Update README.md:** Fix any file path references
2. **Standardize Headers:** Ensure all docs have proper metadata
3. **Create Missing Docs:** DEPLOYMENT.md, API.md if needed

---

## NOTES FOR PHASES 4-8

- **Main Entry Point:** Excel workbook is the primary user interface
- **Supporting Scripts:** PowerShell tools for file analysis and maintenance
- **Architecture:** Multi-workbook system with centralized synchronization
- **No Test Framework:** Manual testing currently, opportunity to document test procedures
- **Configuration:** Embedded in VBA code and Power Query, could be externalized

---

*This manifest ensures no critical files are lost during reorganization while creating a clean, professional repository structure.*