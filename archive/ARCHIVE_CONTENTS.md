# ARCHIVE_CONTENTS.md

**Created:** 2025-08-10  
**Purpose:** Documentation of archived files and rationale for archival

## Archive Directory Structure

```
archive/
├── deprecated/          # Older versions of main application
├── vba-backups/        # Historical VBA code backups  
├── word-docs/          # Word document versions of documentation
├── logs/               # PowerShell script execution logs
└── ARCHIVE_CONTENTS.md # This file
```

---

## DEPRECATED (archive/deprecated/)

These are older versions of the main SQRCT Excel workbook. Archived because:
- Superseded by newer version (v2.5)
- Contain outdated functionality  
- Maintained for historical reference and potential rollback needs

| File | Original Version | Archive Reason |
|------|------------------|----------------|
| `01. TOOL_SQRCT_Main_v1.75 (working!)...` | v1.75 | Superseded by v2.5, old working version |
| `01. TOOL_SQRCT_Main_v1.85beta...` | v1.85 beta | Beta version, superseded by stable releases |
| `01. TOOL_SQRCT_Main_v2.0 (safe-copy!)...` | v2.0 | Superseded by v2.5 |

**Recovery Instructions:** If needed, these can be moved back to root directory and renamed appropriately.

---

## VBA-BACKUPS (archive/vba-backups/)

Historical VBA module backups from Ryan's workbook. Archived because:
- Redundant with current modules in `/src/vba/core/`
- Kept for version history and emergency recovery
- Contains experimental code snippets

| File | Purpose | Archive Reason |
|------|---------|----------------|
| `Module_Dashboard (backup).bas` | Dashboard backup | Redundant with current version |
| `Module_Dashboard.bas` | Dashboard module | Duplicate of current code |
| `Module_Identity.bas` | Identity module | Duplicate of current code |
| `Sheet2 (SQRCT Dashboard).cls` | Dashboard sheet class | Duplicate of current code |
| `ThisWorkbook.bas` | Workbook events | Duplicate of current code |
| `mod*.bas` files | Core modules | Duplicates of current core modules |
| `possible good mod_dashboard.vb` | Experimental code | Unverified experimental version |

**Recovery Instructions:** Can be imported into VBA editor if current modules become corrupted.

---

## WORD-DOCS (archive/word-docs/)

Word document versions of documentation. Archived because:
- Redundant with Markdown versions which are easier to maintain
- Word format not suitable for version control
- Markdown versions are more accessible and portable

| File | Markdown Equivalent | Archive Reason |
|------|---------------------|----------------|
| `ARCHITECTURE.md.docx` | `/docs/ARCHITECTURE.md` | Redundant format |
| `CONTRIBUTING.md.docx` | `/CONTRIBUTING.md` | Redundant format |
| `readme.txt.docx` | `/README.md` | Redundant format |

**Recovery Instructions:** Content has been preserved in Markdown format. Word versions only needed if specific formatting is required.

---

## LOGS (archive/logs/)

PowerShell script execution logs. Archived because:
- Historical execution data, not needed for daily operations
- Consume significant disk space
- Useful for debugging and audit trail

| Directory/File | Purpose | Archive Reason |
|----------------|---------|----------------|
| `PS_FileAnalysis_Logs-20250509T004711Z-1-001/` | File analysis log folder | Historical data |
| `PathLengthAnalysis_Transcript_*.log` | Path length analysis logs | Script execution history |
| `Run-ComprehensiveFileAnalysis_*.csv/.log` | Comprehensive analysis outputs | Historical analysis data |

**Recovery Instructions:** These can be referenced for historical analysis patterns or moved to script output directories if needed.

---

## RESTORATION PROCEDURES

### To Restore a Deprecated Excel File:
1. Copy file from `archive/deprecated/` to repository root
2. Rename to appropriate current naming convention  
3. Update any hardcoded paths that may have changed
4. Test thoroughly before using in production

### To Restore VBA Modules:
1. Open VBA Editor in Excel (Alt+F11)
2. Right-click module to replace
3. Select "Remove [ModuleName]"  
4. Go to File → Import File
5. Select file from `archive/vba-backups/`
6. Save workbook

### To Access Historical Documentation:
1. Word documents can be opened directly from `archive/word-docs/`
2. Compare with current Markdown versions if needed
3. Extract any missing content that should be in current docs

### To Reference Historical Logs:
1. Files can be opened directly from `archive/logs/`
2. Use for debugging recurring issues
3. Reference for baseline performance comparisons

---

## DISK SPACE IMPACT

| Category | Approximate Size | Justification |
|----------|------------------|---------------|
| Deprecated Excel files | ~15-20 MB each | Critical for rollback capability |
| VBA backups | ~1-2 MB total | Essential for code recovery |
| Word docs | ~1 MB total | Low cost, historical reference |  
| Log files | ~5-10 MB total | Audit trail and debugging history |

**Total Archive Size:** ~65-85 MB

---

## MAINTENANCE SCHEDULE

- **Quarterly Review:** Assess if archived items can be permanently deleted
- **Annual Cleanup:** Remove files older than 2 years unless specifically needed
- **Version Cleanup:** When new major versions are released, consider removing oldest archived versions

---

## SECURITY CONSIDERATIONS

- Archived Excel files may contain sensitive business data
- Ensure archive directory has same access restrictions as main repository
- VBA modules should be scanned before restoration
- Log files may contain system path information

---

*This archive preserves the repository's history while maintaining a clean working structure. All archived items have been moved based on the analysis documented in CLEANUP_MANIFEST.md.*