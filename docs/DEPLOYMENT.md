# SQRCT Deployment Guide

**Last Updated:** 2025-08-10  
**Version:** 4.0.0  
**Document Type:** Deployment Instructions

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Production Environment Setup](#production-environment-setup)
- [Installation Steps](#installation-steps)
- [Configuration](#configuration)
- [Testing Deployment](#testing-deployment)
- [Post-Deployment Tasks](#post-deployment-tasks)
- [Rollback Procedures](#rollback-procedures)
- [Monitoring](#monitoring)
- [Troubleshooting](#troubleshooting)

---

## Overview

This document provides step-by-step instructions for deploying the SQRCT (Strategic Quote Recovery & Conversion Tracker) system to a production environment. SQRCT is an Excel-based application with VBA macros and Power Query components that requires careful deployment to ensure all dependencies are properly configured.

### Deployment Architecture

```
Production Environment
├── Network Share (CSV Data Sources)
├── User Workstations (Excel 2016+)
├── Shared Network Location (Master Workbook)
└── SyncTool Workstation (Admin)
```

---

## Prerequisites

### System Requirements

**For All Users:**
- Windows 10 or higher
- Microsoft Excel 2016 or later (Excel 365 recommended)
- Power Query support (included in Excel 2016+)
- VBA macro support enabled
- Network connectivity to shared resources

**For Administrators:**
- Administrative access to shared network locations
- PowerShell execution policy configured for scripts
- Git client (for code version control)

### Network Infrastructure

- **CSV Data Source Path:** Network location with read access for all users
- **Shared Workbook Location:** Network path with read/write access for team members
- **SyncTool Location:** Accessible by designated admin users

### Security Requirements

- File system permissions properly configured
- Macro security settings adjusted for trusted locations
- Network access to required paths
- User accounts with appropriate permissions

---

## Production Environment Setup

### 1. Network Share Configuration

Create the following directory structure on your network share:

```
\\your-server\SQRCT-Production\
├── Data\                    # CSV data sources
├── Workbooks\              # User workbooks
│   ├── Master\             # Master workbook location
│   ├── Ryan\               # Ryan's workbook
│   └── Ally\               # Ally's workbook
├── SyncTool\               # Synchronization tool
└── Scripts\                # PowerShell utilities
```

### 2. Permission Setup

Configure the following permissions:

| Path | User/Group | Permissions |
|------|------------|-------------|
| `\\server\SQRCT-Production\Data\` | All Users | Read Only |
| `\\server\SQRCT-Production\Workbooks\Master\` | All Users | Read; Admin: Read/Write |
| `\\server\SQRCT-Production\Workbooks\Ryan\` | Ryan, Admin | Read/Write |
| `\\server\SQRCT-Production\Workbooks\Ally\` | Ally, Admin | Read/Write |
| `\\server\SQRCT-Production\SyncTool\` | Admin Only | Read/Write |
| `\\server\SQRCT-Production\Scripts\` | All Users | Read/Execute |

### 3. Trusted Locations Configuration

Configure Excel Trusted Locations on each user's workstation:

1. Open Excel → File → Options → Trust Center → Trust Center Settings
2. Navigate to Trusted Locations
3. Add the following paths:
   - `\\your-server\SQRCT-Production\`
   - Check "Subfolders of this location are also trusted"

---

## Installation Steps

### Step 1: Deploy Excel Workbooks

1. **Copy Main Workbook to Network**
   ```bash
   copy src\SQRCT_Main.xlsm "\\server\SQRCT-Production\Workbooks\Master\"
   ```

2. **Create User-Specific Copies**
   ```bash
   copy "\\server\SQRCT-Production\Workbooks\Master\SQRCT_Main.xlsm" "\\server\SQRCT-Production\Workbooks\Ryan\SQRCT_Ryan.xlsm"
   copy "\\server\SQRCT-Production\Workbooks\Master\SQRCT_Main.xlsm" "\\server\SQRCT-Production\Workbooks\Ally\SQRCT_Ally.xlsm"
   ```

### Step 2: Configure VBA Modules

For each workbook (Master, Ryan, Ally):

1. **Open Excel Workbook**
   - Open the workbook in Excel
   - Press `Alt+F11` to open VBA Editor

2. **Import Core Modules** (Required for all workbooks)
   - File → Import File
   - Import each file from `src\vba\core\`:
     - `modUtilities.bas`
     - `modFormatting.bas`
     - `modArchival.bas`
     - `modPerformanceDashboard.bas`

3. **Import Workbook-Specific Modules**
   - **For Master workbook:** Import from `src\vba\workbooks\master\`
   - **For Ryan workbook:** Import from `src\vba\workbooks\ryan\`
   - **For Ally workbook:** Import from `src\vba\workbooks\ally\`

4. **Save and Close**
   - Save each workbook after importing modules
   - Close VBA Editor

### Step 3: Configure Power Query

For each workbook:

1. **Open Power Query Editor**
   - Data → Get Data → Launch Power Query Editor

2. **Import Query Files**
   - Advanced Editor → Copy and paste content from each `.pq` file in `src\power_query\`
   - Create queries in this order:
     1. `Query - CSVQuotes.pq`
     2. `Query - ExistingQuotes.pq`
     3. `Query - MasterQuotes_Raw.pq`
     4. `Query - MasterQuotes_Final.pq`
     5. All other supporting queries

3. **Update Data Source Paths**
   - Edit `Query - CSVQuotes.pq` to point to production CSV path
   - Update any hardcoded paths to use production network locations

4. **Test and Apply**
   - Refresh all queries to test connectivity
   - Close & Apply to save changes

### Step 4: Deploy SyncTool

1. **Create SyncTool Workbook**
   ```bash
   copy src\SQRCT_Main.xlsm "\\server\SQRCT-Production\SyncTool\SQRCT_SyncTool.xlsm"
   ```

2. **Import SyncTool Modules**
   - Open `SQRCT_SyncTool.xlsm`
   - Press `Alt+F11`
   - Import all files from `src\vba\workbooks\sync_tool\`

3. **Configure SyncTool Settings**
   - Open the SyncTool workbook
   - Configure file paths on the dashboard to point to production workbooks
   - Test synchronization functionality

### Step 5: Deploy PowerShell Scripts

1. **Copy Scripts**
   ```bash
   xcopy Scripts\*.ps1 "\\server\SQRCT-Production\Scripts\" /Y
   ```

2. **Update Script Paths**
   - Edit each PowerShell script to use production paths
   - Update log directory paths as needed

---

## Configuration

### 1. Environment-Specific Settings

Create a configuration sheet in each workbook with these settings:

| Setting | Production Value |
|---------|------------------|
| CSV_DATA_PATH | `\\server\SQRCT-Production\Data\` |
| MASTER_WORKBOOK_PATH | `\\server\SQRCT-Production\Workbooks\Master\` |
| LOG_DIRECTORY | `\\server\SQRCT-Production\Logs\` |
| SYNC_FREQUENCY | Daily at 6 PM |

### 2. User Identity Configuration

Verify that each workbook has the correct identity set:

- **Ryan's workbook:** `Module_Identity.bas` → `WORKBOOK_IDENTITY = "RZ"`
- **Ally's workbook:** `Module_Identity.bas` → `WORKBOOK_IDENTITY = "AF"`
- **Master workbook:** `Module_Identity.bas` → `WORKBOOK_IDENTITY = "MASTER"`

### 3. Phase Validation Setup

Ensure the `PHASE_LIST` named range is properly configured in each workbook:

1. Go to Formulas → Name Manager
2. Verify `PHASE_LIST` exists and points to correct range
3. Update phase validation dropdown if needed

---

## Testing Deployment

### 1. Individual Workbook Testing

For each workbook:

- [ ] Open workbook without errors
- [ ] Click "Standard Refresh" button
- [ ] Verify data loads correctly
- [ ] Test user edit functionality (columns K-N)
- [ ] Verify Active/Archive views generate correctly
- [ ] Test navigation buttons
- [ ] Confirm formatting applied correctly

### 2. Multi-User Testing

- [ ] Have Ryan and Ally make test edits simultaneously
- [ ] Run SyncTool to merge edits
- [ ] Verify conflict resolution works correctly
- [ ] Check Master workbook updated properly

### 3. End-to-End Testing

- [ ] Complete daily workflow from CSV refresh to sync
- [ ] Test error scenarios (missing CSV, locked files)
- [ ] Verify performance with production data volume
- [ ] Test PowerShell scripts execution

### 4. Performance Testing

- [ ] Measure dashboard refresh times
- [ ] Test with maximum expected data volume
- [ ] Verify memory usage is acceptable
- [ ] Check network bandwidth requirements

---

## Post-Deployment Tasks

### 1. User Training

- [ ] Conduct user training sessions
- [ ] Provide quick reference guides
- [ ] Document any environment-specific procedures
- [ ] Set up support contact information

### 2. Documentation

- [ ] Update network paths in documentation
- [ ] Create user-specific guides
- [ ] Document any customizations made
- [ ] Update architecture diagrams with production paths

### 3. Monitoring Setup

- [ ] Configure file access logging
- [ ] Set up automated CSV data checks
- [ ] Create synchronization monitoring
- [ ] Establish backup procedures

### 4. Security Review

- [ ] Audit file permissions
- [ ] Review macro security settings
- [ ] Verify network access controls
- [ ] Document security configuration

---

## Rollback Procedures

### Emergency Rollback

If critical issues occur:

1. **Immediately notify all users to stop using system**
2. **Restore from backup branch:**
   ```bash
   git checkout pre-cleanup-backup-2025-08-10
   ```
3. **Redeploy previous version using same steps above**
4. **Restore data from most recent known good state**

### Planned Rollback

For planned rollbacks:

1. **Schedule maintenance window**
2. **Backup current user data**
3. **Deploy previous version**
4. **Migrate any critical data changes**
5. **Test thoroughly before releasing**

### Data Recovery

In case of data corruption:

1. **Stop all user access immediately**
2. **Restore Master workbook from backup**
3. **Extract UserEdits from individual workbooks**
4. **Run manual synchronization**
5. **Verify data integrity before resuming operations**

---

## Monitoring

### Daily Checks

- [ ] Verify CSV data refresh completed
- [ ] Check synchronization logs
- [ ] Monitor user edit activity
- [ ] Verify all workbooks accessible

### Weekly Checks

- [ ] Review error logs
- [ ] Analyze performance metrics
- [ ] Check disk space usage
- [ ] Verify backup completion

### Monthly Reviews

- [ ] User feedback collection
- [ ] Performance trend analysis
- [ ] Security audit
- [ ] Documentation updates

---

## Troubleshooting

### Common Issues

**Issue:** "Macro Security Warning"
- **Cause:** Workbook not in trusted location
- **Solution:** Add network path to Excel Trusted Locations

**Issue:** "Power Query Data Source Error"
- **Cause:** Incorrect CSV path or permissions
- **Solution:** Verify network path and user permissions

**Issue:** "SyncTool Cannot Access Workbooks"
- **Cause:** File locks or permission issues
- **Solution:** Ensure workbooks are closed, check permissions

**Issue:** "VBA Compile Error"
- **Cause:** Missing module references
- **Solution:** Verify all modules imported correctly

### Performance Issues

**Slow Refresh Times:**
1. Check network connectivity
2. Verify CSV file size reasonable
3. Optimize Power Query steps
4. Consider data archival

**High Memory Usage:**
1. Close unused workbooks
2. Reduce dashboard data size
3. Optimize VBA array operations
4. Consider splitting large datasets

### Error Logging

Enable detailed logging by setting debug flags in VBA:

```vba
Private Const DEBUG_MODE As Boolean = True
Private Const LOG_LEVEL As String = "VERBOSE"
```

Check logs in:
- `\\server\SQRCT-Production\Logs\`
- Individual workbook error logs
- Windows Event Viewer (for system-level issues)

---

## Support Contacts

- **Primary Admin:** [Contact Information]
- **Technical Support:** [Contact Information]
- **Network Administrator:** [Contact Information]
- **Business Owner:** [Contact Information]

---

*This deployment guide ensures a successful production rollout of SQRCT while maintaining data integrity and system security. Follow all steps carefully and test thoroughly at each stage.*