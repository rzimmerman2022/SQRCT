# SQRCT - Strategic Quote Recovery & Conversion Tracker

```
███████╗ ██████╗ ██████╗  ██████╗████████╗
██╔════╝██╔═══██╗██╔══██╗██╔════╝╚══██╔══╝
███████╗██║   ██║██████╔╝██║        ██║   
╚════██║██║▄▄ ██║██╔══██╗██║        ██║   
███████║╚██████╔╝██║  ██║╚██████╗   ██║   
╚══════╝ ╚══▀▀═╝ ╚═╝  ╚═╝ ╚═════╝   ╚═╝   
Strategic Quote Recovery & Conversion Tracker
```

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Excel-green.svg)](https://docs.microsoft.com/en-us/office/vba/)
[![Power Query](https://img.shields.io/badge/Power%20Query-M%20Language-blue.svg)](https://docs.microsoft.com/en-us/power-query/)

## 📋 Table of Contents

- [Overview](#-overview)
- [System Architecture](#-system-architecture)
- [Repository Structure](#-repository-structure)
- [Installation](#-installation)
- [Usage](#-usage)
- [Development](#-development)
- [Contributing](#-contributing)
- [Changelog](#-changelog)
- [License](#-license)

## 🎯 Overview

SQRCT is an enterprise-grade Excel-based system designed to track and manage the follow-up process for sales quotes. The system improves quote conversion rates by providing visibility into outstanding quotes, automating follow-up stage calculation, and facilitating consistent engagement by the sales team.

### Key Features

- **Automated Quote Tracking**: Ingests daily CSV exports and maintains historical data
- **Smart Stage Calculation**: Automatically determines follow-up stages based on quote age and interaction history
- **Multi-User Support**: Individual workbooks for team members (Ryan "RZ", Ally "AF") with centralized synchronization
- **Conflict Resolution**: Intelligent merge system resolves editing conflicts using timestamps
- **Active/Archive Views**: Separate views for active quotes requiring follow-up and archived/completed quotes
- **Data Validation**: Built-in phase validation with auto-complete functionality

### Technology Stack

- **Microsoft Excel**: Primary platform (.xlsm files with macro support)
- **Power Query (M Language)**: Data ingestion and transformation pipeline
- **VBA (Visual Basic for Applications)**: User interface, automation, and business logic
- **Git**: Version control for code modules (manual export/import process)

## 🏗️ System Architecture

The system consists of three main components working in concert:

### 1. Data Processing Pipeline (Power Query)

```
CSV Files ──┐
            ├──> MasterQuotes_Raw ──> MasterQuotes_Final ──> Excel Table
Excel Table ┘
```

**Key Queries:**
- `CSVQuotes.pq`: Ingests daily CSV quote exports from network location
- `ExistingQuotes.pq`: Loads historical quote data from Excel table
- `MasterQuotes_Raw.pq`: Combines CSV and historical data
- `MasterQuotes_Final.pq`: Core processing - deduplication, stage calculation, data cleaning

### 2. User Interface Layer (VBA)

```
Power Query Data ──> Dashboard ──> User Edits ──> Hidden Sheet
                         │              │
                         └──────────────┘
                          Event Capture
```

**Core Modules:**
- `Module_Dashboard.bas`: Main dashboard refresh and data management
- `modArchival.bas`: Generates Active/Archive filtered views
- `modUtilities.bas`: Shared utilities and validation functions
- `modFormatting.bas`: UI formatting and styling functions

### 3. Synchronization System (SyncTool)

```
Ryan's Edits ──┐
               ├──> Conflict Resolution ──> Master Workbook
Ally's Edits ──┘         (Timestamp)
```

**SyncTool Modules:**
- `Module_SyncTool_Manager.bas`: Orchestrates synchronization workflow
- `Module_File_Processor.bas`: Reads/writes UserEdits across workbooks
- `Module_Conflict_Handler.bas`: Implements timestamp-based conflict resolution

## 📁 Repository Structure

```
SQRCT/
├── 📄 README.md                     # This file - main documentation
├── 📄 LICENSE                       # MIT License
├── 📄 CONTRIBUTING.md               # Contribution guidelines
├── 📄 .gitignore                    # Git ignore rules
├── 📄 .copilotignore                # Copilot ignore rules
├── 📂 src/                          # Source code directory
│   ├── 📂 vba/                      # VBA modules and classes
│   │   ├── 📂 core/                 # Shared modules used across workbooks
│   │   │   ├── modArchival.bas      # Archive/Active view management
│   │   │   ├── modFormatting.bas    # UI formatting functions
│   │   │   ├── modUtilities.bas     # General utility functions
│   │   │   └── modPerformanceDashboard.bas  # Performance metrics
│   │   └── 📂 workbooks/            # Workbook-specific VBA code
│   │       ├── 📂 ally/             # Ally's workbook modules
│   │       │   ├── Module_Dashboard.bas
│   │       │   ├── Module_Identity.bas  # Sets WORKBOOK_IDENTITY = "AF"
│   │       │   └── Sheet12 (SQRCT Dashboard).cls
│   │       ├── 📂 master/           # Master workbook modules
│   │       │   ├── Module1.bas      # Dashboard logic
│   │       │   └── Sheet2 (SQRCT Dashboard).cls
│   │       ├── 📂 ryan/             # Ryan's workbook modules
│   │       │   ├── Module_Dashboard.bas
│   │       │   ├── Module_Identity.bas  # Sets WORKBOOK_IDENTITY = "RZ"
│   │       │   ├── Sheet2 (SQRCT Dashboard).cls
│   │       │   ├── ThisWorkbook.bas
│   │       │   ├── backup/          # Historical VBA backups
│   │       │   └── mod*.bas         # Core module copies
│   │       └── 📂 sync_tool/        # Synchronization tool modules
│   │           ├── Module_SyncTool_Manager.bas
│   │           ├── Module_File_Processor.bas
│   │           ├── Module_Conflict_Handler.bas
│   │           ├── Module_Constants.bas
│   │           ├── Module_Format_Helpers.bas
│   │           ├── Module_StartUp.bas
│   │           ├── Module_SyncTool_Logger.bas
│   │           ├── Module_SyncTool_UI.bas
│   │           ├── Module_UIHandlers.bas
│   │           ├── Module_Utilities.bas
│   │           └── ThisWorkbook.cls
│   └── 📂 power_query/              # Power Query M language scripts
│       ├── Query - CSVQuotes.pq     # CSV data ingestion
│       ├── Query - ExistingQuotes.pq # Historical data query
│       ├── Query - MasterQuotes_Raw.pq # Data combination
│       ├── Query - MasterQuotes_Final.pq # Final processing
│       ├── DocNum_LatestLocation.pq # Document tracking
│       ├── Map_Form_DocNum.pq       # Form mapping
│       ├── OrderConf_*.pq           # Order confirmation queries
│       └── Query - CLIENT QUOTES.pg  # Client quote query
├── 📂 docs/                         # Documentation
│   ├── 📄 ARCHITECTURE.md           # Detailed technical architecture
│   ├── 📂 updates/                  # Project update history
│   │   ├── SQRCT - Update 041725.txt
│   │   ├── SQRCT - Update 041825.txt
│   │   ├── SQRCT - Update 041825-2.txt
│   │   └── SQRCT - Update 042025.txt
│   └── 📂 word/                     # Word document archives
│       ├── ARCHITECTURE.md.docx
│       ├── CONTRIBUTING.md.docx
│       └── readme.txt.docx
└── 📂 archives/                     # Historical files
    └── 📂 commits/                  # Commit history
        ├── commit_message.txt
        ├── commit_summary_20250421.txt
        └── folder_hierarchy.txt
```

## 🚀 Installation

### Prerequisites

- **Microsoft Excel 2016+** with:
  - Power Query support
  - VBA macro support enabled
  - Trust Center settings configured for macros
- **Network Access** to:
  - CSV quote export location
  - Shared workbook locations
- **Git** (for version control of VBA/M code)

### Setup Instructions

1. **Clone the Repository**
   ```bash
   git clone <repository-url>
   cd SQRCT
   ```

2. **Excel Workbook Setup**
   > ⚠️ **Important**: VBA code must be manually imported into Excel workbooks
   
   For each workbook (Ryan, Ally, Master):
   - Open the `.xlsm` file in Excel
   - Press `Alt+F11` to open VBA Editor
   - Import modules from corresponding `src/vba/workbooks/[user]/` folder
   - Import shared modules from `src/vba/core/`
   - Save the workbook

3. **Power Query Configuration**
   - Open Power Query Editor (`Data` → `Get Data` → `Launch Power Query Editor`)
   - Import `.pq` files from `src/power_query/`
   - Update file paths in queries to match your environment
   - Refresh all queries to test connectivity

4. **SyncTool Setup**
   - Open the SyncTool.xlsm workbook
   - Import all modules from `src/vba/workbooks/sync_tool/`
   - Configure file paths on the SyncTool dashboard

## 💻 Usage

### Daily Workflow

1. **Morning Data Refresh**
   - Open your assigned workbook (Ryan.xlsm or Ally.xlsm)
   - Click "Standard Refresh" button on SQRCT Dashboard
   - System will:
     - Save any existing edits
     - Refresh Power Query data
     - Restore your edits
     - Update Active/Archive views

2. **Quote Management**
   - Review quotes on SQRCT Dashboard
   - Edit engagement columns (K-N):
     - **K**: Engagement Phase (dropdown with validation)
     - **L**: Last Contact Date
     - **M**: Email Contact
     - **N**: User Comments
   - Changes are automatically saved to hidden UserEdits sheet

3. **Synchronization** (Team Lead)
   - Open SyncTool.xlsm
   - Verify file paths are correct
   - Click "Start Sync" button
   - Review any conflicts in MergeData sheet
   - Confirm synchronization
   - Master workbook is updated with all user edits

### Navigation

- **SQRCT Dashboard**: Main working view with all quotes
- **SQRCT Active**: Filtered view of quotes requiring action
- **SQRCT Archive**: Historical/completed quotes
- **Row 2 Controls**: Refresh buttons, view counts, last update timestamp

## 🔧 Development

### Code Organization

#### VBA Module Structure
```vba
'=====================================================================
' Module :  ModuleName
' Purpose:  Clear description of module purpose
' REVISED:  Date - Description of changes
'=====================================================================

Option Explicit

' Module-level constants
Private Const CONSTANT_NAME As String = "value"

' Public procedures
Public Sub MainProcedure()
    ' Implementation
End Sub
```

#### Power Query Best Practices
```m
let
    // Step 1: Clear description
    Source = ...,
    
    // Step 2: Transformation description
    Transformed = ...,
    
    // Final step
    Result = ...
in
    Result
```

### Making Changes

1. **VBA Code Changes**
   - Export module from Excel VBA Editor
   - Make changes in text editor
   - Commit to Git
   - Import updated module back to Excel
   - Test thoroughly

2. **Power Query Changes**
   - Edit in Power Query Editor
   - Export M code to `.pq` file
   - Commit to Git
   - Document any schema changes

3. **Testing Protocol**
   - Test in isolated workbook copy first
   - Verify data integrity
   - Check synchronization behavior
   - Document test results

### Debugging

- **VBA**: Use `Debug.Print` statements and Immediate Window
- **Power Query**: Use Table.Buffer() to materialize intermediate results
- **SyncTool**: Check SyncLog and ErrorLog sheets for details

## 🤝 Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for detailed guidelines.

### Quick Contribution Guide

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Make changes following code standards
4. Test thoroughly
5. Commit with clear messages (`git commit -m 'Add AmazingFeature'`)
6. Push to branch (`git push origin feature/AmazingFeature`)
7. Open a Pull Request

## 📝 Changelog

### Version 4.0.0 (July 2025)
- Complete repository restructure to gold standard organization
- Enhanced documentation for AI-assisted development
- Improved error handling and logging
- Added comprehensive conflict resolution in SyncTool

### Version 3.0.0 (April 2025)
- Implemented multi-user workbook architecture
- Added SyncTool for edit consolidation
- Enhanced Active/Archive view separation

### Version 2.0.0 (March 2025)
- Integrated Power Query for data processing
- Automated stage calculation logic
- Added user edit tracking

### Version 1.0.0 (February 2025)
- Initial release
- Basic quote tracking functionality
- Manual data entry system

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Sales team for requirements and feedback
- IT department for infrastructure support
- All contributors who have helped improve SQRCT

## 📧 Support

For issues, questions, or suggestions:
- Create an issue in the repository
- Contact the development team
- See [ARCHITECTURE.md](docs/ARCHITECTURE.md) for technical details

---

*Last updated: August 2025 | SQRCT v4.0.0*