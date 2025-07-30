# SQRCT - Strategic Quote Recovery & Conversion Tracker

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![VBA](https://img.shields.io/badge/VBA-Excel-green.svg)](https://docs.microsoft.com/en-us/office/vba/)
[![Power Query](https://img.shields.io/badge/Power%20Query-M%20Language-blue.svg)](https://docs.microsoft.com/en-us/power-query/)

## 📋 Project Overview

SQRCT is an enterprise-grade Excel-based system designed to track and manage the follow-up process for sales quotes. The system improves quote conversion rates by providing visibility into outstanding quotes, automating follow-up stage calculation, and facilitating consistent engagement by the sales team.

### 🏗️ System Architecture

The system operates within Microsoft Excel using:
- **Power Query (M Language)** for data ingestion and transformation
- **Visual Basic for Applications (VBA)** for automation and user interfaces
- **Multi-workbook architecture** for scalable user management

## 📁 Repository Structure

```
SQRCT/
├── 📂 src/                          # Source code
│   ├── 📂 vba/                      # VBA modules and classes
│   │   ├── 📂 core/                 # Shared VBA modules
│   │   │   ├── modArchival.bas      # Archive management
│   │   │   ├── modFormatting.bas    # UI formatting
│   │   │   ├── modUtilities.bas     # Utility functions
│   │   │   └── modPerformanceDashboard.bas  # Performance metrics
│   │   └── 📂 workbooks/            # Workbook-specific VBA
│   │       ├── 📂 ally/             # Ally's workbook modules
│   │       ├── 📂 master/           # Master workbook modules
│   │       ├── 📂 ryan/             # Ryan's workbook modules
│   │       │   └── 📂 backup/       # Backup VBA files
│   │       └── 📂 sync_tool/        # Synchronization tool modules
│   └── 📂 power_query/              # Power Query M scripts
│       ├── Query - CSVQuotes.pq     # CSV data ingestion
│       ├── Query - ExistingQuotes.pq # Historical data
│       ├── Query - MasterQuotes_Final.pq # Final processing
│       └── [other .pq files]        # Additional queries
├── 📂 docs/                         # Documentation
│   ├── 📂 updates/                  # Project update logs
│   └── 📂 word/                     # Word document archives
├── 📂 archives/                     # Historical files
│   └── 📂 commits/                  # Commit history files
├── 📂 tests/                        # Test plans and data
├── 📂 scripts/                      # Utility scripts
├── 📂 .github/                      # GitHub workflows
├── ARCHITECTURE.md                  # Technical architecture
├── README.md                        # This file
└── .gitignore                      # Git ignore rules
```

## 🚀 Quick Start

### Prerequisites

- **Microsoft Excel** (with Power Query and VBA macro support)
- **Git** for version control
- **Network access** to required file locations (see Configuration)

### Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd SQRCT
   ```

2. **Access the source code:**
   - **VBA modules:** Located in `src/vba/` organized by workbook type
   - **Power Query scripts:** Located in `src/power_query/`
   - **Core shared modules:** Located in `src/vba/core/`

3. **Apply changes to Excel workbooks:**
   > ⚠️ **Manual Process**: Changes must be manually imported into operational `.xlsm/.xlsx` files using Excel's VBA Editor

## 💻 Development Workflow

### Code Organization

- **`src/vba/core/`** - Shared VBA modules used across workbooks
- **`src/vba/workbooks/`** - Workbook-specific VBA code
- **`src/power_query/`** - M language scripts for data processing

### Working with VBA Code

1. **Extract code changes** from Excel workbooks
2. **Commit to repository** for version control
3. **Import updated code** back to Excel files
4. **Test thoroughly** before deployment

### Power Query Development

1. **Edit M scripts** in `src/power_query/`
2. **Import queries** into Excel workbooks
3. **Test data transformation** logic
4. **Update documentation** as needed

## ⚙️ Configuration

### File Paths
- Power Query and SyncTool rely on specific network paths
- **Recommendation:** Use configuration sheets instead of hardcoded paths
- Document all file path dependencies

### Security
- **No hardcoded passwords** in VBA or M code
- Use Excel's built-in security features
- Rely on filesystem/SharePoint permissions for access control

## 🔄 Usage Workflow

### Daily Operations

1. **User Work:**
   - Users open their individual Excel workbooks
   - Make edits to engagement columns (K-N) on SQRCT Dashboard
   - VBA automatically saves changes to hidden "UserEdits" sheet

2. **Synchronization:**
   - Open SyncTool Excel workbook
   - Select file paths for user and master workbooks
   - Click "Sync" to merge all user edits
   - System resolves conflicts using timestamps

3. **Master Update:**
   - Master workbook refreshes with consolidated data
   - Power Query updates with latest source data
   - Dashboard displays merged results

## 🧪 Testing

### Current Testing Approach
- **Manual verification** of Power Query outputs
- **Functional testing** of VBA components
- **Integration testing** of synchronization process
- **User acceptance testing** with actual workflows

### Test Resources
- Use `tests/` directory for test plans and sample data
- Document test cases and expected outcomes
- Maintain test data separate from production

## 🤝 Contributing

We welcome contributions! Please:

1. **Read the contribution guidelines** in `docs/CONTRIBUTING.md`
2. **Follow VBA and M language coding standards**
3. **Use proper branching strategy** for changes
4. **Submit pull requests** with detailed descriptions
5. **Report issues** using GitHub Issues

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📧 Support & Contact

- **Issues:** Submit through [GitHub Issues](../../issues)
- **Questions:** Contact the development team
- **Documentation:** See `docs/ARCHITECTURE.md` for technical details

---

*Last updated: July 2025 | SQRCT v4.0.0*