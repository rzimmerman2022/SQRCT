# Contributing to SQRCT

## Table of Contents

- [Welcome](#welcome)
- [Getting Started](#getting-started)
- [Development Environment](#development-environment)
- [Code Organization](#code-organization)
- [Coding Standards](#coding-standards)
- [Development Workflow](#development-workflow)
- [Testing Guidelines](#testing-guidelines)
- [Submitting Changes](#submitting-changes)
- [Code Review Process](#code-review-process)
- [Security Guidelines](#security-guidelines)
- [Documentation](#documentation)
- [Community](#community)

## Welcome

Thank you for your interest in contributing to the Strategic Quote Recovery & Conversion Tracker (SQRCT) project! This guide provides comprehensive information for developers, including AI coding assistants, to effectively contribute to the project.

### Project Overview

SQRCT is an Excel-based quote tracking system that uses:
- **VBA** for business logic and UI
- **Power Query (M)** for data processing
- **Git** for version control (manual export/import workflow)

## Getting Started

### Prerequisites

1. **Software Requirements**
   - Microsoft Excel 2016+ (with Power Query and VBA support)
   - Git 2.x or higher
   - Text editor with VBA syntax highlighting (recommended: VS Code with VBA extension)

2. **Access Requirements**
   - Repository access (read/write permissions)
   - Test data access (anonymized quote data)
   - Network paths for CSV sources (test environment)

3. **Knowledge Requirements**
   - Intermediate VBA programming
   - Basic Power Query (M language)
   - Git version control
   - Excel object model understanding

### Initial Setup

```bash
# 1. Clone the repository
git clone <repository-url>
cd SQRCT

# 2. Create your feature branch
git checkout -b feature/your-feature-name

# 3. Set up your Excel development environment
# See "Development Environment" section below
```

## Development Environment

### Excel Configuration

1. **Macro Security Settings**
   ```
   File → Options → Trust Center → Trust Center Settings
   - Enable all macros (for development only)
   - Trust access to the VBA project object model
   ```

2. **VBA Editor Settings**
   ```
   Tools → Options
   - Editor tab:
     ✓ Auto Syntax Check
     ✓ Require Variable Declaration
     ✓ Auto List Members
     ✓ Auto Quick Info
   - Editor Format tab:
     - Font: Consolas, 10pt
     - Tab Width: 4
   ```

3. **References Setup**
   In VBA Editor → Tools → References, ensure these are checked:
   - Microsoft Excel 16.0 Object Library
   - Microsoft Office 16.0 Object Library
   - Microsoft Scripting Runtime
   - Microsoft VBScript Regular Expressions 5.5

### Power Query Setup

1. **Enable Power Query**
   - Data tab → Get & Transform Data group should be visible
   - If not, enable through Excel Options → Add-ins

2. **Configure Query Options**
   - Data → Get Data → Query Options
   - Privacy: Set to "Ignore Privacy Levels" for development
   - Data Load: Disable background refresh for debugging

## Code Organization

### Directory Structure

```
SQRCT/
├── src/
│   ├── vba/
│   │   ├── core/                    # Shared modules (ALWAYS import these)
│   │   │   ├── modArchival.bas      # View management functions
│   │   │   ├── modFormatting.bas    # UI formatting utilities
│   │   │   ├── modUtilities.bas     # General utilities
│   │   │   └── modPerformanceDashboard.bas  # Metrics tracking
│   │   └── workbooks/
│   │       ├── ally/                # Ally-specific modules
│   │       ├── master/              # Master workbook modules
│   │       ├── ryan/                # Ryan-specific modules
│   │       └── sync_tool/           # Synchronization tool
│   └── power_query/
│       ├── Query - *.pq             # Individual query files
│       └── OrderConf_*.pq           # Order confirmation queries
├── docs/                            # Documentation
├── archives/                        # Historical files
└── [root files]                     # README, LICENSE, etc.
```

### Module Dependencies

```
Core Modules (Required in all workbooks):
├── modUtilities.bas (no dependencies)
├── modFormatting.bas (depends on: modUtilities)
├── modArchival.bas (depends on: modUtilities, modFormatting)
└── modPerformanceDashboard.bas (depends on: modUtilities)

Workbook Modules:
├── Module_Dashboard.bas (depends on: all core modules)
├── Module_Identity.bas (standalone)
└── Sheet modules (depends on: Module_Dashboard)
```

## Coding Standards

### VBA Standards

#### Naming Conventions

```vba
' Constants
Private Const MAX_ROWS As Long = 100000
Public Const DASHBOARD_SHEET_NAME As String = "SQRCT Dashboard"

' Variables
Dim currentRow As Long
Dim docNumber As String
Dim isValid As Boolean

' Procedures
Public Sub RefreshDashboard()
Private Function ValidateDocumentNumber(docNum As String) As Boolean

' Objects
Dim wsData As Worksheet
Dim rngTarget As Range
```

#### Code Structure

```vba
'=====================================================================
' Module:   ModuleName
' Purpose:  Clear, concise description of module purpose
' Author:   Developer Name
' Created:  YYYY-MM-DD
' Modified: YYYY-MM-DD - Description of changes
'=====================================================================

Option Explicit

'---------------------------------------------------------------------
' Module-level declarations
'---------------------------------------------------------------------
Private Const MODULE_NAME As String = "ModuleName"

'---------------------------------------------------------------------
' Public procedures
'---------------------------------------------------------------------
Public Sub MainProcedure()
    ' Purpose: What this procedure does
    ' Parameters: List parameters if any
    ' Returns: What it returns if Function
    
    On Error GoTo ErrorHandler
    
    ' Main logic here
    
    Exit Sub
    
ErrorHandler:
    HandleError "MainProcedure", Err.Number, Err.Description
End Sub

'---------------------------------------------------------------------
' Private procedures
'---------------------------------------------------------------------
Private Sub HandleError(procName As String, errNum As Long, errDesc As String)
    ' Centralized error handling
    Debug.Print MODULE_NAME & "." & procName & " Error: " & errNum & " - " & errDesc
    MsgBox "An error occurred. Please check the logs.", vbExclamation
End Sub
```

#### Best Practices

1. **Always use Option Explicit**
2. **Declare variables at the smallest scope needed**
3. **Use meaningful variable names (no single letters except loop counters)**
4. **Implement error handling in all Public procedures**
5. **Avoid GoTo except for error handling**
6. **Use early binding when possible**
7. **Comment complex logic, not obvious code**

### Power Query Standards

#### Query Structure

```m
let
    //========================================
    // CONFIGURATION
    //========================================
    FilePath = "\\network\path\to\data.csv",
    
    //========================================
    // DATA INGESTION
    //========================================
    Source = Csv.Document(
        File.Contents(FilePath),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]
    ),
    
    //========================================
    // TRANSFORMATION
    //========================================
    // Step 1: Promote headers
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    
    // Step 2: Change types
    ChangedTypes = Table.TransformColumnTypes(
        PromotedHeaders,
        {
            {"Document Number", type text},
            {"Amount", Currency.Type},
            {"Date", type date}
        }
    ),
    
    // Step 3: Filter and clean
    FilteredRows = Table.SelectRows(
        ChangedTypes, 
        each [Amount] > 0 and [Date] <> null
    )
in
    FilteredRows
```

#### Best Practices

1. **Use clear step names that describe the transformation**
2. **Group related steps with comments**
3. **Avoid hardcoded values - use parameters**
4. **Handle errors gracefully with try...otherwise**
5. **Document data source assumptions**
6. **Optimize for performance (filter early, transform late)**

## Development Workflow

### Branch Workflow

```bash
# 1. Start from updated main
git checkout main
git pull origin main

# 2. Create feature branch
git checkout -b feature/quote-export-enhancement

# 3. Make changes and commit
git add .
git commit -m "feat(export): add CSV export for archived quotes"

# 4. Push branch
git push origin feature/quote-export-enhancement

# 5. Create Pull Request on GitHub
```

### Commit Message Format

Follow [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<scope>): <subject>

<body>

<footer>
```

#### Types
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `style`: Code style changes (formatting, etc.)
- `refactor`: Code refactoring
- `perf`: Performance improvements
- `test`: Adding tests
- `chore`: Maintenance tasks

#### Examples
```
feat(dashboard): add real-time quote age calculation

- Calculate age dynamically based on current date
- Display in new column with conditional formatting
- Update refresh logic to recalculate on demand

Closes #123
```

## Testing Guidelines

### Manual Testing Protocol

#### 1. Unit Testing (VBA)
```vba
' Test individual functions
Sub Test_ValidateDocumentNumber()
    Debug.Assert ValidateDocumentNumber("SMOQ12345") = True
    Debug.Assert ValidateDocumentNumber("INVALID") = False
    Debug.Assert ValidateDocumentNumber("") = False
    Debug.Print "ValidateDocumentNumber tests passed"
End Sub
```

#### 2. Integration Testing
- Test complete workflows (refresh → edit → sync)
- Verify data integrity across workbooks
- Test error scenarios

#### 3. User Acceptance Testing
- Have actual users test in sandbox environment
- Document feedback and issues
- Verify performance with production-size data

### Test Documentation

Create test documents in the following format:

```markdown
# Test Case: Dashboard Refresh

## Objective
Verify dashboard refresh maintains user edits

## Prerequisites
- Test workbook with sample data
- UserEdits sheet populated

## Steps
1. Open test workbook
2. Make edits in columns K-N
3. Click "Standard Refresh"
4. Verify edits persist

## Expected Results
- Power Query data updates
- User edits remain intact
- No errors displayed

## Actual Results
[Fill during testing]

## Pass/Fail
[Mark after testing]
```

## Submitting Changes

### Pre-Submission Checklist

- [ ] Code follows VBA/Power Query standards
- [ ] All procedures have error handling
- [ ] Complex logic is documented
- [ ] Manual testing completed
- [ ] Documentation updated
- [ ] No hardcoded paths/credentials
- [ ] Commit messages follow convention

### Pull Request Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Tested in development environment
- [ ] Tested with production-size data
- [ ] Regression testing completed

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No merge conflicts

## Screenshots
[If applicable]
```

## Code Review Process

### For Authors

1. **Self-review first** - Check your own PR before requesting review
2. **Provide context** - Explain why changes were made
3. **Be responsive** - Address feedback promptly
4. **Test suggestions** - Verify reviewer suggestions work

### For Reviewers

1. **Be constructive** - Suggest improvements, don't just criticize
2. **Check functionality** - Pull code and test locally
3. **Verify standards** - Ensure coding standards are followed
4. **Consider performance** - Flag potential performance issues

### Review Checklist

- [ ] **Functionality**: Does it work as intended?
- [ ] **Code Quality**: Is it readable and maintainable?
- [ ] **Standards**: Does it follow our conventions?
- [ ] **Error Handling**: Are errors handled gracefully?
- [ ] **Security**: Are there any security concerns?
- [ ] **Performance**: Will it scale appropriately?
- [ ] **Documentation**: Is it properly documented?

## Security Guidelines

### Code Security

1. **Never hardcode credentials**
   ```vba
   ' BAD
   Const DB_PASSWORD As String = "mypassword123"
   
   ' GOOD
   Dim dbPassword As String
   dbPassword = Environ("SQRCT_DB_PASSWORD")
   ```

2. **Validate all inputs**
   ```vba
   Function SafeGetCell(ws As Worksheet, row As Long, col As Long) As Variant
       On Error GoTo SafeExit
       If row > 0 And row <= ws.Rows.Count And col > 0 And col <= ws.Columns.Count Then
           SafeGetCell = ws.Cells(row, col).Value
       End If
   SafeExit:
   End Function
   ```

3. **Use least privilege principle**
   - Only request necessary file permissions
   - Limit macro capabilities to required functions

### Data Security

1. **Anonymize test data** - Remove real customer information
2. **Secure file paths** - Use UNC paths with proper permissions
3. **Audit trail** - Maintain logs of data modifications

## Documentation

### Code Documentation

```vba
'---------------------------------------------------------------------
' Procedure: CalculateQuoteAge
' Purpose:   Calculates the age of a quote in business days
' 
' Parameters:
'   quoteDate (Date): The date the quote was created
'   endDate (Date): Optional end date, defaults to today
'
' Returns:
'   Long: Number of business days between dates
'
' Example:
'   daysOld = CalculateQuoteAge(#1/1/2024#)
'
' Notes:
'   - Excludes weekends
'   - Does not account for holidays
'---------------------------------------------------------------------
Public Function CalculateQuoteAge(quoteDate As Date, Optional endDate As Date) As Long
    ' Implementation here
End Function
```

### User Documentation

When adding features, update:
1. **README.md** - High-level feature description
2. **ARCHITECTURE.md** - Technical implementation details
3. **In-app help** - User-facing instructions

## Community

### Getting Help

- **GitHub Issues**: Bug reports and feature requests
- **Discussions**: General questions and ideas
- **Code Comments**: Check inline documentation
- **Architecture Doc**: Technical details in `/docs/ARCHITECTURE.md`

### Code of Conduct

- Be respectful and inclusive
- Welcome newcomers
- Focus on constructive feedback
- Assume positive intent

### Recognition

Contributors are recognized in:
- Git commit history
- Pull request mentions
- Release notes
- Annual contributor summary

---

## Quick Reference

### Common Tasks

```bash
# Update from main
git pull origin main

# Run VBA tests
# Open Excel → Alt+F11 → Run Test_All procedure

# Export VBA module
# VBA Editor → Right-click module → Export

# Import Power Query
# Power Query Editor → Home → Advanced Editor → Copy/Paste

# Create release
git tag -a v1.2.3 -m "Release version 1.2.3"
git push origin v1.2.3
```

### Helpful Resources

- [Excel VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Power Query M Reference](https://docs.microsoft.com/en-us/powerquery-m/)
- [Git Documentation](https://git-scm.com/doc)
- [Conventional Commits](https://www.conventionalcommits.org/)

---

*Thank you for contributing to SQRCT! Your efforts help improve quote management for the entire sales team.*