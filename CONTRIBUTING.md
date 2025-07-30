# Contributing to SQRCT

Thank you for your interest in contributing to the Strategic Quote Recovery & Conversion Tracker (SQRCT) project! This guide will help you understand our development process and coding standards.

## 🚀 Getting Started

### Prerequisites
- Microsoft Excel (with Power Query and VBA support)
- Git for version control
- Basic understanding of VBA and Power Query (M language)
- Access to development/test data sources

### Development Environment Setup
1. **Clone the repository**
2. **Set up Excel with macro security** enabled for development
3. **Configure VBA Editor** with proper settings:
   - Enable "Require Variable Declaration"
   - Set appropriate error trapping options
   - Configure indentation (4 spaces recommended)

## 📁 Code Organization

### Repository Structure
Our codebase follows gold standard organization:

- **`src/vba/core/`** - Shared VBA modules used across all workbooks
- **`src/vba/workbooks/`** - Workbook-specific VBA implementations
- **`src/power_query/`** - Power Query M language scripts
- **`docs/`** - All documentation files
- **`tests/`** - Test plans and sample data

### Working with VBA Code
1. **Extract VBA modules** from Excel workbooks using the VBA Editor
2. **Save as `.bas`, `.cls`, or `.frm` files** in appropriate directories
3. **Commit changes** to the repository
4. **Import updated code** back to Excel files for testing

### Working with Power Query
1. **Export M scripts** as `.pq` text files
2. **Store in `src/power_query/`** with descriptive names
3. **Document dependencies** and data source requirements
4. **Test transformations** with sample data

## 🎯 Coding Standards

### VBA Guidelines
- **Use explicit variable declarations** (`Option Explicit`)
- **Follow consistent naming conventions**:
  - Variables: `camelCase` (e.g., `currentRow`, `docNumber`)
  - Constants: `UPPER_CASE` (e.g., `MAX_ROWS`, `DEFAULT_PATH`)
  - Functions/Subs: `PascalCase` (e.g., `RefreshDashboard`, `ProcessData`)
- **Include error handling** using `On Error GoTo` statements
- **Add meaningful comments** for complex logic
- **Use consistent indentation** (4 spaces)

### Power Query (M) Standards
- **Use descriptive step names** in queries
- **Document complex transformations** with comments
- **Avoid hardcoded values** where possible
- **Use proper data type conversions**
- **Structure queries** for readability and maintainability

### Documentation Standards
- **Update README.md** when adding new features
- **Document API changes** in ARCHITECTURE.md
- **Include inline comments** for complex code sections
- **Maintain changelog** for version tracking

## 🔄 Development Workflow

### Branching Strategy
1. **Create feature branches** from `main`: `feature/description`
2. **Create bugfix branches** from `main`: `bugfix/description`
3. **Use descriptive branch names** that explain the change

### Commit Message Format
Follow conventional commit format:
```
type(scope): description

[optional body]

[optional footer]
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

Examples:
- `feat(dashboard): add new performance metrics display`
- `fix(sync): resolve conflict resolution timestamp logic`
- `docs(readme): update installation instructions`

### Pull Request Process
1. **Ensure all tests pass** (manual testing for VBA/Excel)
2. **Update documentation** as needed
3. **Provide clear PR description** with:
   - What was changed and why
   - How to test the changes
   - Any breaking changes or migration notes
4. **Request review** from project maintainers
5. **Address feedback** and update as needed

## 🧪 Testing Guidelines

### Manual Testing Requirements
- **Test VBA functionality** in Excel environment
- **Verify Power Query transformations** with sample data
- **Test user workflows** end-to-end
- **Validate data accuracy** after synchronization
- **Check error handling** scenarios

### Test Documentation
- **Document test cases** in `tests/` directory
- **Include sample data** for testing scenarios
- **Maintain test checklists** for release validation
- **Record known issues** and workarounds

## 🐛 Bug Reports

When reporting bugs, please include:
- **Excel version** and environment details
- **Steps to reproduce** the issue
- **Expected vs actual behavior**
- **Screenshots** or error messages if applicable
- **Sample data** (anonymized) if relevant

## 💡 Feature Requests

For new features, please provide:
- **Clear use case** and business justification
- **Detailed requirements** and acceptance criteria
- **Mockups or examples** if applicable
- **Consideration of existing architecture**

## 📋 Code Review Checklist

### For Reviewers
- [ ] Code follows established conventions
- [ ] Error handling is appropriate
- [ ] Documentation is updated
- [ ] Changes are backward compatible
- [ ] Security considerations are addressed
- [ ] Performance impact is acceptable

### For Contributors
- [ ] Code is properly formatted
- [ ] Variables are properly declared
- [ ] Error handling is implemented
- [ ] Documentation is updated
- [ ] Testing has been performed
- [ ] Commit messages are clear

## 🔒 Security Guidelines

- **Never commit** sensitive data or credentials
- **Use configuration files** for environment-specific settings
- **Follow principle of least privilege** for file access
- **Validate user inputs** to prevent injection attacks
- **Use Excel's built-in security features** appropriately

## 📞 Getting Help

- **GitHub Issues** - For bug reports and feature requests
- **GitHub Discussions** - For questions and community support
- **Documentation** - Check `docs/ARCHITECTURE.md` for technical details
- **Code Comments** - Review inline documentation in source files

## 📄 License

By contributing to SQRCT, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to SQRCT! Your efforts help improve the quote management process for the entire team.