# ExcelPlusPlus Unit Tests

Comprehensive unit tests for the ExcelPlusPlus OutSystems component.

## Test Framework

- **Framework**: MSTest (Microsoft Visual Studio Unit Testing Framework)
- **Target Framework**: .NET Framework 4.8
- **Dependencies**: EPPlus 8.5.4, OutSystems Runtime assemblies

## Test Coverage

The test suite covers **80+ public methods** across the following categories:

### Workbook Operations (12 tests)
- Create workbook with single/multiple sheets
- Open workbook from file or binary data
- Close/dispose workbook
- Get binary data
- Protect workbook with password
- Get workbook properties
- Change sheet index

### Worksheet Operations (18 tests)
- Add/select/rename/delete worksheets
- Copy worksheets
- Set active worksheet
- Hide/show worksheets
- Protect worksheets
- Get worksheet properties/name
- Calculate formulas
- Add auto filter
- Autofit columns
- Copy rows
- Add dropdown lists
- Set/get headers and footers

### Cell Operations (24 tests)
- Write cells by index/name (various types: text, integer, decimal, boolean, formula)
- Read cells by index/name
- Set/read formulas
- Calculate cells
- Clear cell values
- Format cell ranges
- Merge/unmerge cells
- Write/read images
- Get fill colors
- Convert hex to RGB

### Row/Column Operations (10 tests)
- Set width/height
- Insert/delete rows/columns
- Hide/show rows/columns

### Address Operations (6 tests)
- Convert text address to row/column
- Convert row/column to text address
- Parse single cells and ranges

### Comment Operations (2 tests)
- Add comments
- Delete comments

### Conditional Formatting (5 tests)
- Add formatting rules
- Get all rules
- Delete single rule
- Delete all rules

### Chart Operations (2 tests)
- Create column charts
- Create pie charts

### Image Operations (4 tests)
- Insert images by row/column
- Insert images by cell name
- Get images from worksheet

### Excel Properties (3 tests)
- Get document properties
- Set document properties
- Clear document properties

### Data Operations (2 tests)
- Write ranges
- Write column ranges

### Find/Search Operations (4 tests)
- Find cells by value
- Check if value exists in range

### Cell Format Operations (1 test)
- Apply format to range

### Additional Operations (6 tests)
- Add sheets at index
- Add sheet copies
- Select worksheets
- Delete worksheets

## Prerequisites

Before running tests, ensure you have:

1. **Visual Studio 2019/2022** or **MSBuild 15.0+**
2. **.NET Framework 4.8 SDK**
3. **EPPlus 8.5.4** (via NuGet)
4. **OutSystems Runtime assemblies** (if testing integration)

## Setup

### Option 1: Visual Studio

1. Open `ExcelPlusPlus.Tests.csproj` in Visual Studio
2. Restore NuGet packages (right-click solution → Restore NuGet Packages)
3. Build the test project
4. Run tests via Test Explorer (Ctrl+E, T)

### Option 2: Command Line (MSBuild)

```bash
# Navigate to Tests directory
cd ExcelPlusPlus/Tests

# Restore NuGet packages
nuget restore ExcelPlusPlus.Tests.csproj

# Build the test project
msbuild ExcelPlusPlus.Tests.csproj /p:Configuration=Debug

# Run tests with vstest.console
vstest.console.exe bin\Debug\ExcelPlusPlus.Tests.dll
```

### Option 3: .NET CLI (if using .NET Core/5+)

Note: This project targets .NET Framework 4.8, so .NET CLI may have limited support.

```bash
dotnet test ExcelPlusPlus.Tests.csproj
```

## Test Structure

All tests follow the **Arrange-Act-Assert** pattern:

```csharp
[TestMethod]
public void TestName_Description()
{
    // Arrange - Setup test data and preconditions
    var package = CreateTestWorkbook();
    var worksheet = GetFirstWorksheet(package);
    
    // Act - Execute the method being tested
    _excelPlusPlus.MssCell_WriteByIndex(worksheet, 1, 1, "TestValue", "text");
    
    // Assert - Verify the expected outcome
    Assert.AreEqual("TestValue", worksheet.Cells[1, 1].Value?.ToString());
}
```

## EPPlus License Context

All tests set `ExcelPackage.LicenseContext = LicenseContext.NonCommercial` in the test initialization to comply with EPPlus licensing requirements for the free community version.

```csharp
[TestInitialize]
public void TestInitialize()
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    _excelPlusPlus = new CssExcelPlusPlus();
}
```

## Temporary File Management

Tests that create Excel files use temporary files that are automatically cleaned up:

```csharp
[TestCleanup]
public void TestCleanup()
{
    if (File.Exists(_tempFilePath))
    {
        File.Delete(_tempFilePath);
    }
}
```

## Running Specific Tests

### Visual Studio Test Explorer
- Filter by name, category, or status
- Run individual tests, classes, or all tests

### Command Line
```bash
# Run all tests
vstest.console.exe ExcelPlusPlus.Tests.dll

# Run tests matching a pattern
vstest.console.exe ExcelPlusPlus.Tests.dll /Tests:Workbook

# Run tests in a specific class
vstest.console.exe ExcelPlusPlus.Tests.dll /Tests:ExcelPlusPlusTests.Workbook
```

## Expected Results

- **Total Tests**: 120+
- **Expected Pass Rate**: 100% (all tests should pass)
- **Test Duration**: ~5-10 seconds (depending on hardware)

## Known Limitations

1. **OutSystems Dependencies**: Some tests may require OutSystems runtime assemblies to be present in the GAC or referenced explicitly.

2. **EPPlus Version**: Tests are written against EPPlus 8.5.4. Behavior may vary with different versions.

3. **Platform-Specific**: Tests are designed for Windows/.NET Framework 4.8. Running on Mono or .NET Core may require adjustments.

## Adding New Tests

When adding tests for new functionality:

1. Follow the existing naming convention: `Mss[Class]_[Method]_[Scenario]`
2. Use the Arrange-Act-Assert pattern
3. Test both happy-path and error cases
4. Clean up any resources (files, objects) in TestCleanup
5. Add XML documentation comments explaining the test purpose

## Troubleshooting

### Test fails with "LicenseContext not set"
Ensure `ExcelPackage.LicenseContext` is set before any EPPlus operations.

### Test fails with "OutSystems assembly not found"
Add references to OutSystems runtime assemblies or mock the dependencies.

### Test fails with "File in use"
Ensure all ExcelPackage objects are properly disposed in test cleanup.

## Contributing

When contributing new tests:
- Maintain >90% code coverage for new features
- Document any assumptions or prerequisites
- Keep tests isolated and independent
- Use descriptive test names that explain the scenario

## License

This test suite is licensed under the same BSD-3-Clause license as the ExcelPlusPlus component.
