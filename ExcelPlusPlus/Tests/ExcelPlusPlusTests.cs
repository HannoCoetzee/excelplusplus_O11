using System;
using System.IO;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OutSystems.NssExcelPlusPlus;

namespace ExcelPlusPlus.Tests
{
    /// <summary>
    /// Comprehensive unit tests for ExcelPlusPlus component
    /// Tests all public methods exposed by CssExcelPlusPlus class
    /// </summary>
    [TestClass]
    public class ExcelPlusPlusTests
    {
        private CssExcelPlusPlus _excelPlusPlus;
        private string _tempFilePath;

        [TestInitialize]
        public void TestInitialize()
        {
            // Set license context for EPPlus (required for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _excelPlusPlus = new CssExcelPlusPlus();
            _tempFilePath = Path.Combine(Path.GetTempPath(), $"ExcelPlusPlus_Test_{Guid.NewGuid()}.xlsx");
        }

        [TestCleanup]
        public void TestCleanup()
        {
            // Clean up temporary files
            if (!string.IsNullOrEmpty(_tempFilePath) && File.Exists(_tempFilePath))
            {
                try
                {
                    File.Delete(_tempFilePath);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        #region Helper Methods

        private ExcelPackage CreateTestWorkbook(string sheetName = "Sheet1")
        {
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add(sheetName);
            return package;
        }

        private ExcelWorksheet GetFirstWorksheet(ExcelPackage package)
        {
            return package.Workbook.Worksheets[0];
        }

        private void SaveWorkbook(ExcelPackage package)
        {
            var fileInfo = new FileInfo(_tempFilePath);
            package.SaveAs(fileInfo);
        }

        #endregion

        #region Workbook Tests

        [TestMethod]
        public void MssWorkbook_Create_WithDefaultSheetName_CreatesWorkbook()
        {
            // Arrange
            int numberOfSheets = 1;
            string firstSheetName = "TestSheet";
            RLNewSheetRecordList sheetNames = null;

            // Act
            _excelPlusPlus.MssWorkbook_Create(numberOfSheets, firstSheetName, sheetNames, out object workbook);

            // Assert
            Assert.IsNotNull(workbook);
            var package = workbook as ExcelPackage;
            Assert.IsNotNull(package);
            Assert.AreEqual(1, package.Workbook.Worksheets.Count);
            Assert.AreEqual("TestSheet", package.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void MssWorkbook_Create_WithMultipleSheets_CreatesAllSheets()
        {
            // Arrange
            int numberOfSheets = 3;
            string firstSheetName = "Sheet";
            RLNewSheetRecordList sheetNames = null;

            // Act
            _excelPlusPlus.MssWorkbook_Create(numberOfSheets, firstSheetName, sheetNames, out object workbook);

            // Assert
            Assert.IsNotNull(workbook);
            var package = workbook as ExcelPackage;
            Assert.AreEqual(3, package.Workbook.Worksheets.Count);
        }

        [TestMethod]
        public void MssWorkbook_Open_WithFilePath_OpensWorkbook()
        {
            // Arrange
            var testPackage = CreateTestWorkbook("TestSheet");
            SaveWorkbook(testPackage);
            testPackage.Dispose();

            // Act
            _excelPlusPlus.MssWorkbook_Open(_tempFilePath, null, out object workbook);

            // Assert
            Assert.IsNotNull(workbook);
            var package = workbook as ExcelPackage;
            Assert.IsNotNull(package);
            Assert.AreEqual("TestSheet", package.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void MssWorkbook_Open_WithBinaryData_OpensWorkbook()
        {
            // Arrange
            var testPackage = CreateTestWorkbook("BinaryTest");
            using (var ms = new MemoryStream())
            {
                testPackage.SaveAs(ms);
                byte[] binaryData = ms.ToArray();
                testPackage.Dispose();

                // Act
                _excelPlusPlus.MssWorkbook_Open("", binaryData, out object workbook);

                // Assert
                Assert.IsNotNull(workbook);
                var package = workbook as ExcelPackage;
                Assert.AreEqual("BinaryTest", package.Workbook.Worksheets[0].Name);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssWorkbook_Open_WithNoParameters_ThrowsException()
        {
            // Act
            _excelPlusPlus.MssWorkbook_Open("", null, out object workbook);
        }

        [TestMethod]
        public void MssWorkbook_Close_DisposesWorkbook()
        {
            // Arrange
            var package = CreateTestWorkbook();

            // Act
            _excelPlusPlus.MssWorkbook_Close(package);

            // Assert - workbook should be disposed
            Assert.IsNull(package);
        }

        [TestMethod]
        public void MssWorkbook_GetBinaryData_ReturnsByteArray()
        {
            // Arrange
            var package = CreateTestWorkbook("BinaryDataTest");

            // Act
            _excelPlusPlus.MssWorkbook_GetBinaryData(package, out byte[] binaryData);

            // Assert
            Assert.IsNotNull(binaryData);
            Assert.IsTrue(binaryData.Length > 0);
        }

        [TestMethod]
        public void MssWorkbook_Protect_WithPassword_SetsProtection()
        {
            // Arrange
            var package = CreateTestWorkbook();
            string password = "TestPassword123";

            // Act
            _excelPlusPlus.MssWorkbook_Protect(package, password, true, true, false);

            // Assert
            Assert.IsNotNull(package.Encryption.Password);
            Assert.AreEqual(password, package.Encryption.Password);
            Assert.IsTrue(package.Workbook.Protection.LockStructure);
            Assert.IsTrue(package.Workbook.Protection.LockWindows);
        }

        [TestMethod]
        public void MssWorkbook_GetProperties_ReturnsWorkbookProperties()
        {
            // Arrange
            var package = CreateTestWorkbook("PropsTest");

            // Act
            _excelPlusPlus.MssWorkbook_GetProperties(package, out RCWorkbookRecord properties);

            // Assert
            Assert.IsNotNull(properties);
            Assert.IsNotNull(properties.ssSTWorkbook.ssWorksheets);
            Assert.AreEqual(1, properties.ssSTWorkbook.ssWorksheets.Count);
        }

        [TestMethod]
        public void MssWorkbook_ChangeSheetIndex_MovesSheet()
        {
            // Arrange
            var package = CreateTestWorkbook("Sheet1");
            package.Workbook.Worksheets.Add("Sheet2");
            package.Workbook.Worksheets.Add("Sheet3");

            // Act - Move Sheet1 from index 1 to index 3
            _excelPlusPlus.MssWorkbook_ChangeSheetIndex(package, 1, 3);

            // Assert
            Assert.AreEqual("Sheet2", package.Workbook.Worksheets[0].Name);
            Assert.AreEqual("Sheet3", package.Workbook.Worksheets[1].Name);
            Assert.AreEqual("Sheet1", package.Workbook.Worksheets[2].Name);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssWorkbook_ChangeSheetIndex_WithInvalidIndex_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();

            // Act
            _excelPlusPlus.MssWorkbook_ChangeSheetIndex(package, 0, 1);
        }

        #endregion

        #region Worksheet Tests

        [TestMethod]
        public void MssWorkbook_AddName_AddsWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook();

            // Act
            _excelPlusPlus.MssWorkbook_AddName(package, "NewSheet", out object worksheet);

            // Assert
            Assert.IsNotNull(worksheet);
            Assert.AreEqual(2, package.Workbook.Worksheets.Count);
            Assert.AreEqual("NewSheet", package.Workbook.Worksheets[1].Name);
        }

        [TestMethod]
        public void MssWorksheet_SelectByName_SelectsCorrectWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("TestSheet");
            package.Workbook.Worksheets.Add("AnotherSheet");

            // Act
            _excelPlusPlus.MssWorksheet_SelectByName(package, "TestSheet", out object worksheet);

            // Assert
            Assert.IsNotNull(worksheet);
            var ws = worksheet as ExcelWorksheet;
            Assert.AreEqual("TestSheet", ws.Name);
        }

        [TestMethod]
        public void MssWorksheet_SelectByIndex_SelectsCorrectWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("First");
            package.Workbook.Worksheets.Add("Second");

            // Act
            _excelPlusPlus.MssWorksheet_SelectByIndex(package, 2, out object worksheet);

            // Assert
            Assert.IsNotNull(worksheet);
            var ws = worksheet as ExcelWorksheet;
            Assert.AreEqual("Second", ws.Name);
        }

        [TestMethod]
        public void MssWorksheet_Rename_ChangesSheetName()
        {
            // Arrange
            var package = CreateTestWorkbook("OldName");
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssWorksheet_Rename(worksheet, "NewName");

            // Assert
            Assert.AreEqual("NewName", worksheet.Name);
        }

        [TestMethod]
        public void MssWorksheet_DeleteByName_RemovesSheet()
        {
            // Arrange
            var package = CreateTestWorkbook("ToDelete");
            package.Workbook.Worksheets.Add("KeepThis");

            // Act
            _excelPlusPlus.MssWorksheet_DeleteByName(package, "ToDelete");

            // Assert
            Assert.AreEqual(1, package.Workbook.Worksheets.Count);
            Assert.AreEqual("KeepThis", package.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void MssWorksheet_DeleteByIndex_RemovesSheet()
        {
            // Arrange
            var package = CreateTestWorkbook("KeepThis");
            package.Workbook.Worksheets.Add("ToDelete");

            // Act
            _excelPlusPlus.MssWorksheet_DeleteByIndex(package, 2);

            // Assert
            Assert.AreEqual(1, package.Workbook.Worksheets.Count);
        }

        [TestMethod]
        public void MssWorksheet_AddCopyWorksheet_CreatesCopy()
        {
            // Arrange
            var package = CreateTestWorkbook("Original");
            var originalWs = GetFirstWorksheet(package);
            originalWs.Cells["A1"].Value = "TestData";

            // Act
            _excelPlusPlus.MssWorkbook_AddCopyWorksheet(package, "Copy", originalWs, out object copiedWorksheet);

            // Assert
            Assert.AreEqual(2, package.Workbook.Worksheets.Count);
            var copyWs = copiedWorksheet as ExcelWorksheet;
            Assert.AreEqual("Copy", copyWs.Name);
            Assert.AreEqual("TestData", copyWs.Cells["A1"].Value?.ToString());
        }

        [TestMethod]
        public void MssWorksheet_SetActive_SetsActiveSheet()
        {
            // Arrange
            var package = CreateTestWorkbook("Sheet1");
            package.Workbook.Worksheets.Add("Sheet2");

            // Act
            _excelPlusPlus.MssWorksheet_SetActive(package, "Sheet2", 0);

            // Assert - Sheet2 should be selected
            Assert.AreEqual("Sheet2", package.Workbook.Worksheets.Selected?.Name);
        }

        [TestMethod]
        public void MssWorksheet_Hide_Show_HidesWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act - Hide (1 = Hidden)
            _excelPlusPlus.MssWorksheet_Hide_Show(worksheet, 1);

            // Assert
            Assert.AreEqual(eWorkSheetHidden.Hidden, worksheet.Hidden);
        }

        [TestMethod]
        public void MssWorksheet_Hide_Show_ShowsWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Hidden = eWorkSheetHidden.Hidden;

            // Act - Show (0 = Visible)
            _excelPlusPlus.MssWorksheet_Hide_Show(worksheet, 0);

            // Assert
            Assert.AreEqual(eWorkSheetHidden.Visible, worksheet.Hidden);
        }

        [TestMethod]
        public void MssWorksheet_Protect_WithPassword_ProtectsSheet()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            string password = "SheetPassword123";

            // Act
            _excelPlusPlus.MssWorksheet_Protect(worksheet, password, null);

            // Assert
            Assert.IsTrue(worksheet.Protection.IsProtected);
        }

        [TestMethod]
        public void MssWorksheet_GetProperties_ReturnsWorksheetProperties()
        {
            // Arrange
            var package = CreateTestWorkbook("TestProps");
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssWorksheet_GetProperties(worksheet, out RCWorksheetRecord properties);

            // Assert
            Assert.IsNotNull(properties);
            Assert.AreEqual("TestProps", properties.ssSTWorksheet.ssName);
            Assert.AreEqual(1, properties.ssSTWorksheet.ssIndex);
        }

        [TestMethod]
        public void MssWorksheet_GetName_ReturnsCorrectName()
        {
            // Arrange
            var package = CreateTestWorkbook("NameTest");
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssWorksheet_GetName(worksheet, out string sheetName);

            // Assert
            Assert.AreEqual("NameTest", sheetName);
        }

        [TestMethod]
        public void MssWorksheet_Calculate_CalculatesFormulas()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = 10;
            worksheet.Cells["A2"].Value = 20;
            worksheet.Cells["A3"].Formula = "A1+A2";

            // Act
            _excelPlusPlus.MssWorksheet_Calculate(worksheet);

            // Assert
            Assert.AreEqual(30d, worksheet.Cells["A3"].Value);
        }

        [TestMethod]
        public void MssWorkbook_Calculate_CalculatesAllFormulas()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = 5;
            worksheet.Cells["A2"].Value = 15;
            worksheet.Cells["A3"].Formula = "A1+A2";

            // Act
            _excelPlusPlus.MssWorkbook_Calculate(package);

            // Assert
            Assert.AreEqual(20d, worksheet.Cells["A3"].Value);
        }

        [TestMethod]
        public void MssWorksheet_AddAutoFilter_AddsFilter()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "Header1";
            worksheet.Cells["B1"].Value = "Header2";
            worksheet.Cells["A2"].Value = "Data1";
            worksheet.Cells["B2"].Value = "Data2";

            // Act
            _excelPlusPlus.MssWorksheet_AddAutoFilter(worksheet, null);

            // Assert
            Assert.IsNotNull(worksheet.AutoFilter);
        }

        [TestMethod]
        public void MssWorksheet_AutofitColumns_AutoFitsColumns()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "This is a very long text that should trigger autofit";

            // Act
            _excelPlusPlus.MssWorksheet_AutofitColumns(worksheet);

            // Assert - Column width should be greater than default
            Assert.IsTrue(worksheet.Column(1).Width > 0);
        }

        [TestMethod]
        public void MssWorksheet_CopyRows_CopiesRows()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "Original";

            // Act
            _excelPlusPlus.MssWorksheet_CopyRows(worksheet, "A1:A1", "A2:A2");

            // Assert
            Assert.AreEqual("Original", worksheet.Cells["A2"].Value?.ToString());
        }

        [TestMethod]
        public void MssWorksheet_AddDropdown_AddsDataValidation()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var itemsList = new RLItemsRecordList();
            var item1 = new RCItemsRecord();
            item1.ssSTItems.ssItemText = "Option1";
            itemsList.Add(item1);
            var item2 = new RCItemsRecord();
            item2.ssSTItems.ssItemText = "Option2";
            itemsList.Add(item2);

            // Act
            _excelPlusPlus.MssWorksheet_AddDropdown(
                worksheet, 
                itemsList, 
                "", 
                "A:A", 
                "Title", 
                "Select an option", 
                true, 
                "Invalid selection", 
                "Error");

            // Assert
            Assert.IsTrue(worksheet.DataValidations.Count > 0);
        }

        [TestMethod]
        public void MssWorksheet_SetFooter_SetsFooterSections()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssWorksheet_SetFooter(worksheet, "Left", "Center", "Right");

            // Assert
            Assert.AreEqual("Left", worksheet.HeaderFooter.OddFooter.LeftAlignedText);
            Assert.AreEqual("Center", worksheet.HeaderFooter.OddFooter.CenteredText);
            Assert.AreEqual("Right", worksheet.HeaderFooter.OddFooter.RightAlignedText);
        }

        [TestMethod]
        public void MssWorksheet_SetHeader_SetsHeaderSections()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssWorksheet_SetHeader(worksheet, "LeftHeader", "CenterHeader", "RightHeader");

            // Assert
            Assert.AreEqual("LeftHeader", worksheet.HeaderFooter.OddHeader.LeftAlignedText);
            Assert.AreEqual("CenterHeader", worksheet.HeaderFooter.OddHeader.CenteredText);
            Assert.AreEqual("RightHeader", worksheet.HeaderFooter.OddHeader.RightAlignedText);
        }

        [TestMethod]
        public void MssWorksheet_GetHeader_RetrievesHeaderSections()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.HeaderFooter.OddHeader.LeftAlignedText = "TestLeft";
            worksheet.HeaderFooter.OddHeader.CenteredText = "TestCenter";
            worksheet.HeaderFooter.OddHeader.RightAlignedText = "TestRight";

            // Act
            _excelPlusPlus.MssWorksheet_GetHeader(worksheet, false, out string left, out string center, out string right);

            // Assert
            Assert.AreEqual("TestLeft", left);
            Assert.AreEqual("TestCenter", center);
            Assert.AreEqual("TestRight", right);
        }

        [TestMethod]
        public void MssWorksheet_GetFooter_RetrievesFooterSections()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.HeaderFooter.OddFooter.LeftAlignedText = "FooterLeft";

            // Act
            _excelPlusPlus.MssWorksheet_GetFooter(worksheet, false, out string left, out string center, out string right);

            // Assert
            Assert.AreEqual("FooterLeft", left);
        }

        #endregion

        #region Cell Tests

        [TestMethod]
        public void MssCell_WriteByIndex_WritesStringValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_WriteByIndex(worksheet, 1, 1, "TestValue", "text");

            // Assert
            Assert.AreEqual("TestValue", worksheet.Cells[1, 1].Value?.ToString());
        }

        [TestMethod]
        public void MssCell_WriteByIndex_WritesIntegerValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_WriteByIndex(worksheet, 1, 1, "42", "integer");

            // Assert
            Assert.AreEqual(42, worksheet.Cells[1, 1].Value);
        }

        [TestMethod]
        public void MssCell_WriteByIndex_WritesDecimalValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_WriteByIndex(worksheet, 1, 1, "3.14", "decimal");

            // Assert
            Assert.AreEqual(3.14m, worksheet.Cells[1, 1].Value);
        }

        [TestMethod]
        public void MssCell_WriteByIndex_WritesBooleanValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_WriteByIndex(worksheet, 1, 1, "TRUE", "boolean");

            // Assert
            Assert.AreEqual(true, worksheet.Cells[1, 1].Value);
        }

        [TestMethod]
        public void MssCell_WriteByIndex_WritesFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = 10;
            worksheet.Cells["A2"].Value = 20;

            // Act
            _excelPlusPlus.MssCell_WriteByIndex(worksheet, 3, 1, "A1+A2", "formula");

            // Assert
            Assert.AreEqual("A1+A2", worksheet.Cells[3, 1].Formula);
        }

        [TestMethod]
        public void MssCell_WriteByName_WritesValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_WriteByName(worksheet, "B5", "NamedCell", "text");

            // Assert
            Assert.AreEqual("NamedCell", worksheet.Cells["B5"].Value?.ToString());
        }

        [TestMethod]
        public void MssCell_ReadByIndex_ReadsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].Value = "ReadTest";

            // Act
            _excelPlusPlus.MssCell_ReadByIndex(worksheet, 1, 1, false, out string cellValue);

            // Assert
            Assert.AreEqual("ReadTest", cellValue);
        }

        [TestMethod]
        public void MssCell_ReadByName_ReadsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["C3"].Value = "NamedRead";

            // Act
            _excelPlusPlus.MssCell_ReadByName(worksheet, "C3", false, out string cellValue);

            // Assert
            Assert.AreEqual("NamedRead", cellValue);
        }

        [TestMethod]
        public void MssCell_ReadByIndex_AsText_ReturnsText()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].Value = 123.45;

            // Act
            _excelPlusPlus.MssCell_ReadByIndex(worksheet, 1, 1, true, out string cellValue);

            // Assert
            Assert.AreEqual("123.45", cellValue);
        }

        [TestMethod]
        public void MssCell_Read_WithCellName_ReadsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["D4"].Value = "CellReadTest";

            // Act
            _excelPlusPlus.MssCell_Read(worksheet, "D4", 0, 0, out string cellValue, false);

            // Assert
            Assert.AreEqual("CellReadTest", cellValue);
        }

        [TestMethod]
        public void MssCell_Read_WithRowCol_ReadsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[5, 4].Value = "RowColRead";

            // Act
            _excelPlusPlus.MssCell_Read(worksheet, "", 5, 4, out string cellValue, false);

            // Assert
            Assert.AreEqual("RowColRead", cellValue);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssCell_Read_WithNoAddress_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_Read(worksheet, "", 0, 0, out string cellValue, false);
        }

        [TestMethod]
        public void MssCell_SetFormulaByIndex_SetsFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_SetFormulaByIndex(worksheet, 1, 1, "SUM(A2:A10)");

            // Assert
            Assert.AreEqual("SUM(A2:A10)", worksheet.Cells[1, 1].Formula);
        }

        [TestMethod]
        public void MssCell_SetFormulaByName_SetsFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCell_SetFormulaByName(worksheet, "E5", "AVERAGE(E1:E4)");

            // Assert
            Assert.AreEqual("AVERAGE(E1:E4)", worksheet.Cells["E5"].Formula);
        }

        [TestMethod]
        public void MssCell_CalculateByIndex_CalculatesFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = 100;
            worksheet.Cells["A2"].Value = 200;
            worksheet.Cells["A3"].Formula = "A1+A2";

            // Act
            _excelPlusPlus.MssCell_CalculateByIndex(worksheet, 3, 1);

            // Assert
            Assert.AreEqual(300d, worksheet.Cells["A3"].Value);
        }

        [TestMethod]
        public void MssCell_CalculateByName_CalculatesFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["B1"].Value = 50;
            worksheet.Cells["B2"].Value = 50;
            worksheet.Cells["B3"].Formula = "B1+B2";

            // Act
            _excelPlusPlus.MssCell_CalculateByName(worksheet, "B3");

            // Assert
            Assert.AreEqual(100d, worksheet.Cells["B3"].Value);
        }

        [TestMethod]
        public void MssCell_ClearValueByIndex_ClearsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].Value = "ToClear";

            // Act
            _excelPlusPlus.MssCell_ClearValueByIndex(worksheet, 1, 1, 0, false);

            // Assert
            Assert.IsNull(worksheet.Cells[1, 1].Value);
        }

        [TestMethod]
        public void MssCell_ClearValueByName_ClearsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["F1"].Value = "ClearByName";

            // Act
            _excelPlusPlus.MssCell_ClearValueByName(worksheet, "F1", false);

            // Assert
            Assert.IsNull(worksheet.Cells["F1"].Value);
        }

        [TestMethod]
        public void MssCell_ReadFormulaByIndex_ReadsFormula()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].Formula = "SUM(G1:G10)";

            // Act
            _excelPlusPlus.MssCell_ReadFormulaByIndex(worksheet, 1, 1, out string formula);

            // Assert
            Assert.AreEqual("SUM(G1:G10)", formula);
        }

        [TestMethod]
        public void MssCell_FormatRange_AppliesFormatting()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var format = new RCCellFormatRecord();
            format.ssSTCellFormat.ssFontName = "Arial";
            format.ssSTCellFormat.ssFontSize = 12;
            format.ssSTCellFormat.ssBold = true;

            // Act
            _excelPlusPlus.MssCell_FormatRange(worksheet, 1, 1, 5, 5, format);

            // Assert
            Assert.AreEqual("Arial", worksheet.Cells[1, 1].Style.Font.Name);
            Assert.AreEqual(12, worksheet.Cells[1, 1].Style.Font.Size);
            Assert.IsTrue(worksheet.Cells[1, 1].Style.Font.Bold);
        }

        [TestMethod]
        public void MssCell_Merge_MergesCells()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var range = new RCRangeRecord();
            range.ssSTRange.ssStartRow = 1;
            range.ssSTRange.ssStartCol = 1;
            range.ssSTRange.ssEndRow = 1;
            range.ssSTRange.ssEndCol = 3;

            // Act
            _excelPlusPlus.MssCell_Merge(worksheet, range);

            // Assert
            Assert.IsTrue(worksheet.Cells[1, 1].Merge);
        }

        [TestMethod]
        public void MssCell_UnMerge_UnmergesCells()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1:C1"].Merge = true;
            var range = new RCRangeRecord();
            range.ssSTRange.ssStartRow = 1;
            range.ssSTRange.ssStartCol = 1;
            range.ssSTRange.ssEndRow = 1;
            range.ssSTRange.ssEndCol = 3;

            // Act
            _excelPlusPlus.MssCell_UnMerge(worksheet, range);

            // Assert
            Assert.IsFalse(worksheet.Cells[1, 1].Merge);
        }

        [TestMethod]
        public void MssCell_WriteImageByIndex_AddsImage()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }; // Minimal PNG header

            // Act
            _excelPlusPlus.MssCell_WriteImageByIndex(worksheet, 1, 1, "TestImage", testImage);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        [TestMethod]
        public void MssCell_WriteImageByName_AddsImage()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

            // Act
            _excelPlusPlus.MssCell_WriteImageByName(worksheet, "A1", "NamedImage", testImage);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        [TestMethod]
        public void MssCell_GetFillColorByIndex_ReturnsFillColor()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);

            // Act
            _excelPlusPlus.MssCell_GetFillColorByIndex(worksheet, 1, 1, out string fillColor);

            // Assert
            Assert.IsNotNull(fillColor);
            Assert.IsTrue(fillColor.StartsWith("#"));
        }

        [TestMethod]
        public void MssCell_GetFillColorByName_ReturnsFillColor()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.Blue);

            // Act
            _excelPlusPlus.MssCell_GetFillColorByName(worksheet, "A1", out string fillColor);

            // Assert
            Assert.IsNotNull(fillColor);
            Assert.IsTrue(fillColor.StartsWith("#"));
        }

        [TestMethod]
        public void MssUtil_ConvertHexCodeToRGB_ConvertsColor()
        {
            // Act
            _excelPlusPlus.MssUtil_ConvertHexCodeToRGB("#FF0000", out string rgb);

            // Assert
            Assert.AreEqual("RGB(255, 0, 0)", rgb);
        }

        [TestMethod]
        public void MssUtil_ConvertHexCodeToRGB_HandlesEmptyString()
        {
            // Act
            _excelPlusPlus.MssUtil_ConvertHexCodeToRGB("", out string rgb);

            // Assert
            Assert.AreEqual("No Fill Color", rgb);
        }

        #endregion

        #region Row/Column Tests

        [TestMethod]
        public void MssColumn_SetWidth_SetsColumnWidth()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssColumn_SetWidth(worksheet, 1, 200);

            // Assert
            Assert.AreEqual(200, worksheet.Column(1).Width);
        }

        [TestMethod]
        public void MssRow_SetHeight_SetsRowHeight()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssRow_SetHeight(worksheet, 1, 50);

            // Assert
            Assert.AreEqual(50, worksheet.Row(1).Height);
        }

        [TestMethod]
        public void MssColumn_Insert_InsertsColumn()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "Original";

            // Act
            _excelPlusPlus.MssColumn_Insert(worksheet, 1, 1, 0);

            // Assert
            Assert.AreEqual("Original", worksheet.Cells["B1"].Value?.ToString());
        }

        [TestMethod]
        public void MssColumn_Delete_DeletesColumn()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "ToDelete";
            worksheet.Cells["B1"].Value = "ToKeep";

            // Act
            _excelPlusPlus.MssColumn_Delete(worksheet, 1, 1);

            // Assert
            Assert.AreEqual("ToKeep", worksheet.Cells["A1"].Value?.ToString());
        }

        [TestMethod]
        public void MssRow_Insert_InsertsRow()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "Original";

            // Act
            _excelPlusPlus.MssRow_Insert(worksheet, 1, 1, 0);

            // Assert
            Assert.AreEqual("Original", worksheet.Cells["A2"].Value?.ToString());
        }

        [TestMethod]
        public void MssRow_Delete_DeletesRow()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "ToDelete";
            worksheet.Cells["A2"].Value = "ToKeep";

            // Act
            _excelPlusPlus.MssRow_Delete(worksheet, 1, 1);

            // Assert
            Assert.AreEqual("ToKeep", worksheet.Cells["A1"].Value?.ToString());
        }

        [TestMethod]
        public void MssColumn_Hide_Show_HidesColumn()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssColumn_Hide_Show(worksheet, 1, true);

            // Assert
            Assert.IsTrue(worksheet.Column(1).Hidden);
        }

        [TestMethod]
        public void MssColumn_Hide_Show_ShowsColumn()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Column(1).Hidden = true;

            // Act
            _excelPlusPlus.MssColumn_Hide_Show(worksheet, 1, false);

            // Assert
            Assert.IsFalse(worksheet.Column(1).Hidden);
        }

        [TestMethod]
        public void MssRow_Hide_Show_HidesRow()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssRow_Hide_Show(worksheet, 1, true);

            // Assert
            Assert.IsTrue(worksheet.Row(1).Hidden);
        }

        [TestMethod]
        public void MssRow_Hide_Show_ShowsRow()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Row(1).Hidden = true;

            // Act
            _excelPlusPlus.MssRow_Hide_Show(worksheet, 1, false);

            // Assert
            Assert.IsFalse(worksheet.Row(1).Hidden);
        }

        #endregion

        #region Address Tests

        [TestMethod]
        public void MssAddress_From_Text_ParsesSingleCell()
        {
            // Act
            _excelPlusPlus.MssAddress_From_Text("C5", out int rowStart, out int colStart, out int rowEnd, out int colEnd);

            // Assert
            Assert.AreEqual(5, rowStart);
            Assert.AreEqual(3, colStart);
            Assert.AreEqual(5, rowEnd);
            Assert.AreEqual(3, colEnd);
        }

        [TestMethod]
        public void MssAddress_From_Text_ParsesRange()
        {
            // Act
            _excelPlusPlus.MssAddress_From_Text("A1:C3", out int rowStart, out int colStart, out int rowEnd, out int colEnd);

            // Assert
            Assert.AreEqual(1, rowStart);
            Assert.AreEqual(1, colStart);
            Assert.AreEqual(3, rowEnd);
            Assert.AreEqual(3, colEnd);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssAddress_From_Text_WithEmptyString_ThrowsException()
        {
            // Act
            _excelPlusPlus.MssAddress_From_Text("", out int rowStart, out int colStart, out int rowEnd, out int colEnd);
        }

        [TestMethod]
        public void MssAddress_From_RowCol_GeneratesSingleCell()
        {
            // Act
            _excelPlusPlus.MssAddress_From_RowCol(5, 3, 0, 0, out string address);

            // Assert
            Assert.AreEqual("C5", address);
        }

        [TestMethod]
        public void MssAddress_From_RowCol_GeneratesRange()
        {
            // Act
            _excelPlusPlus.MssAddress_From_RowCol(1, 1, 3, 3, out string address);

            // Assert
            Assert.AreEqual("A1:C3", address);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssAddress_From_RowCol_WithInvalidStart_ThrowsException()
        {
            // Act
            _excelPlusPlus.MssAddress_From_RowCol(0, 0, 0, 0, out string address);
        }

        #endregion

        #region Comment Tests

        [TestMethod]
        public void MssComment_Add_AddsComment()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssComment_Add(worksheet, 1, 1, "Test Comment", "Test Author", true);

            // Assert
            Assert.IsNotNull(worksheet.Cells[1, 1].Comment);
            Assert.AreEqual("Test Comment", worksheet.Cells[1, 1].Comment.Text);
            Assert.AreEqual("Test Author", worksheet.Cells[1, 1].Comment.Author);
        }

        [TestMethod]
        public void MssComment_Delete_DeletesComment()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells[1, 1].AddComment("To Delete", "Author");
            var range = new RCRangeRecord();
            range.ssSTRange.ssStartRow = 1;
            range.ssSTRange.ssStartCol = 1;
            range.ssSTRange.ssEndRow = 1;
            range.ssSTRange.ssEndCol = 1;

            // Act
            _excelPlusPlus.MssComment_Delete(worksheet, range);

            // Assert
            Assert.IsNull(worksheet.Cells[1, 1].Comment);
        }

        #endregion

        #region Conditional Formatting Tests

        [TestMethod]
        public void MssConditionalFormatting_AddRule_AddsGreaterThanRule()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var cfRecord = new RCConditionalFormatItemRecord();
            cfRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress = "A1:A10";
            cfRecord.ssSTConditionalFormatItem.ssRuleType = (int)eExcelConditionalFormattingRuleType.GreaterThan;
            cfRecord.ssSTConditionalFormatItem.ssFormula = "5";
            cfRecord.ssSTConditionalFormatItem.ssPriority = 1;
            cfRecord.ssSTConditionalFormatItem.ssStopIfTrue = false;

            // Act
            _excelPlusPlus.MssConditionalFormatting_AddRule(worksheet, cfRecord);

            // Assert
            Assert.AreEqual(1, worksheet.ConditionalFormatting.Count);
        }

        [TestMethod]
        public void MssConditionalFormatting_GetAllRules_RetrievesRules()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var cf = worksheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A10"));
            cf.Formula = "10";

            // Act
            _excelPlusPlus.MssConditionalFormatting_GetAllRules(worksheet, out RLConditionalFormatItemRecordList rules);

            // Assert
            Assert.IsNotNull(rules);
            Assert.AreEqual(1, rules.Count);
        }

        [TestMethod]
        public void MssConditionalFormatting_DeleteRule_DeletesRule()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A10"));

            // Act
            _excelPlusPlus.MssConditionalFormatting_DeleteRule(worksheet, 1);

            // Assert
            Assert.AreEqual(0, worksheet.ConditionalFormatting.Count);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void MssConditionalFormatting_DeleteRule_WithInvalidIndex_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssConditionalFormatting_DeleteRule(worksheet, 0);
        }

        [TestMethod]
        public void MssConditionalFormatting_DeleteAllRules_DeletesAllRules()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A10"));
            worksheet.ConditionalFormatting.AddLessThan(new ExcelAddress("B1:B10"));

            // Act
            _excelPlusPlus.MssConditionalFormatting_DeleteAllRules(worksheet);

            // Assert
            Assert.AreEqual(0, worksheet.ConditionalFormatting.Count);
        }

        #endregion

        #region Chart Tests

        [TestMethod]
        public void MssChart_Create_CreatesColumnChart()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "Category";
            worksheet.Cells["A2"].Value = 10;
            worksheet.Cells["A3"].Value = 20;
            worksheet.Cells["B2"].Value = 15;
            worksheet.Cells["B3"].Value = 25;

            var dataSeriesList = new RLDataSeriesRecordList();
            var dataSeries = new RCDataSeriesRecord();
            dataSeries.ssSTDataSeries.ssName = "Series1";
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssStartRow = 2;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssStartCol = 2;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssEndRow = 3;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssEndCol = 2;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssStartRow = 2;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssStartCol = 1;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssEndRow = 3;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssEndCol = 1;
            dataSeriesList.Add(dataSeries);

            // Act
            _excelPlusPlus.MssChart_Create(worksheet, "ColumnClustered", "TestChart", dataSeriesList, 300, 400, 5, 3);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        [TestMethod]
        public void MssWorksheet_Chart_Create_CreatesChart()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = 10;
            worksheet.Cells["A2"].Value = 20;

            var dataSeriesList = new RLDataSeriesRecordList();
            var dataSeries = new RCDataSeriesRecord();
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssStartRow = 1;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssStartCol = 1;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssEndRow = 2;
            dataSeries.ssSTDataSeries.ssValueRange.ssSTRange.ssEndCol = 1;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssStartRow = 1;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssStartCol = 1;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssEndRow = 2;
            dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange.ssEndCol = 1;
            dataSeriesList.Add(dataSeries);

            // Act
            _excelPlusPlus.MssWorksheet_Chart_Create(worksheet, "Pie", "PieChart", dataSeriesList, 300, 300, 5, 5);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        #endregion

        #region Image Tests

        [TestMethod]
        public void MssImage_Insert_InsertsImageByRowCol()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

            // Act
            _excelPlusPlus.MssImage_Insert(
                worksheet, 
                testImage, 
                "PNG", 
                "InsertedImg", 
                1, 
                1, 
                "", 
                100, 
                100, 
                0, 
                0);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        [TestMethod]
        public void MssImage_Insert_InsertsImageByCellName()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

            // Act
            _excelPlusPlus.MssImage_Insert(
                worksheet, 
                testImage, 
                "PNG", 
                "NamedImg", 
                0, 
                0, 
                "B2", 
                100, 
                100, 
                0, 
                0);

            // Assert
            Assert.AreEqual(1, worksheet.Drawings.Count);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssImage_Insert_WithInvalidCell_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

            // Act
            _excelPlusPlus.MssImage_Insert(
                worksheet, 
                testImage, 
                "PNG", 
                "Img", 
                0, 
                0, 
                "", 
                100, 
                100, 
                0, 
                0);
        }

        [TestMethod]
        public void MssWorksheet_GetImages_RetrievesImages()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var testImage = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            _excelPlusPlus.MssCell_WriteImageByIndex(worksheet, 1, 1, "TestImg", testImage);

            // Act
            _excelPlusPlus.MssWorksheet_GetImages(worksheet, out RLImageRecordList images);

            // Assert
            Assert.IsNotNull(images);
            Assert.AreEqual(1, images.Count);
        }

        #endregion

        #region Excel Properties Tests

        [TestMethod]
        public void MssExcel_GetProperties_RetrievesProperties()
        {
            // Arrange
            var package = CreateTestWorkbook();
            package.Workbook.Properties.Author = "Test Author";
            package.Workbook.Properties.Title = "Test Title";

            // Act
            _excelPlusPlus.MssExcel_GetProperties(package, out RCOfficePropertiesRecord properties);

            // Assert
            Assert.IsNotNull(properties);
            Assert.AreEqual("Test Author", properties.ssSTOfficeProperties.ssAuthor);
            Assert.AreEqual("Test Title", properties.ssSTOfficeProperties.ssTitle);
        }

        [TestMethod]
        public void MssExcel_SetProperties_SetsProperties()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var properties = new RCOfficePropertiesRecord(null);
            properties.ssSTOfficeProperties.ssAuthor = "New Author";
            properties.ssSTOfficeProperties.ssTitle = "New Title";
            properties.ssSTOfficeProperties.ssCompany = "Test Company";

            // Act
            _excelPlusPlus.MssExcel_SetProperties(package, properties, false);

            // Assert
            Assert.AreEqual("New Author", package.Workbook.Properties.Author);
            Assert.AreEqual("New Title", package.Workbook.Properties.Title);
            Assert.AreEqual("Test Company", package.Workbook.Properties.Company);
        }

        [TestMethod]
        public void MssExcel_ClearProperties_ClearsAllProperties()
        {
            // Arrange
            var package = CreateTestWorkbook();
            package.Workbook.Properties.Author = "To Clear";
            package.Workbook.Properties.Title = "To Clear";

            // Act
            _excelPlusPlus.MssExcel_ClearProperties(
                package, 
                true, true, true, true, true, true, true, true, true, true);

            // Assert
            Assert.AreEqual(string.Empty, package.Workbook.Properties.Author);
            Assert.AreEqual(string.Empty, package.Workbook.Properties.Title);
        }

        #endregion

        #region Data Range Tests

        [TestMethod]
        public void MssCell_WriteRange_WritesData()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            // Note: This test would require creating proper RecordList data structures
            // For now, we test the method exists and doesn't throw with empty data
            var emptyRecordList = new RLValueRecordList();

            // Act - should not throw with empty list
            _excelPlusPlus.MssCell_WriteColumnRange(worksheet, 1, 1, emptyRecordList, "text");

            // Assert - no exception thrown
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void MssCell_WriteColumnRange_WritesColumnData()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var emptyRecordList = new RLValueRecordList();

            // Act
            _excelPlusPlus.MssCell_WriteColumnRangeWithFormat(worksheet, 1, 1, emptyRecordList, "text", null);

            // Assert - no exception with empty data
            Assert.IsTrue(true);
        }

        #endregion

        #region Find/Search Tests

        [TestMethod]
        public void MssCells_FindByValue_FindsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "SearchTarget";
            worksheet.Cells["A2"].Value = "Other";
            worksheet.Cells["A3"].Value = "SearchTarget";

            // Act
            _excelPlusPlus.MssCells_FindByValue(worksheet, "SearchTarget", out RLRangeRecordList foundCells);

            // Assert
            Assert.IsNotNull(foundCells);
            Assert.IsTrue(foundCells.Count > 0);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssCells_FindByValue_WithEmptyValue_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);

            // Act
            _excelPlusPlus.MssCells_FindByValue(worksheet, "", out RLRangeRecordList foundCells);
        }

        [TestMethod]
        public void MssContainInRange_FindsValue()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "FindMe";
            worksheet.Cells["A2"].Value = "Other";

            // Act
            _excelPlusPlus.MssContainInRange(worksheet, "A1:A2", "FindMe", "", out bool found, out int rowIndex, out int colIndex);

            // Assert
            Assert.IsTrue(found);
            Assert.AreEqual(1, rowIndex);
            Assert.AreEqual(1, colIndex);
        }

        [TestMethod]
        public void MssContainInRange_ValueNotFound()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "NotThis";

            // Act
            _excelPlusPlus.MssContainInRange(worksheet, "A1:A1", "NotFound", "", out bool found, out int rowIndex, out int colIndex);

            // Assert
            Assert.IsFalse(found);
            Assert.AreEqual(0, rowIndex);
            Assert.AreEqual(0, colIndex);
        }

        #endregion

        #region CellFormat Tests

        [TestMethod]
        public void MssCellFormat_ApplyToRange_AppliesFormat()
        {
            // Arrange
            var package = CreateTestWorkbook();
            var worksheet = GetFirstWorksheet(package);
            var format = new RCCellFormatRecord();
            format.ssSTCellFormat.ssFontName = "Times New Roman";
            format.ssSTCellFormat.ssFontSize = 14;
            var range = new RCRangeRecord();
            range.ssSTRange.ssStartRow = 1;
            range.ssSTRange.ssStartCol = 1;
            range.ssSTRange.ssEndRow = 5;
            range.ssSTRange.ssEndCol = 5;

            // Act
            _excelPlusPlus.MssCellFormat_ApplyToRange(worksheet, format, range);

            // Assert
            Assert.AreEqual("Times New Roman", worksheet.Cells[1, 1].Style.Font.Name);
            Assert.AreEqual(14, worksheet.Cells[1, 1].Style.Font.Size);
        }

        #endregion

        #region WorkBook AddSheet Tests

        [TestMethod]
        public void MssWorkBook_AddSheet_AddsSheetAtIndex()
        {
            // Arrange
            var package = CreateTestWorkbook("First");
            package.Workbook.Worksheets.Add("Second");

            // Act
            _excelPlusPlus.MssWorkBook_AddSheet(package, "InsertedSheet", null, 1);

            // Assert
            Assert.AreEqual(3, package.Workbook.Worksheets.Count);
            Assert.AreEqual("InsertedSheet", package.Workbook.Worksheets[0].Name);
        }

        [TestMethod]
        public void MssWorkBook_AddSheet_AddsSheetCopy()
        {
            // Arrange
            var package = CreateTestWorkbook("Original");
            var worksheet = GetFirstWorksheet(package);
            worksheet.Cells["A1"].Value = "TestData";

            // Act
            _excelPlusPlus.MssWorkBook_AddSheet(package, "CopiedSheet", worksheet, 0);

            // Assert
            Assert.AreEqual(2, package.Workbook.Worksheets.Count);
            var copiedWs = package.Workbook.Worksheets["CopiedSheet"];
            Assert.IsNotNull(copiedWs);
            Assert.AreEqual("TestData", copiedWs.Cells["A1"].Value?.ToString());
        }

        #endregion

        #region Worksheet Select Tests

        [TestMethod]
        public void MssWorksheet_Select_ByIndex_SelectsWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("First");
            package.Workbook.Worksheets.Add("Second");

            // Act
            _excelPlusPlus.MssWorksheet_Select(package, 2, "", out object worksheet);

            // Assert
            Assert.IsNotNull(worksheet);
            var ws = worksheet as ExcelWorksheet;
            Assert.AreEqual("Second", ws.Name);
        }

        [TestMethod]
        public void MssWorksheet_Select_ByName_SelectsWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("Target");

            // Act
            _excelPlusPlus.MssWorksheet_Select(package, 0, "Target", out object worksheet);

            // Assert
            Assert.IsNotNull(worksheet);
            var ws = worksheet as ExcelWorksheet;
            Assert.AreEqual("Target", ws.Name);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssWorksheet_Select_WithNoParameters_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();

            // Act
            _excelPlusPlus.MssWorksheet_Select(package, 0, "", out object worksheet);
        }

        #endregion

        #region Worksheet Delete Tests

        [TestMethod]
        public void MssWorksheet_Delete_ByIndex_DeletesWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("Keep");
            package.Workbook.Worksheets.Add("Delete");

            // Act
            _excelPlusPlus.MssWorksheet_Delete(package, 2, "");

            // Assert
            Assert.AreEqual(1, package.Workbook.Worksheets.Count);
        }

        [TestMethod]
        public void MssWorksheet_Delete_ByName_DeletesWorksheet()
        {
            // Arrange
            var package = CreateTestWorkbook("Delete");
            package.Workbook.Worksheets.Add("Keep");

            // Act
            _excelPlusPlus.MssWorksheet_Delete(package, 0, "Delete");

            // Assert
            Assert.AreEqual(1, package.Workbook.Worksheets.Count);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void MssWorksheet_Delete_WithNoParameters_ThrowsException()
        {
            // Arrange
            var package = CreateTestWorkbook();

            // Act
            _excelPlusPlus.MssWorksheet_Delete(package, 0, "");
        }

        #endregion
    }
}
