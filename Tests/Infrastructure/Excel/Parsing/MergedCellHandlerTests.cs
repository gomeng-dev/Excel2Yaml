using System;
using System.Linq;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Tests.Common;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Tests.Infrastructure.Excel.Parsing
{
    public class MergedCellHandlerTests
    {
        private readonly ISimpleLogger _mockLogger = new MockLogger();
        private readonly MergedCellHandler _handler;

        public MergedCellHandlerTests()
        {
            _handler = new MergedCellHandler(_mockLogger);
        }

        public void GetMergedRegionsInRow_ShouldReturnEmptyListWhenNoMergedCells()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(1, 1).Value = "A";
                worksheet.Cell(1, 2).Value = "B";

                // Act
                var result = _handler.GetMergedRegionsInRow(worksheet, 1);

                // Assert
                TestAssert.AreEqual(0, result.Count);
            }
        }

        public void GetMergedRegionsInRow_ShouldFindMergedRegionsInSpecifiedRow()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Range("A1:C1").Merge();
                worksheet.Range("D2:F2").Merge();

                // Act
                var result1 = _handler.GetMergedRegionsInRow(worksheet, 1);
                var result2 = _handler.GetMergedRegionsInRow(worksheet, 2);

                // Assert
                TestAssert.AreEqual(1, result1.Count);
                TestAssert.AreEqual(1, result2.Count);
            }
        }

        public void GetMergedRegionsInRow_ShouldFindMultipleMergedRegionsInSameRow()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Range("A1:B1").Merge();
                worksheet.Range("D1:F1").Merge();

                // Act
                var result = _handler.GetMergedRegionsInRow(worksheet, 1);

                // Assert
                TestAssert.AreEqual(2, result.Count);
            }
        }

        public void GetMergedCellRange_ShouldReturnSingleCellRangeWhenNotMerged()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var cell = worksheet.Cell(1, 2);
                var mergedRegions = _handler.GetMergedRegionsInRow(worksheet, 1);

                // Act
                var (startColumn, endColumn) = _handler.GetMergedCellRange(cell, mergedRegions);

                // Assert
                TestAssert.AreEqual(2, startColumn);
                TestAssert.AreEqual(2, endColumn);
            }
        }

        public void GetMergedCellRange_ShouldReturnMergedRangeWhenCellIsMerged()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Range("B1:E1").Merge();
                var cell = worksheet.Cell(1, 3); // Cell C1 is within merged range
                var mergedRegions = _handler.GetMergedRegionsInRow(worksheet, 1);

                // Act
                var (startColumn, endColumn) = _handler.GetMergedCellRange(cell, mergedRegions);

                // Assert
                TestAssert.AreEqual(2, startColumn); // Column B
                TestAssert.AreEqual(5, endColumn);   // Column E
            }
        }

        public void GetMergedCellRange_ShouldHandleVerticallyMergedCells()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Range("B1:B3").Merge();
                var cell = worksheet.Cell(2, 2); // Cell B2 is within vertically merged range
                var mergedRegions = _handler.GetMergedRegionsInRow(worksheet, 2);

                // Act
                var (startColumn, endColumn) = _handler.GetMergedCellRange(cell, mergedRegions);

                // Assert
                TestAssert.AreEqual(2, startColumn); // Column B
                TestAssert.AreEqual(2, endColumn);   // Column B (same column for vertical merge)
            }
        }

        // Mock logger for testing
        private class MockLogger : ISimpleLogger
        {
            public void Debug(string message) { }
            public void Debug(string messageTemplate, params object[] args) { }
            public void Information(string message) { }
            public void Information(string messageTemplate, params object[] args) { }
            public void Warning(string message) { }
            public void Warning(string messageTemplate, params object[] args) { }
            public void Error(string message) { }
            public void Error(string messageTemplate, params object[] args) { }
            public void Error(Exception exception, string message) { }
            public void Error(Exception exception, string messageTemplate, params object[] args) { }
        }
    }
}