using System;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using ExcelToYamlAddin.Tests.Common;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Tests.Infrastructure.Excel.Parsing
{
    public class ParsingContextTests
    {
        public void Constructor_ShouldInitializeAllPropertiesCorrectly()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                // Act
                var context = new ParsingContext(
                    worksheet,
                    schemeStartRow,
                    schemeEndRowNumber: 10,
                    dataStartRowNumber: 11,
                    firstCellNumber: 1,
                    lastCellNumber: 5
                );

                // Assert
                TestAssert.AreEqual(worksheet, context.Worksheet);
                TestAssert.AreEqual(schemeStartRow, context.SchemeStartRow);
                TestAssert.AreEqual(10, context.SchemeEndRowNumber);
                TestAssert.AreEqual(11, context.DataStartRowNumber);
                TestAssert.AreEqual(1, context.FirstCellNumber);
                TestAssert.AreEqual(5, context.LastCellNumber);
            }
        }

        public void Constructor_ShouldThrowExceptionForNullWorksheet()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                // Act & Assert
                TestAssert.Throws<ArgumentNullException>(() =>
                    new ParsingContext(null, schemeStartRow, 10, 11, 1, 5)
                );
            }
        }

        public void Constructor_ShouldThrowExceptionForNullSchemeStartRow()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");

                // Act & Assert
                TestAssert.Throws<ArgumentNullException>(() =>
                    new ParsingContext(worksheet, null, 10, 11, 1, 5)
                );
            }
        }

        public void Constructor_ShouldThrowExceptionForInvalidSchemeEndRowNumber()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                // Act & Assert
                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, 0, 11, 1, 5)
                );

                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, -1, 11, 1, 5)
                );
            }
        }

        public void Constructor_ShouldThrowExceptionWhenDataStartRowIsBeforeSchemeEnd()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                // Act & Assert
                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, 10, 10, 1, 5)
                );

                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, 10, 9, 1, 5)
                );
            }
        }

        public void Constructor_ShouldThrowExceptionForInvalidCellNumbers()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                // Act & Assert
                // Invalid first cell number
                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, 10, 11, 0, 5)
                );

                // Last cell number less than first
                TestAssert.Throws<ArgumentException>(() =>
                    new ParsingContext(worksheet, schemeStartRow, 10, 11, 5, 4)
                );
            }
        }

        public void GetDataEndRowNumber_ShouldReturnLastUsedRowNumber()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);
                
                // Add some data
                worksheet.Cell(15, 1).Value = "Data";

                var context = new ParsingContext(worksheet, schemeStartRow, 10, 11, 1, 5);

                // Act
                var result = context.GetDataEndRowNumber();

                // Assert
                TestAssert.AreEqual(15, result);
            }
        }

        public void GetDataEndRowNumber_ShouldReturnDataStartRowWhenNoData()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var schemeStartRow = worksheet.Row(2);

                var context = new ParsingContext(worksheet, schemeStartRow, 10, 11, 1, 5);

                // Act
                var result = context.GetDataEndRowNumber();

                // Assert
                TestAssert.AreEqual(11, result); // Should return data start row
            }
        }

        public void ToString_ShouldReturnFormattedSummary()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("TestSheet");
                var schemeStartRow = worksheet.Row(2);

                var context = new ParsingContext(worksheet, schemeStartRow, 10, 11, 1, 5);

                // Act
                var result = context.ToString();

                // Assert
                TestAssert.IsTrue(result.Contains("TestSheet"));
                TestAssert.IsTrue(result.Contains("2-10")); // Scheme rows
                TestAssert.IsTrue(result.Contains("11")); // Data start row
                TestAssert.IsTrue(result.Contains("1-5")); // Columns
            }
        }
    }
}