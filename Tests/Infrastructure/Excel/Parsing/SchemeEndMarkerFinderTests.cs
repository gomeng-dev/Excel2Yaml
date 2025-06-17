using System;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Tests.Infrastructure.Excel.Parsing
{
    public class SchemeEndMarkerFinderTests
    {
        private readonly ISimpleLogger _mockLogger = new MockLogger();
        private readonly SchemeEndMarkerFinder _finder;

        public SchemeEndMarkerFinderTests()
        {
            _finder = new SchemeEndMarkerFinder(_mockLogger);
        }

        public void FindSchemeEndRow_ShouldFindMarkerInValidRow()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(5, 1).Value = SchemeConstants.Markers.SchemeEnd;

                // Act
                var result = _finder.FindSchemeEndRow(worksheet);

                // Assert
                TestAssert.AreEqual(5, result);
            }
        }

        public void FindSchemeEndRow_ShouldReturnIllegalRowWhenNotFound()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(1, 1).Value = "Some data";

                // Act
                var result = _finder.FindSchemeEndRow(worksheet);

                // Assert
                TestAssert.AreEqual(SchemeConstants.RowNumbers.IllegalRow, result);
            }
        }

        public void FindSchemeEndRow_ShouldSkipCommentRow()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(SchemeConstants.RowNumbers.CommentRow + 1, 1).Value = SchemeConstants.Markers.SchemeEnd;
                worksheet.Cell(6, 1).Value = SchemeConstants.Markers.SchemeEnd;

                // Act
                var result = _finder.FindSchemeEndRow(worksheet);

                // Assert
                TestAssert.AreEqual(6, result); // Should skip comment row and find the next one
            }
        }

        public void ContainsEndMarker_ShouldReturnTrueForValidMarker()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(1, 1).Value = SchemeConstants.Markers.SchemeEnd;
                var row = worksheet.Row(1);

                // Act
                var result = _finder.ContainsEndMarker(row);

                // Assert
                TestAssert.IsTrue(result);
            }
        }

        public void ContainsEndMarker_ShouldReturnFalseForEmptyCell()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var row = worksheet.Row(1);

                // Act
                var result = _finder.ContainsEndMarker(row);

                // Assert
                TestAssert.IsFalse(result);
            }
        }

        public void ContainsEndMarker_ShouldBeCaseInsensitive()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                worksheet.Cell(1, 1).Value = "$SCHEME_END"; // Upper case
                var row = worksheet.Row(1);

                // Act
                var result = _finder.ContainsEndMarker(row);

                // Assert
                TestAssert.IsTrue(result);
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