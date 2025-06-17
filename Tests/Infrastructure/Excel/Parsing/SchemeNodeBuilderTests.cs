using System;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Tests.Common;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Tests.Infrastructure.Excel.Parsing
{
    public class SchemeNodeBuilderTests
    {
        private readonly ISimpleLogger _mockLogger = new MockLogger();
        private readonly SchemeNodeBuilder _builder;

        public SchemeNodeBuilderTests()
        {
            _builder = new SchemeNodeBuilder(_mockLogger);
        }

        public void BuildFromCell_ShouldReturnNullForEmptyCell()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var cell = worksheet.Cell(1, 1);

                // Act
                var result = _builder.BuildFromCell(cell);

                // Assert
                TestAssert.IsNull(result);
            }
        }

        public void BuildFromCell_ShouldReturnNullForIgnoreMarker()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var cell = worksheet.Cell(1, 1);
                cell.Value = SchemeConstants.Markers.Ignore;

                // Act
                var result = _builder.BuildFromCell(cell);

                // Assert
                TestAssert.IsNull(result);
            }
        }

        public void BuildFromCell_ShouldCreatePropertyNode()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var cell = worksheet.Cell(2, 3);
                cell.Value = "propertyName";

                // Act
                var result = _builder.BuildFromCell(cell);

                // Assert
                TestAssert.IsNotNull(result);
                TestAssert.AreEqual("propertyName", result.Key);
                TestAssert.AreEqual(SchemeNodeType.Property, result.NodeType);
                TestAssert.AreEqual(2, result.Position.Row);
                TestAssert.AreEqual(3, result.Position.Column);
            }
        }

        public void BuildFromCell_ShouldCreateMapNode()
        {
            // Arrange
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Test");
                var cell = worksheet.Cell(1, 1);
                cell.Value = SchemeConstants.NodeTypes.Map;

                // Act
                var result = _builder.BuildFromCell(cell);

                // Assert
                TestAssert.IsNotNull(result);
                TestAssert.AreEqual(SchemeConstants.NodeTypes.Map, result.Key);
                TestAssert.AreEqual(SchemeNodeType.Map, result.NodeType);
            }
        }

        public void Build_ShouldCreateArrayNode()
        {
            // Act
            var result = _builder.Build(SchemeConstants.NodeTypes.Array, 5, 10);

            // Assert
            TestAssert.IsNotNull(result);
            TestAssert.AreEqual(SchemeConstants.NodeTypes.Array, result.Key);
            TestAssert.AreEqual(SchemeNodeType.Array, result.NodeType);
            TestAssert.AreEqual(5, result.Position.Row);
            TestAssert.AreEqual(10, result.Position.Column);
        }

        public void Build_ShouldCreateKeyNode()
        {
            // Act
            var result = _builder.Build(SchemeConstants.NodeTypes.Key, 3, 4);

            // Assert
            TestAssert.IsNotNull(result);
            TestAssert.AreEqual(SchemeConstants.NodeTypes.Key, result.Key);
            TestAssert.AreEqual(SchemeNodeType.Key, result.NodeType);
        }

        public void Build_ShouldCreateValueNode()
        {
            // Act
            var result = _builder.Build(SchemeConstants.NodeTypes.Value, 7, 8);

            // Assert
            TestAssert.IsNotNull(result);
            TestAssert.AreEqual(SchemeConstants.NodeTypes.Value, result.Key);
            TestAssert.AreEqual(SchemeNodeType.Value, result.NodeType);
        }

        public void Build_ShouldThrowExceptionForEmptyValue()
        {
            // Act & Assert
            TestAssert.Throws<ArgumentException>(() => _builder.Build("", 1, 1));
            TestAssert.Throws<ArgumentException>(() => _builder.Build(null, 1, 1));
        }

        public void ShouldIgnore_ShouldReturnTrueForIgnoreMarker()
        {
            // Act & Assert
            TestAssert.IsTrue(_builder.ShouldIgnore(SchemeConstants.Markers.Ignore));
            TestAssert.IsTrue(_builder.ShouldIgnore("^")); // Case insensitive
            TestAssert.IsTrue(_builder.ShouldIgnore(""));
            TestAssert.IsTrue(_builder.ShouldIgnore("   "));
            TestAssert.IsTrue(_builder.ShouldIgnore(null));
        }

        public void ShouldIgnore_ShouldReturnFalseForValidValues()
        {
            // Act & Assert
            TestAssert.IsFalse(_builder.ShouldIgnore("property"));
            TestAssert.IsFalse(_builder.ShouldIgnore(SchemeConstants.NodeTypes.Map));
            TestAssert.IsFalse(_builder.ShouldIgnore(SchemeConstants.NodeTypes.Array));
            TestAssert.IsFalse(_builder.ShouldIgnore("someValue"));
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