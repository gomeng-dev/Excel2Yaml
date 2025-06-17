using System;
using System.Linq;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;

namespace ExcelToYamlAddin.Tests.Domain.Entities
{
    /// <summary>
    /// Scheme 엔티티 단위 테스트
    /// </summary>
    public class SchemeTests
    {
        public static void RunAllTests()
        {
            Console.WriteLine("=== Scheme Tests ===");
            
            Test_Create_ValidInput_CreatesScheme();
            Test_Create_InvalidSheetName_ThrowsException();
            Test_Create_NullRoot_ThrowsException();
            Test_Create_InvalidContentStartRow_ThrowsException();
            Test_Create_EndRowLessThanStartRow_ThrowsException();
            Test_Empty_CreatesEmptyScheme();
            Test_GetLinearNodes_ReturnsAllNodes();
            Test_GetNodesByType_FiltersCorrectly();
            Test_GetNodesByDepth_FiltersCorrectly();
            Test_FindNodeByPath_ReturnsCorrectNode();
            Test_Validate_ValidScheme_ReturnsSuccess();
            Test_Clone_CreatesDeepCopy();
            Test_DataRowCount_CalculatesCorrectly();
            
            Console.WriteLine("Scheme 테스트 완료!\n");
        }

        private static void Test_Create_ValidInput_CreatesScheme()
        {
            // Arrange
            var sheetName = "Sheet1";
            var root = SchemeNode.CreateRoot();
            var contentStartRow = 3;
            var endRow = 10;
            
            // Act
            var scheme = Scheme.Create(sheetName, root, contentStartRow, endRow);
            
            // Assert
            TestAssert.AreEqual(sheetName, scheme.SheetName);
            TestAssert.AreEqual(root, scheme.Root);
            TestAssert.AreEqual(contentStartRow, scheme.ContentStartRow);
            TestAssert.AreEqual(endRow, scheme.EndRow);
            TestAssert.IsTrue(scheme.IsValid);
            
            Console.WriteLine("✓ Create_ValidInput_CreatesScheme");
        }

        private static void Test_Create_InvalidSheetName_ThrowsException()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            
            // Act & Assert
            try
            {
                var scheme = Scheme.Create("", root, 3, 10);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Validation.InvalidSheetName));
                Console.WriteLine("✓ Create_InvalidSheetName_ThrowsException");
            }
        }

        private static void Test_Create_NullRoot_ThrowsException()
        {
            // Act & Assert
            try
            {
                var scheme = Scheme.Create("Sheet1", null, 3, 10);
                TestAssert.Fail("Should throw ArgumentNullException");
            }
            catch (ArgumentNullException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Schema.RootNodeIsNull));
                Console.WriteLine("✓ Create_NullRoot_ThrowsException");
            }
        }

        private static void Test_Create_InvalidContentStartRow_ThrowsException()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            
            // Act & Assert
            try
            {
                var scheme = Scheme.Create("Sheet1", root, 0, 10);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Schema.InvalidContentStartRow));
                Console.WriteLine("✓ Create_InvalidContentStartRow_ThrowsException");
            }
        }

        private static void Test_Create_EndRowLessThanStartRow_ThrowsException()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            
            // Act & Assert
            try
            {
                var scheme = Scheme.Create("Sheet1", root, 10, 5);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Schema.EndRowLessThanStartRow));
                Console.WriteLine("✓ Create_EndRowLessThanStartRow_ThrowsException");
            }
        }

        private static void Test_Empty_CreatesEmptyScheme()
        {
            // Act
            var scheme = Scheme.Empty("Sheet1");
            
            // Assert
            TestAssert.AreEqual("Sheet1", scheme.SheetName);
            TestAssert.IsNotNull(scheme.Root);
            TestAssert.AreEqual(SchemeConstants.Sheet.DataStartRow, scheme.ContentStartRow);
            TestAssert.AreEqual(SchemeConstants.Sheet.DataStartRow, scheme.EndRow);
            TestAssert.IsTrue(scheme.IsValid);
            
            Console.WriteLine("✓ Empty_CreatesEmptyScheme");
        }

        private static void Test_GetLinearNodes_ReturnsAllNodes()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var child1 = SchemeNode.Create("child1", 2, 1);
            var child2 = SchemeNode.Create("child2", 2, 2);
            root.AddChild(child1);
            root.AddChild(child2);
            
            var scheme = Scheme.Create("Sheet1", root, 3, 10);
            
            // Act
            var linearNodes = scheme.GetLinearNodes().ToList();
            
            // Assert
            TestAssert.AreEqual(3, linearNodes.Count);
            TestAssert.IsTrue(linearNodes.Contains(root));
            TestAssert.IsTrue(linearNodes.Contains(child1));
            TestAssert.IsTrue(linearNodes.Contains(child2));
            
            Console.WriteLine("✓ GetLinearNodes_ReturnsAllNodes");
        }

        private static void Test_GetNodesByType_FiltersCorrectly()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var propertyNode = SchemeNode.Create("property", 2, 1);
            var arrayNode = SchemeNode.Create("items$[]", 2, 2);
            root.AddChild(propertyNode);
            root.AddChild(arrayNode);
            
            var scheme = Scheme.Create("Sheet1", root, 3, 10);
            
            // Act
            var propertyNodes = scheme.GetNodesByType(SchemeNodeType.Property).ToList();
            var arrayNodes = scheme.GetNodesByType(SchemeNodeType.Array).ToList();
            
            // Assert
            TestAssert.AreEqual(1, propertyNodes.Count);
            TestAssert.AreEqual(propertyNode, propertyNodes[0]);
            TestAssert.AreEqual(1, arrayNodes.Count);
            TestAssert.AreEqual(arrayNode, arrayNodes[0]);
            
            Console.WriteLine("✓ GetNodesByType_FiltersCorrectly");
        }

        private static void Test_GetNodesByDepth_FiltersCorrectly()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var level1 = SchemeNode.Create("level1${}", 2, 1);
            var level2 = SchemeNode.Create("level2", 3, 2);
            root.AddChild(level1);
            level1.AddChild(level2);
            
            var scheme = Scheme.Create("Sheet1", root, 4, 10);
            
            // Act
            var depth0Nodes = scheme.GetNodesByDepth(0).ToList();
            var depth1Nodes = scheme.GetNodesByDepth(1).ToList();
            var depth2Nodes = scheme.GetNodesByDepth(2).ToList();
            
            // Assert
            TestAssert.AreEqual(1, depth0Nodes.Count);
            TestAssert.AreEqual(root, depth0Nodes[0]);
            TestAssert.AreEqual(1, depth1Nodes.Count);
            TestAssert.AreEqual(level1, depth1Nodes[0]);
            TestAssert.AreEqual(1, depth2Nodes.Count);
            TestAssert.AreEqual(level2, depth2Nodes[0]);
            
            Console.WriteLine("✓ GetNodesByDepth_FiltersCorrectly");
        }

        private static void Test_FindNodeByPath_ReturnsCorrectNode()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var config = SchemeNode.Create("config${}", 2, 1);
            var database = SchemeNode.Create("database${}", 3, 2);
            var connection = SchemeNode.Create("connection", 4, 3);
            
            root.AddChild(config);
            config.AddChild(database);
            database.AddChild(connection);
            
            var scheme = Scheme.Create("Sheet1", root, 5, 10);
            
            // Act
            var foundNode = scheme.FindNodeByPath("config/database/connection");
            var notFound = scheme.FindNodeByPath("config/nonexistent");
            
            // Assert
            TestAssert.AreEqual(connection, foundNode);
            TestAssert.IsNull(notFound);
            
            Console.WriteLine("✓ FindNodeByPath_ReturnsCorrectNode");
        }

        private static void Test_Validate_ValidScheme_ReturnsSuccess()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var child = SchemeNode.Create("child", 2, 1);
            root.AddChild(child);
            
            var scheme = Scheme.Create("Sheet1", root, 3, 10);
            
            // Act
            var result = scheme.Validate();
            
            // Assert
            TestAssert.IsTrue(result.IsValid);
            TestAssert.IsFalse(result.HasErrors);
            TestAssert.AreEqual(0, result.Errors.Count);
            
            Console.WriteLine("✓ Validate_ValidScheme_ReturnsSuccess");
        }

        private static void Test_Clone_CreatesDeepCopy()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var child = SchemeNode.Create("child", 2, 1);
            root.AddChild(child);
            
            var original = Scheme.Create("Sheet1", root, 3, 10);
            
            // Act
            var clone = original.Clone();
            
            // Assert
            TestAssert.AreEqual(original.SheetName, clone.SheetName);
            TestAssert.AreEqual(original.ContentStartRow, clone.ContentStartRow);
            TestAssert.AreEqual(original.EndRow, clone.EndRow);
            TestAssert.AreNotEqual(original.Root.Id, clone.Root.Id);
            TestAssert.AreEqual(original.NodeCount, clone.NodeCount);
            
            Console.WriteLine("✓ Clone_CreatesDeepCopy");
        }

        private static void Test_DataRowCount_CalculatesCorrectly()
        {
            // Arrange
            var testCases = new[]
            {
                (3, 10, 8),   // 10 - 3 + 1 = 8
                (5, 5, 1),    // 5 - 5 + 1 = 1
                (1, 100, 100) // 100 - 1 + 1 = 100
            };
            
            foreach (var (startRow, endRow, expectedCount) in testCases)
            {
                // Act
                var scheme = Scheme.Create("Sheet1", SchemeNode.CreateRoot(), startRow, endRow);
                
                // Assert
                TestAssert.AreEqual(expectedCount, scheme.DataRowCount, 
                    $"DataRowCount for rows {startRow}-{endRow} should be {expectedCount}");
            }
            
            Console.WriteLine("✓ DataRowCount_CalculatesCorrectly");
        }
    }
}