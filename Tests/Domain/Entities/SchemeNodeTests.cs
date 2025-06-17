using System;
using System.Linq;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;

namespace ExcelToYamlAddin.Tests.Domain.Entities
{
    /// <summary>
    /// SchemeNode 엔티티 단위 테스트
    /// </summary>
    public class SchemeNodeTests
    {
        public static void RunAllTests()
        {
            Console.WriteLine("=== SchemeNode Tests ===");
            
            Test_Create_ValidInput_CreatesNode();
            Test_Create_EmptySchemeName_ThrowsException();
            Test_CreateRoot_CreatesRootNode();
            Test_AddChild_ValidChild_AddsSuccessfully();
            Test_AddChild_ToValueNode_ThrowsException();
            Test_AddChild_ToIgnoreNode_ThrowsException();
            Test_GetFullPath_ReturnsCorrectPath();
            Test_FindChildByKey_ReturnsCorrectChild();
            Test_FindChildByType_ReturnsCorrectChild();
            Test_Linear_ReturnsAllNodesInOrder();
            Test_Clone_CreatesDeepCopy();
            Test_Validate_ValidNode_ReturnsSuccess();
            Test_ExtractKey_FromSchemeName();
            
            Console.WriteLine("SchemeNode 테스트 완료!\n");
        }

        private static void Test_Create_ValidInput_CreatesNode()
        {
            // Arrange
            var schemeName = "name";
            var row = 1;
            var column = 1;
            
            // Act
            var node = SchemeNode.Create(schemeName, row, column);
            
            // Assert
            TestAssert.AreEqual("name", node.Key);
            TestAssert.AreEqual(SchemeNodeType.Property, node.NodeType);
            TestAssert.AreEqual(schemeName, node.SchemeName);
            TestAssert.AreEqual(1, node.Position.Row);
            TestAssert.AreEqual(1, node.Position.Column);
            
            Console.WriteLine("✓ Create_ValidInput_CreatesNode");
        }

        private static void Test_Create_EmptySchemeName_ThrowsException()
        {
            // Arrange & Act & Assert
            try
            {
                var node = SchemeNode.Create("", 1, 1);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Schema.SchemeNameIsEmpty));
                Console.WriteLine("✓ Create_EmptySchemeName_ThrowsException");
            }
        }

        private static void Test_CreateRoot_CreatesRootNode()
        {
            // Act
            var root = SchemeNode.CreateRoot();
            
            // Assert
            TestAssert.AreEqual("", root.Key);
            TestAssert.AreEqual(SchemeNodeType.Map, root.NodeType);
            TestAssert.IsTrue(root.IsRoot);
            TestAssert.AreEqual(0, root.Depth);
            TestAssert.AreEqual(SchemeConstants.Position.RootNodeRow, root.Position.Row);
            TestAssert.AreEqual(SchemeConstants.Position.RootNodeColumn, root.Position.Column);
            
            Console.WriteLine("✓ CreateRoot_CreatesRootNode");
        }

        private static void Test_AddChild_ValidChild_AddsSuccessfully()
        {
            // Arrange
            var parent = SchemeNode.Create("parent${}", 1, 1);
            var child = SchemeNode.Create("child", 2, 2);
            
            // Act
            parent.AddChild(child);
            
            // Assert
            TestAssert.AreEqual(1, parent.ChildCount);
            TestAssert.AreEqual(parent, child.Parent);
            TestAssert.AreEqual(1, child.Depth);
            TestAssert.IsTrue(parent.Children.Contains(child));
            
            Console.WriteLine("✓ AddChild_ValidChild_AddsSuccessfully");
        }

        private static void Test_AddChild_ToValueNode_ThrowsException()
        {
            // Arrange
            var valueNode = SchemeNode.Create("$value", 1, 1);
            var child = SchemeNode.Create("child", 2, 2);
            
            // Act & Assert
            try
            {
                valueNode.AddChild(child);
                TestAssert.Fail("Should throw InvalidOperationException");
            }
            catch (InvalidOperationException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Validation.CannotAddChildToValueNode));
                Console.WriteLine("✓ AddChild_ToValueNode_ThrowsException");
            }
        }

        private static void Test_AddChild_ToIgnoreNode_ThrowsException()
        {
            // Arrange
            var ignoreNode = SchemeNode.Create("^", 1, 1);
            var child = SchemeNode.Create("child", 2, 2);
            
            // Act & Assert
            try
            {
                ignoreNode.AddChild(child);
                TestAssert.Fail("Should throw InvalidOperationException");
            }
            catch (InvalidOperationException ex)
            {
                TestAssert.IsTrue(ex.Message.Contains(ErrorMessages.Validation.CannotAddChildToIgnoreNode));
                Console.WriteLine("✓ AddChild_ToIgnoreNode_ThrowsException");
            }
        }

        private static void Test_GetFullPath_ReturnsCorrectPath()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var level1 = SchemeNode.Create("config${}", 2, 1);
            var level2 = SchemeNode.Create("database${}", 3, 2);
            var level3 = SchemeNode.Create("connection", 4, 3);
            
            root.AddChild(level1);
            level1.AddChild(level2);
            level2.AddChild(level3);
            
            // Act
            var path = level3.GetFullPath();
            
            // Assert
            TestAssert.AreEqual("config/database/connection", path);
            
            Console.WriteLine("✓ GetFullPath_ReturnsCorrectPath");
        }

        private static void Test_FindChildByKey_ReturnsCorrectChild()
        {
            // Arrange
            var parent = SchemeNode.Create("parent${}", 1, 1);
            var child1 = SchemeNode.Create("child1", 2, 2);
            var child2 = SchemeNode.Create("child2", 2, 3);
            
            parent.AddChild(child1);
            parent.AddChild(child2);
            
            // Act
            var found = parent.FindChildByKey("child2");
            
            // Assert
            TestAssert.AreEqual(child2, found);
            TestAssert.IsNull(parent.FindChildByKey("nonexistent"));
            
            Console.WriteLine("✓ FindChildByKey_ReturnsCorrectChild");
        }

        private static void Test_FindChildByType_ReturnsCorrectChild()
        {
            // Arrange
            var parent = SchemeNode.Create("parent${}", 1, 1);
            var propertyChild = SchemeNode.Create("property", 2, 2);
            var arrayChild = SchemeNode.Create("items$[]", 2, 3);
            
            parent.AddChild(propertyChild);
            parent.AddChild(arrayChild);
            
            // Act
            var foundArray = parent.FindChildByType(SchemeNodeType.Array);
            var foundProperty = parent.FindChildByType(SchemeNodeType.Property);
            
            // Assert
            TestAssert.AreEqual(arrayChild, foundArray);
            TestAssert.AreEqual(propertyChild, foundProperty);
            
            Console.WriteLine("✓ FindChildByType_ReturnsCorrectChild");
        }

        private static void Test_Linear_ReturnsAllNodesInOrder()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var child1 = SchemeNode.Create("child1", 2, 1);
            var child2 = SchemeNode.Create("child2", 2, 2);
            var grandchild = SchemeNode.Create("grandchild", 3, 1);
            
            root.AddChild(child1);
            root.AddChild(child2);
            child1.AddChild(grandchild);
            
            // Act
            var linearNodes = root.Linear().ToList();
            
            // Assert
            TestAssert.AreEqual(4, linearNodes.Count);
            TestAssert.AreEqual(root, linearNodes[0]);
            TestAssert.AreEqual(child1, linearNodes[1]);
            TestAssert.AreEqual(grandchild, linearNodes[2]);
            TestAssert.AreEqual(child2, linearNodes[3]);
            
            Console.WriteLine("✓ Linear_ReturnsAllNodesInOrder");
        }

        private static void Test_Clone_CreatesDeepCopy()
        {
            // Arrange
            var original = SchemeNode.Create("parent${}", 1, 1);
            var child = SchemeNode.Create("child", 2, 2);
            original.AddChild(child);
            
            // Act
            var clone = original.Clone();
            
            // Assert
            TestAssert.AreNotEqual(original.Id, clone.Id);
            TestAssert.AreEqual(original.Key, clone.Key);
            TestAssert.AreEqual(original.NodeType, clone.NodeType);
            TestAssert.AreEqual(original.ChildCount, clone.ChildCount);
            TestAssert.AreNotEqual(original.Children.First().Id, clone.Children.First().Id);
            
            Console.WriteLine("✓ Clone_CreatesDeepCopy");
        }

        private static void Test_Validate_ValidNode_ReturnsSuccess()
        {
            // Arrange
            var root = SchemeNode.CreateRoot();
            var child = SchemeNode.Create("child", 2, 2);
            root.AddChild(child);
            
            // Act
            var result = root.Validate();
            
            // Assert
            TestAssert.IsTrue(result.IsValid);
            TestAssert.AreEqual(0, result.Errors.Count);
            
            Console.WriteLine("✓ Validate_ValidNode_ReturnsSuccess");
        }

        private static void Test_ExtractKey_FromSchemeName()
        {
            // Arrange
            var testCases = new[]
            {
                ("name", "name"),
                ("items$[]", "items"),
                ("config${}", "config"),
                ("$key", ""),
                ("$value", ""),
                ("data$MAP", "data"),
                ("^", "^")
            };
            
            foreach (var (schemeName, expectedKey) in testCases)
            {
                // Act
                var node = SchemeNode.Create(schemeName, 1, 2); // column 2 to avoid root container logic
                
                // Assert
                TestAssert.AreEqual(expectedKey, node.Key, $"Key for '{schemeName}' should be '{expectedKey}'");
            }
            
            Console.WriteLine("✓ ExtractKey_FromSchemeName");
        }
    }
}
