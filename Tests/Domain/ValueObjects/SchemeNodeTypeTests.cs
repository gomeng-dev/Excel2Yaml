using System;
using System.Linq;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;

namespace ExcelToYamlAddin.Tests.Domain.ValueObjects
{
    /// <summary>
    /// SchemeNodeType 값 객체 단위 테스트
    /// </summary>
    public class SchemeNodeTypeTests
    {
        public static void RunAllTests()
        {
            Console.WriteLine("=== SchemeNodeType Tests ===");
            
            Test_StaticInstances_AreInitialized();
            Test_IsContainer_ReturnsCorrectValue();
            Test_IsDynamic_ReturnsCorrectValue();
            Test_CanHaveData_ReturnsCorrectValue();
            Test_FromSchemeName_Property_ReturnsProperty();
            Test_FromSchemeName_Array_ReturnsArray();
            Test_FromSchemeName_Map_ReturnsMap();
            Test_FromSchemeName_Key_ReturnsKey();
            Test_FromSchemeName_Value_ReturnsValue();
            Test_FromSchemeName_Ignore_ReturnsIgnore();
            Test_GetAll_ReturnsAllTypes();
            Test_Equals_SameType_ReturnsTrue();
            Test_ImplicitStringConversion_ReturnsCode();
            
            Console.WriteLine("SchemeNodeType 테스트 완료!\n");
        }

        private static void Test_StaticInstances_AreInitialized()
        {
            // Assert
            TestAssert.IsNotNull(SchemeNodeType.Property, "Property should be initialized");
            TestAssert.IsNotNull(SchemeNodeType.Map, "Map should be initialized");
            TestAssert.IsNotNull(SchemeNodeType.Array, "Array should be initialized");
            TestAssert.IsNotNull(SchemeNodeType.Key, "Key should be initialized");
            TestAssert.IsNotNull(SchemeNodeType.Value, "Value should be initialized");
            TestAssert.IsNotNull(SchemeNodeType.Ignore, "Ignore should be initialized");
            
            Console.WriteLine("✓ StaticInstances_AreInitialized");
        }

        private static void Test_IsContainer_ReturnsCorrectValue()
        {
            // Assert
            TestAssert.IsTrue(SchemeNodeType.Map.IsContainer, "Map should be container");
            TestAssert.IsTrue(SchemeNodeType.Array.IsContainer, "Array should be container");
            TestAssert.IsFalse(SchemeNodeType.Property.IsContainer, "Property should not be container");
            TestAssert.IsFalse(SchemeNodeType.Key.IsContainer, "Key should not be container");
            TestAssert.IsFalse(SchemeNodeType.Value.IsContainer, "Value should not be container");
            TestAssert.IsFalse(SchemeNodeType.Ignore.IsContainer, "Ignore should not be container");
            
            Console.WriteLine("✓ IsContainer_ReturnsCorrectValue");
        }

        private static void Test_IsDynamic_ReturnsCorrectValue()
        {
            // Assert
            TestAssert.IsTrue(SchemeNodeType.Key.IsDynamic, "Key should be dynamic");
            TestAssert.IsTrue(SchemeNodeType.Value.IsDynamic, "Value should be dynamic");
            TestAssert.IsFalse(SchemeNodeType.Property.IsDynamic, "Property should not be dynamic");
            TestAssert.IsFalse(SchemeNodeType.Map.IsDynamic, "Map should not be dynamic");
            TestAssert.IsFalse(SchemeNodeType.Array.IsDynamic, "Array should not be dynamic");
            TestAssert.IsFalse(SchemeNodeType.Ignore.IsDynamic, "Ignore should not be dynamic");
            
            Console.WriteLine("✓ IsDynamic_ReturnsCorrectValue");
        }

        private static void Test_CanHaveData_ReturnsCorrectValue()
        {
            // Assert
            TestAssert.IsTrue(SchemeNodeType.Property.CanHaveData, "Property can have data");
            TestAssert.IsTrue(SchemeNodeType.Map.CanHaveData, "Map can have data");
            TestAssert.IsTrue(SchemeNodeType.Array.CanHaveData, "Array can have data");
            TestAssert.IsTrue(SchemeNodeType.Key.CanHaveData, "Key can have data");
            TestAssert.IsTrue(SchemeNodeType.Value.CanHaveData, "Value can have data");
            TestAssert.IsFalse(SchemeNodeType.Ignore.CanHaveData, "Ignore cannot have data");
            
            Console.WriteLine("✓ CanHaveData_ReturnsCorrectValue");
        }

        private static void Test_FromSchemeName_Property_ReturnsProperty()
        {
            // Arrange
            var testCases = new[] { "name", "title", "description", "" };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Property, nodeType, $"'{schemeName}' should return Property");
            }
            
            Console.WriteLine("✓ FromSchemeName_Property_ReturnsProperty");
        }

        private static void Test_FromSchemeName_Array_ReturnsArray()
        {
            // Arrange
            var testCases = new[] 
            { 
                SchemeConstants.Markers.ArrayStart,
                "items$[]",
                "data$[]$array"
            };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Array, nodeType, $"'{schemeName}' should return Array");
            }
            
            Console.WriteLine("✓ FromSchemeName_Array_ReturnsArray");
        }

        private static void Test_FromSchemeName_Map_ReturnsMap()
        {
            // Arrange
            var testCases = new[] 
            { 
                SchemeConstants.Markers.MapStart,
                "config${}",
                "settings${}$map"
            };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Map, nodeType, $"'{schemeName}' should return Map");
            }
            
            Console.WriteLine("✓ FromSchemeName_Map_ReturnsMap");
        }

        private static void Test_FromSchemeName_Key_ReturnsKey()
        {
            // Arrange
            var testCases = new[] 
            { 
                SchemeConstants.Markers.DynamicKey,
                "id$key",
                "name$key"
            };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Key, nodeType, $"'{schemeName}' should return Key");
            }
            
            Console.WriteLine("✓ FromSchemeName_Key_ReturnsKey");
        }

        private static void Test_FromSchemeName_Value_ReturnsValue()
        {
            // Arrange
            var testCases = new[] 
            { 
                SchemeConstants.Markers.DynamicValue,
                "data$value",
                "content$value"
            };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Value, nodeType, $"'{schemeName}' should return Value");
            }
            
            Console.WriteLine("✓ FromSchemeName_Value_ReturnsValue");
        }

        private static void Test_FromSchemeName_Ignore_ReturnsIgnore()
        {
            // Arrange
            var testCases = new[] 
            { 
                SchemeConstants.Markers.Ignore,
                "^"
            };
            
            foreach (var schemeName in testCases)
            {
                // Act
                var nodeType = SchemeNodeType.FromSchemeName(schemeName);
                
                // Assert
                TestAssert.AreEqual(SchemeNodeType.Ignore, nodeType, $"'{schemeName}' should return Ignore");
            }
            
            Console.WriteLine("✓ FromSchemeName_Ignore_ReturnsIgnore");
        }

        private static void Test_GetAll_ReturnsAllTypes()
        {
            // Act
            var allTypes = SchemeNodeType.GetAll().ToList();
            
            // Assert
            TestAssert.AreEqual(6, allTypes.Count, "Should return 6 types");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Property), "Should contain Property");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Map), "Should contain Map");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Array), "Should contain Array");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Key), "Should contain Key");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Value), "Should contain Value");
            TestAssert.IsTrue(allTypes.Contains(SchemeNodeType.Ignore), "Should contain Ignore");
            
            Console.WriteLine("✓ GetAll_ReturnsAllTypes");
        }

        private static void Test_Equals_SameType_ReturnsTrue()
        {
            // Arrange
            var property1 = SchemeNodeType.Property;
            var property2 = SchemeNodeType.Property;
            var map1 = SchemeNodeType.Map;
            
            // Assert
            TestAssert.IsTrue(property1 == property2, "Same types should be equal");
            TestAssert.IsTrue(map1.Equals(SchemeNodeType.Map), "Same types should be equal");
            TestAssert.IsFalse(SchemeNodeType.Property == SchemeNodeType.Map, "Different types should not be equal");
            TestAssert.IsTrue(SchemeNodeType.Property != SchemeNodeType.Map, "Different types should not be equal");
            
            Console.WriteLine("✓ Equals_SameType_ReturnsTrue");
        }

        private static void Test_ImplicitStringConversion_ReturnsCode()
        {
            // Act
            string propertyCode = SchemeNodeType.Property;
            string mapCode = SchemeNodeType.Map;
            string arrayCode = SchemeNodeType.Array;
            
            // Assert
            TestAssert.AreEqual("PROPERTY", propertyCode, "Property code should be PROPERTY");
            TestAssert.AreEqual("MAP", mapCode, "Map code should be MAP");
            TestAssert.AreEqual("ARRAY", arrayCode, "Array code should be ARRAY");
            
            Console.WriteLine("✓ ImplicitStringConversion_ReturnsCode");
        }
    }
}