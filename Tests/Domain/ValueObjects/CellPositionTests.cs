using System;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;

namespace ExcelToYamlAddin.Tests.Domain.ValueObjects
{
    /// <summary>
    /// CellPosition 값 객체 단위 테스트
    /// </summary>
    public class CellPositionTests
    {
        public static void RunAllTests()
        {
            Console.WriteLine("=== CellPosition Tests ===");
            
            Test_Constructor_ValidValues_CreatesInstance();
            Test_Constructor_InvalidRow_ThrowsException();
            Test_Constructor_InvalidColumn_ThrowsException();
            Test_FromAddress_ValidAddress_CreatesInstance();
            Test_FromAddress_InvalidAddress_ThrowsException();
            Test_ToAddress_ReturnsCorrectAddress();
            Test_Equals_SameValues_ReturnsTrue();
            Test_Equals_DifferentValues_ReturnsFalse();
            Test_ToString_ReturnsFormattedString();
            
            Console.WriteLine("CellPosition 테스트 완료!\n");
        }

        private static void Test_Constructor_ValidValues_CreatesInstance()
        {
            // Arrange & Act
            var position = new CellPosition(1, 1);
            
            // Assert
            TestAssert.AreEqual(1, position.Row, "Row should be 1");
            TestAssert.AreEqual(1, position.Column, "Column should be 1");
            Console.WriteLine("✓ Constructor_ValidValues_CreatesInstance");
        }

        private static void Test_Constructor_InvalidRow_ThrowsException()
        {
            // Arrange & Act & Assert
            try
            {
                var position = new CellPosition(0, 1);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.AreEqual(ErrorMessages.Validation.RowLessThanOne, ex.Message);
                Console.WriteLine("✓ Constructor_InvalidRow_ThrowsException");
            }
        }

        private static void Test_Constructor_InvalidColumn_ThrowsException()
        {
            // Arrange & Act & Assert
            try
            {
                var position = new CellPosition(1, 0);
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException ex)
            {
                TestAssert.AreEqual(ErrorMessages.Validation.ColumnLessThanOne, ex.Message);
                Console.WriteLine("✓ Constructor_InvalidColumn_ThrowsException");
            }
        }

        private static void Test_FromAddress_ValidAddress_CreatesInstance()
        {
            // Arrange & Act
            var position = CellPosition.FromAddress("A1");
            
            // Assert
            TestAssert.AreEqual(1, position.Row, "Row should be 1");
            TestAssert.AreEqual(1, position.Column, "Column should be 1");
            Console.WriteLine("✓ FromAddress_ValidAddress_CreatesInstance");
        }

        private static void Test_FromAddress_InvalidAddress_ThrowsException()
        {
            // Arrange & Act & Assert
            try
            {
                var position = CellPosition.FromAddress("InvalidAddress");
                TestAssert.Fail("Should throw ArgumentException");
            }
            catch (ArgumentException)
            {
                Console.WriteLine("✓ FromAddress_InvalidAddress_ThrowsException");
            }
        }

        private static void Test_ToAddress_ReturnsCorrectAddress()
        {
            // Arrange
            var testCases = new[]
            {
                (1, 1, "A1"),
                (10, 26, "Z10"),
                (100, 27, "AA100"),
                (1, 52, "AZ1"),
                (1, 53, "BA1")
            };

            foreach (var (row, column, expected) in testCases)
            {
                // Act
                var position = new CellPosition(row, column);
                var address = position.Address;

                // Assert
                TestAssert.AreEqual(expected, address, $"Address for ({row},{column}) should be {expected}");
            }
            
            Console.WriteLine("✓ ToAddress_ReturnsCorrectAddress");
        }

        private static void Test_Equals_SameValues_ReturnsTrue()
        {
            // Arrange
            var position1 = new CellPosition(5, 10);
            var position2 = new CellPosition(5, 10);
            
            // Act & Assert
            TestAssert.IsTrue(position1.Equals(position2), "Same positions should be equal");
            TestAssert.IsTrue(position1 == position2, "Same positions should be equal with == operator");
            Console.WriteLine("✓ Equals_SameValues_ReturnsTrue");
        }

        private static void Test_Equals_DifferentValues_ReturnsFalse()
        {
            // Arrange
            var position1 = new CellPosition(5, 10);
            var position2 = new CellPosition(5, 11);
            var position3 = new CellPosition(6, 10);
            
            // Act & Assert
            TestAssert.IsFalse(position1.Equals(position2), "Different positions should not be equal");
            TestAssert.IsFalse(position1.Equals(position3), "Different positions should not be equal");
            TestAssert.IsTrue(position1 != position2, "Different positions should not be equal with != operator");
            Console.WriteLine("✓ Equals_DifferentValues_ReturnsFalse");
        }

        private static void Test_ToString_ReturnsFormattedString()
        {
            // Arrange
            var position = new CellPosition(5, 10);
            
            // Act
            var result = position.ToString();
            
            // Assert
            TestAssert.AreEqual("[5, 10]", result, "ToString should return formatted string");
            Console.WriteLine("✓ ToString_ReturnsFormattedString");
        }
    }
}