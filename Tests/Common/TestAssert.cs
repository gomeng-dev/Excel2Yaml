using System;

namespace ExcelToYamlAddin.Tests.Common
{
    /// <summary>
    /// 테스트용 Assert 헬퍼 클래스
    /// </summary>
    internal static class TestAssert
    {
        public static void AreEqual<T>(T expected, T actual, string message = "")
        {
            if (!Equals(expected, actual))
            {
                throw new Exception($"Assert failed: Expected '{expected}' but was '{actual}'. {message}");
            }
        }

        public static void AreNotEqual<T>(T expected, T actual, string message = "")
        {
            if (Equals(expected, actual))
            {
                throw new Exception($"Assert failed: Expected not equal but both were '{expected}'. {message}");
            }
        }

        public static void IsTrue(bool condition, string message = "")
        {
            if (!condition)
            {
                throw new Exception($"Assert failed: Expected true but was false. {message}");
            }
        }

        public static void IsFalse(bool condition, string message = "")
        {
            if (condition)
            {
                throw new Exception($"Assert failed: Expected false but was true. {message}");
            }
        }

        public static void IsNull(object obj, string message = "")
        {
            if (obj != null)
            {
                throw new Exception($"Assert failed: Expected null but was not null. {message}");
            }
        }

        public static void IsNotNull(object obj, string message = "")
        {
            if (obj == null)
            {
                throw new Exception($"Assert failed: Expected not null but was null. {message}");
            }
        }

        public static void Fail(string message = "")
        {
            throw new Exception($"Assert failed: {message}");
        }

        public static void Throws<TException>(Action action, string message = "") where TException : Exception
        {
            try
            {
                action();
                throw new Exception($"Assert failed: Expected exception of type {typeof(TException).Name} but no exception was thrown. {message}");
            }
            catch (TException)
            {
                // Expected exception was thrown
            }
            catch (Exception ex)
            {
                throw new Exception($"Assert failed: Expected exception of type {typeof(TException).Name} but got {ex.GetType().Name}. {message}");
            }
        }
    }
}