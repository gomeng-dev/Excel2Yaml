using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Tests.Utilities;
using Microsoft.Office.Interop.Excel;

namespace ExcelToYamlAddin.Tests
{
    /// <summary>
    /// 테스트 실행기
    /// </summary>
    public class TestRunner
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<TestRunner>();
        private readonly List<TestResult> _results = new List<TestResult>();

        public class TestResult
        {
            public string TestClass { get; set; }
            public string TestMethod { get; set; }
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public TimeSpan Duration { get; set; }
        }

        /// <summary>
        /// 테스트 시트를 생성하고 테스트를 실행합니다.
        /// </summary>
        public void GenerateTestSheetsAndRun()
        {
            Logger.Information("\n테스트 시트 생성 및 테스트 실행 시작...");
            
            try
            {
                // 테스트 시트 생성
                var sheets = GenerateAllTestSheets();
                Logger.Information($"{sheets.Count}개의 테스트 시트 생성 완료");
                
                // 각 시트에 대해 테스트 실행
                foreach (var sheet in sheets)
                {
                    Logger.Information($"\n'{sheet.Name}' 시트로 테스트 실행 중...");
                    sheet.Activate();
                    RunIntegrationTests();
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "테스트 시트 생성 및 실행 중 오류");
                throw;
            }
        }

        /// <summary>
        /// 모든 테스트 시트를 생성합니다.
        /// </summary>
        public List<Worksheet> GenerateAllTestSheets()
        {
            var sheets = new List<Worksheet>();
            
            try
            {
                sheets.Add(TestSheetGenerator.CreateBasicTestSheet());
                sheets.Add(TestSheetGenerator.CreateArrayTestSheet());
                sheets.Add(TestSheetGenerator.CreateMapTestSheet());
                sheets.Add(TestSheetGenerator.CreateKeyValueTestSheet());
                sheets.Add(TestSheetGenerator.CreateComplexTestSheet());
                
                return sheets;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "테스트 시트 생성 중 오류");
                throw;
            }
        }

        public void RunAllTests()
        {
            Logger.Information("테스트 실행 시작...");
            var startTime = DateTime.Now;

            // Excel 환경 확인
            if (!CheckExcelEnvironment())
            {
                Logger.Error("Excel 환경을 확인할 수 없습니다. 테스트를 중단합니다.");
                return;
            }

            // 테스트 클래스 찾기
            var testClasses = GetTestClasses();
            Logger.Information($"{testClasses.Count}개의 테스트 클래스 발견");

            foreach (var testClass in testClasses)
            {
                RunTestsInClass(testClass);
            }

            // 결과 요약
            var totalTime = DateTime.Now - startTime;
            PrintSummary(totalTime);
        }

        public void RunIntegrationTests()
        {
            Logger.Information("\n통합 테스트 실행 시작...");
            var startTime = DateTime.Now;

            // Excel 환경 확인
            if (!CheckExcelEnvironment())
            {
                Logger.Error("Excel 환경을 확인할 수 없습니다. 테스트를 중단합니다.");
                return;
            }

            // 통합 테스트만 실행
            var integrationTestClasses = GetTestClasses()
                .Where(t => t.Name.Contains("Integration"))
                .ToList();

            Logger.Information($"{integrationTestClasses.Count}개의 통합 테스트 클래스 발견");

            foreach (var testClass in integrationTestClasses)
            {
                RunTestsInClass(testClass);
            }

            // 결과 요약
            var totalTime = DateTime.Now - startTime;
            PrintSummary(totalTime);
        }

        private bool CheckExcelEnvironment()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null)
                {
                    Logger.Error("Excel Application 객체를 찾을 수 없습니다.");
                    return false;
                }

                var activeSheet = app.ActiveSheet;
                if (activeSheet == null)
                {
                    Logger.Warning("활성화된 시트가 없습니다.");
                    return false;
                }

                Logger.Information($"Excel 환경 확인 완료 - 현재 시트: {activeSheet.Name}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Excel 환경 확인 중 오류: {ex.Message}");
                return false;
            }
        }

        private List<Type> GetTestClasses()
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetTypes()
                .Where(t => t.Namespace != null && 
                           t.Namespace.StartsWith("ExcelToYamlAddin.Tests") &&
                           t.Name.EndsWith("Tests") &&
                           !t.IsAbstract &&
                           t.IsClass)
                .ToList();
        }

        private void RunTestsInClass(Type testClass)
        {
            Logger.Information($"\n--- {testClass.Name} 실행 중 ---");

            var testMethods = testClass.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => !m.IsSpecialName && m.DeclaringType == testClass)
                .ToList();

            if (testMethods.Count == 0)
            {
                Logger.Warning($"{testClass.Name}에 테스트 메서드가 없습니다.");
                return;
            }

            object testInstance = null;
            try
            {
                testInstance = Activator.CreateInstance(testClass);
            }
            catch (Exception ex)
            {
                Logger.Error($"{testClass.Name} 인스턴스 생성 실패: {ex.Message}");
                return;
            }

            foreach (var method in testMethods)
            {
                RunTestMethod(testClass, testInstance, method);
            }
        }

        private void RunTestMethod(Type testClass, object testInstance, MethodInfo method)
        {
            var result = new TestResult
            {
                TestClass = testClass.Name,
                TestMethod = method.Name
            };

            var startTime = DateTime.Now;

            try
            {
                Logger.Debug($"  - {method.Name} 실행 중...");
                method.Invoke(testInstance, null);
                result.Success = true;
                Logger.Information($"  ✓ {method.Name} 성공");
            }
            catch (Exception ex)
            {
                result.Success = false;
                var innerEx = ex.InnerException ?? ex;
                result.ErrorMessage = innerEx.Message;
                Logger.Error($"  ✗ {method.Name} 실패: {innerEx.Message}");
                if (innerEx.StackTrace != null)
                {
                    Logger.Debug($"    StackTrace: {innerEx.StackTrace}");
                }
            }
            finally
            {
                result.Duration = DateTime.Now - startTime;
                _results.Add(result);
            }
        }

        private void PrintSummary(TimeSpan totalTime)
        {
            var successCount = _results.Count(r => r.Success);
            var failCount = _results.Count(r => !r.Success);
            var totalCount = _results.Count;

            var summary = new StringBuilder();
            summary.AppendLine("\n");
            summary.AppendLine("=".PadRight(50, '='));
            summary.AppendLine("테스트 결과 요약");
            summary.AppendLine("=".PadRight(50, '='));
            summary.AppendLine($"전체: {totalCount}개");
            summary.AppendLine($"성공: {successCount}개");
            summary.AppendLine($"실패: {failCount}개");
            summary.AppendLine($"소요 시간: {totalTime.TotalSeconds:F2}초");

            if (failCount > 0)
            {
                summary.AppendLine("\n실패한 테스트:");
                foreach (var failed in _results.Where(r => !r.Success))
                {
                    summary.AppendLine($"  - {failed.TestClass}.{failed.TestMethod}: {failed.ErrorMessage}");
                }
            }

            summary.AppendLine("=".PadRight(50, '='));

            Logger.Information(summary.ToString());
        }

        public string GetResultSummary()
        {
            var successCount = _results.Count(r => r.Success);
            var failCount = _results.Count(r => !r.Success);
            return $"테스트 결과: 성공 {successCount}, 실패 {failCount}";
        }
    }
}