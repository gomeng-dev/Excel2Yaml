using System;
using ExcelToYamlAddin.Tests.Domain.ValueObjects;
using ExcelToYamlAddin.Tests.Domain.Entities;

namespace ExcelToYamlAddin.Tests
{
    /// <summary>
    /// 도메인 모델 테스트 실행기
    /// VSTO 애드인 프로젝트에서 수동으로 테스트를 실행하기 위한 클래스
    /// </summary>
    public class TestRunner
    {
        /// <summary>
        /// 모든 도메인 모델 테스트를 실행합니다.
        /// </summary>
        public static void RunAllTests()
        {
            Console.WriteLine("======================================");
            Console.WriteLine("도메인 모델 단위 테스트 실행 시작");
            Console.WriteLine("======================================\n");

            try
            {
                // 값 객체 테스트
                CellPositionTests.RunAllTests();
                SchemeNodeTypeTests.RunAllTests();
                
                // 엔티티 테스트
                SchemeNodeTests.RunAllTests();
                SchemeTests.RunAllTests();

                Console.WriteLine("\n======================================");
                Console.WriteLine("✅ 모든 테스트가 성공적으로 완료되었습니다!");
                Console.WriteLine("======================================");
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n======================================");
                Console.WriteLine("❌ 테스트 실행 중 오류 발생!");
                Console.WriteLine($"오류: {ex.Message}");
                Console.WriteLine($"스택 추적:\n{ex.StackTrace}");
                Console.WriteLine("======================================");
                throw;
            }
        }

        /// <summary>
        /// 특정 테스트 스위트를 실행합니다.
        /// </summary>
        public static void RunTestSuite(string suiteName)
        {
            Console.WriteLine($"테스트 스위트 '{suiteName}' 실행 중...\n");

            try
            {
                switch (suiteName.ToLower())
                {
                    case "cellposition":
                        CellPositionTests.RunAllTests();
                        break;
                    case "schemenodetype":
                        SchemeNodeTypeTests.RunAllTests();
                        break;
                    case "schemenode":
                        SchemeNodeTests.RunAllTests();
                        break;
                    case "scheme":
                        SchemeTests.RunAllTests();
                        break;
                    default:
                        Console.WriteLine($"알 수 없는 테스트 스위트: {suiteName}");
                        Console.WriteLine("사용 가능한 스위트: cellposition, schemenodetype, schemenode, scheme");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"테스트 '{suiteName}' 실행 중 오류: {ex.Message}");
                throw;
            }
        }
    }
}