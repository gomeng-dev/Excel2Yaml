using System;
using Microsoft.Office.Interop.Excel;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;

namespace ExcelToYamlAddin.Tests.Utilities
{
    /// <summary>
    /// 테스트용 Excel 시트를 생성하는 유틸리티 클래스
    /// </summary>
    public static class TestSheetGenerator
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger(nameof(TestSheetGenerator));

        /// <summary>
        /// 기본 테스트 시트를 생성합니다.
        /// </summary>
        public static Worksheet CreateBasicTestSheet(string sheetName = "!TestSheet")
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                // 기존 시트가 있으면 삭제
                DeleteSheetIfExists(workbook, sheetName);
                
                // 새 시트 생성
                var sheet = workbook.Worksheets.Add() as Worksheet;
                sheet.Name = sheetName;
                
                Logger.Information($"테스트 시트 생성됨: {sheetName}");
                
                // 기본 스키마 구조 생성
                CreateBasicSchema(sheet);
                
                // 테스트 데이터 추가
                AddTestData(sheet);
                
                return sheet;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "테스트 시트 생성 실패");
                throw;
            }
        }

        /// <summary>
        /// 배열 테스트 시트를 생성합니다.
        /// </summary>
        public static Worksheet CreateArrayTestSheet(string sheetName = "!ArrayTest")
        {
            var sheet = CreateEmptySheet(sheetName);
            
            // 배열 스키마 구조
            sheet.Cells[1, 1].Value = "# 배열 테스트 시트";
            sheet.Cells[2, 1].Value = SchemeConstants.NodeTypes.Array;
            sheet.Cells[3, 2].Value = "id";
            sheet.Cells[3, 3].Value = "name";
            sheet.Cells[3, 4].Value = "age";
            sheet.Cells[3, 5].Value = "active";
            
            // 병합 셀 (배열 컨테이너)
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 5]].Merge();
            
            // 스키마 끝 마커
            sheet.Cells[4, 1].Value = SchemeConstants.Markers.SchemeEnd;
            sheet.Range[sheet.Cells[4, 1], sheet.Cells[4, 5]].Merge();
            
            // 테스트 데이터
            sheet.Cells[5, 2].Value = 1;
            sheet.Cells[5, 3].Value = "Alice";
            sheet.Cells[5, 4].Value = 25;
            sheet.Cells[5, 5].Value = true;
            
            sheet.Cells[6, 2].Value = 2;
            sheet.Cells[6, 3].Value = "Bob";
            sheet.Cells[6, 4].Value = 30;
            sheet.Cells[6, 5].Value = false;
            
            sheet.Cells[7, 2].Value = 3;
            sheet.Cells[7, 3].Value = "Charlie";
            sheet.Cells[7, 4].Value = 35;
            sheet.Cells[7, 5].Value = true;
            
            Logger.Information($"배열 테스트 시트 생성 완료: {sheetName}");
            return sheet;
        }

        /// <summary>
        /// 맵(객체) 테스트 시트를 생성합니다.
        /// </summary>
        public static Worksheet CreateMapTestSheet(string sheetName = "!MapTest")
        {
            var sheet = CreateEmptySheet(sheetName);
            
            // 맵 스키마 구조
            sheet.Cells[1, 1].Value = "# 맵/객체 테스트 시트";
            sheet.Cells[2, 1].Value = SchemeConstants.NodeTypes.Map;
            sheet.Cells[3, 2].Value = "config";
            sheet.Cells[3, 3].Value = SchemeConstants.NodeTypes.Map;
            sheet.Cells[4, 3].Value = "version";
            sheet.Cells[4, 4].Value = "enabled";
            sheet.Cells[4, 5].Value = "timeout";
            
            // 병합 셀
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 5]].Merge();
            sheet.Range[sheet.Cells[3, 3], sheet.Cells[3, 5]].Merge();
            
            // 스키마 끝 마커
            sheet.Cells[5, 1].Value = SchemeConstants.Markers.SchemeEnd;
            sheet.Range[sheet.Cells[5, 1], sheet.Cells[5, 5]].Merge();
            
            // 테스트 데이터
            sheet.Cells[6, 3].Value = "1.0.0";
            sheet.Cells[6, 4].Value = true;
            sheet.Cells[6, 5].Value = 3000;
            
            Logger.Information($"맵 테스트 시트 생성 완료: {sheetName}");
            return sheet;
        }

        /// <summary>
        /// 동적 키-값 테스트 시트를 생성합니다.
        /// </summary>
        public static Worksheet CreateKeyValueTestSheet(string sheetName = "!KeyValueTest")
        {
            var sheet = CreateEmptySheet(sheetName);
            
            // 동적 키-값 스키마
            sheet.Cells[1, 1].Value = "# 동적 키-값 테스트 시트";
            sheet.Cells[2, 1].Value = SchemeConstants.NodeTypes.Map;
            sheet.Cells[3, 2].Value = SchemeConstants.NodeTypes.Key;
            sheet.Cells[3, 3].Value = SchemeConstants.NodeTypes.Value;
            
            // 병합 셀
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 3]].Merge();
            
            // 스키마 끝 마커
            sheet.Cells[4, 1].Value = SchemeConstants.Markers.SchemeEnd;
            sheet.Range[sheet.Cells[4, 1], sheet.Cells[4, 3]].Merge();
            
            // 테스트 데이터
            sheet.Cells[5, 2].Value = "setting1";
            sheet.Cells[5, 3].Value = "value1";
            
            sheet.Cells[6, 2].Value = "setting2";
            sheet.Cells[6, 3].Value = "value2";
            
            sheet.Cells[7, 2].Value = "setting3";
            sheet.Cells[7, 3].Value = 12345;
            
            Logger.Information($"키-값 테스트 시트 생성 완료: {sheetName}");
            return sheet;
        }

        /// <summary>
        /// 복잡한 중첩 구조 테스트 시트를 생성합니다.
        /// </summary>
        public static Worksheet CreateComplexTestSheet(string sheetName = "!ComplexTest")
        {
            var sheet = CreateEmptySheet(sheetName);
            
            // 복잡한 중첩 스키마
            sheet.Cells[1, 1].Value = "# 복잡한 중첩 구조 테스트";
            sheet.Cells[2, 1].Value = SchemeConstants.NodeTypes.Array;
            sheet.Cells[3, 2].Value = "id";
            sheet.Cells[3, 3].Value = "user";
            sheet.Cells[3, 4].Value = SchemeConstants.NodeTypes.Map;
            sheet.Cells[4, 4].Value = "name";
            sheet.Cells[4, 5].Value = "email";
            sheet.Cells[4, 6].Value = "roles";
            sheet.Cells[4, 7].Value = SchemeConstants.NodeTypes.Array;
            sheet.Cells[5, 7].Value = "role";
            
            // 병합 셀
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 7]].Merge();
            sheet.Range[sheet.Cells[3, 4], sheet.Cells[3, 7]].Merge();
            sheet.Range[sheet.Cells[4, 7], sheet.Cells[4, 7]].Merge();
            
            // 스키마 끝 마커
            sheet.Cells[6, 1].Value = SchemeConstants.Markers.SchemeEnd;
            sheet.Range[sheet.Cells[6, 1], sheet.Cells[6, 7]].Merge();
            
            // 테스트 데이터
            // 첫 번째 사용자
            sheet.Cells[7, 2].Value = 1;
            sheet.Cells[7, 3].Value = "user1";
            sheet.Cells[7, 4].Value = "John Doe";
            sheet.Cells[7, 5].Value = "john@example.com";
            sheet.Cells[7, 7].Value = "admin";
            
            sheet.Cells[8, 7].Value = "editor";
            
            // 두 번째 사용자
            sheet.Cells[9, 2].Value = 2;
            sheet.Cells[9, 3].Value = "user2";
            sheet.Cells[9, 4].Value = "Jane Smith";
            sheet.Cells[9, 5].Value = "jane@example.com";
            sheet.Cells[9, 7].Value = "viewer";
            
            Logger.Information($"복잡한 구조 테스트 시트 생성 완료: {sheetName}");
            return sheet;
        }

        private static Worksheet CreateEmptySheet(string sheetName)
        {
            var app = Globals.ThisAddIn.Application;
            var workbook = app.ActiveWorkbook;
            
            DeleteSheetIfExists(workbook, sheetName);
            
            var sheet = workbook.Worksheets.Add() as Worksheet;
            sheet.Name = sheetName;
            
            return sheet;
        }

        private static void DeleteSheetIfExists(Workbook workbook, string sheetName)
        {
            try
            {
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == sheetName)
                    {
                        var app = workbook.Application;
                        app.DisplayAlerts = false;
                        ws.Delete();
                        app.DisplayAlerts = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"기존 시트 삭제 중 오류: {ex.Message}");
            }
        }

        private static void CreateBasicSchema(Worksheet sheet)
        {
            // 주석 행
            sheet.Cells[1, 1].Value = "# 테스트 시트 - 기본 구조";
            
            // 스키마 정의
            sheet.Cells[2, 1].Value = SchemeConstants.NodeTypes.Map;
            sheet.Cells[3, 2].Value = "property1";
            sheet.Cells[3, 3].Value = "property2";
            sheet.Cells[3, 4].Value = "property3";
            
            // 병합 셀 설정
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 4]].Merge();
            
            // 스키마 끝 마커
            sheet.Cells[4, 1].Value = SchemeConstants.Markers.SchemeEnd;
            sheet.Range[sheet.Cells[4, 1], sheet.Cells[4, 4]].Merge();
            
            // 서식 설정
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[4, 4]].Borders.LineStyle = XlLineStyle.xlContinuous;
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[3, 4]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
        }

        private static void AddTestData(Worksheet sheet)
        {
            // 테스트 데이터 추가
            sheet.Cells[5, 2].Value = "Value1";
            sheet.Cells[5, 3].Value = 123;
            sheet.Cells[5, 4].Value = true;
            
            sheet.Cells[6, 2].Value = "Value2";
            sheet.Cells[6, 3].Value = 456;
            sheet.Cells[6, 4].Value = false;
            
            sheet.Cells[7, 2].Value = "Value3";
            sheet.Cells[7, 3].Value = 789;
            sheet.Cells[7, 4].Value = true;
            
            // 데이터 영역 서식
            sheet.Range[sheet.Cells[5, 1], sheet.Cells[7, 4]].Borders.LineStyle = XlLineStyle.xlContinuous;
        }
    }
}