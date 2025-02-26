using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToJsonAddin.Logging;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    /// <summary>
    /// 시트 분석 및 변환 가능 여부 판단 클래스
    /// </summary>
    public class SheetAnalyzer
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<SheetAnalyzer>();

        // 변환 가능한 시트를 판별하는 메서드
        public static List<Worksheet> GetConvertibleSheets(Workbook workbook)
        {
            var result = new List<Worksheet>();
            
            try
            {
                if (workbook == null) return result;
                
                // 워크북의 모든 시트를 순회
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (IsSheetConvertible(sheet))
                    {
                        result.Add(sheet);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "시트 분석 중 오류 발생");
                Debug.WriteLine($"시트 분석 중 오류 발생: {ex.Message}");
            }
            
            return result;
        }

        // 시트가 변환 가능한지 판별하는 메서드
        private static bool IsSheetConvertible(Worksheet sheet)
        {
            try
            {
                if (sheet == null) return false;
                
                // 시트 이름이 '!'로 시작하는지 확인
                string sheetName = sheet.Name;
                return sheetName != null && sheetName.StartsWith("!");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "시트 분석 중 오류 발생: {0}", sheet?.Name);
                Debug.WriteLine($"시트 분석 중 오류 발생: {ex.Message}");
                return false;
            }
        }
    }
}
