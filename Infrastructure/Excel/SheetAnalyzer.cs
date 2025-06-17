using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelToYamlAddin.Infrastructure.Excel
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
                Logger.Error(ex, ErrorMessages.Conversion.SheetAnalysisError);
                Debug.WriteLine($"{ErrorMessages.Conversion.SheetAnalysisError}: {ex.Message}");
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
                return sheetName != null && sheetName.StartsWith(SchemeConstants.Sheet.ConversionPrefix);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, ErrorMessages.Conversion.SheetAnalysisErrorWithName, sheet?.Name);
                Debug.WriteLine($"{ErrorMessages.Conversion.SheetAnalysisError}: {ex.Message}");
                return false;
            }
        }
    }
}
