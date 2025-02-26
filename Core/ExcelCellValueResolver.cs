using ExcelToJsonAddin.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public static class ExcelCellValueResolver
    {
        // 단순화된 로깅 방식으로 변경
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger("ExcelCellValueResolver");

        public static object GetCellValue(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
            {
                Debug.WriteLine($"[ExcelCellValueResolver] 셀이 null이거나 비어 있습니다.");
                return null;
            }

            try
            {
                Debug.WriteLine($"[ExcelCellValueResolver] 셀 값 추출: 행={cell.Address.RowNumber}, 열={cell.Address.ColumnNumber}, 타입={cell.DataType}");

                // ClosedXML에서는 수식이 자동으로 평가됨
                if (cell.HasFormula)
                {
                    Debug.WriteLine($"[ExcelCellValueResolver] 수식 셀 처리: {cell.FormulaA1}");
                }

                switch (cell.DataType)
                {
                    case XLDataType.Boolean:
                        try
                        {
                            var boolValue = cell.GetBoolean();
                            Debug.WriteLine($"[ExcelCellValueResolver] 불리언 값: {boolValue}");
                            return boolValue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ExcelCellValueResolver] 불리언 값 추출 오류: {ex.Message}");
                            return null;
                        }

                    case XLDataType.Number:
                        try
                        {
                            var numValue = cell.GetDouble();
                            Debug.WriteLine($"[ExcelCellValueResolver] 숫자 값: {numValue}");
                            return numValue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ExcelCellValueResolver] 숫자 값 추출 오류: {ex.Message}");
                            return null;
                        }

                    case XLDataType.Text:
                        try
                        {
                            var stringCellValue = cell.GetString();
                            Debug.WriteLine($"[ExcelCellValueResolver] 문자열 값: {stringCellValue}");
                            if (string.IsNullOrEmpty(stringCellValue))
                            {
                                Debug.WriteLine($"[ExcelCellValueResolver] 빈 문자열 값");
                                return null;
                            }
                            
                            try
                            {
                                // Trim() 메서드 제거하여 공백 유지
                                var str = stringCellValue;
                                
                                // 문자열이 숫자만 포함하는 경우에만 숫자로 변환 시도
                                // 공백을 포함하는 문자열을 숫자로 변환하지 않도록 함
                                if (!string.IsNullOrWhiteSpace(str) && !str.Contains(" ") && 
                                    !str.StartsWith(" ") && !str.EndsWith(" "))
                                {
                                    if (int.TryParse(str, out int intValue))
                                    {
                                        Debug.WriteLine($"[ExcelCellValueResolver] 문자열을 정수로 변환: {intValue}");
                                        return intValue;
                                    }
                                    if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue))
                                    {
                                        Debug.WriteLine($"[ExcelCellValueResolver] 문자열을 실수로 변환: {doubleValue}");
                                        return doubleValue;
                                    }
                                }
                                
                                Debug.WriteLine($"[ExcelCellValueResolver] 문자열 그대로 반환: {str}");
                                return str;
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[ExcelCellValueResolver] 문자열 변환 중 오류: {ex.Message}");
                                return stringCellValue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ExcelCellValueResolver] 문자열 값 추출 오류: {ex.Message}");
                            return null;
                        }

                    case XLDataType.DateTime:
                        try
                        {
                            var dateValue = cell.GetDateTime();
                            Debug.WriteLine($"[ExcelCellValueResolver] 날짜 값: {dateValue}");
                            return dateValue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ExcelCellValueResolver] 날짜 값 추출 오류: {ex.Message}");
                            return null;
                        }

                    default:
                        // 그 외의 타입은 문자열로 반환
                        try
                        {
                            var value = cell.Value.ToString();
                            Debug.WriteLine($"[ExcelCellValueResolver] 기타 타입 값: {value}");
                            return value;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[ExcelCellValueResolver] 기타 타입 값 추출 오류: {ex.Message}");
                            return null;
                        }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelCellValueResolver] 셀 값 추출 중 예외 발생: {ex.Message}");
                return null;
            }
        }
    }
}
