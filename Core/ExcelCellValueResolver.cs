using Microsoft.Extensions.Logging;
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
        private static readonly ILogger Logger = LoggerFactory.Create(builder => 
        {
            builder.SetMinimumLevel(LogLevel.Debug);
        }).CreateLogger(typeof(ExcelCellValueResolver));

        public static object GetCellValue(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
            {
                Logger.LogDebug("셀이 null이거나 비어 있습니다.");
                return null;
            }

            Logger.LogDebug("셀 값 추출: 행={Row}, 열={Column}, 타입={CellType}", 
                cell.Address.RowNumber, cell.Address.ColumnNumber, cell.DataType);

            // ClosedXML에서는 수식이 자동으로 평가됨
            if (cell.HasFormula)
            {
                Logger.LogDebug("수식 셀 처리: {Formula}", cell.FormulaA1);
            }

            switch (cell.DataType)
            {
                case XLDataType.Boolean:
                    var boolValue = cell.GetBoolean();
                    Logger.LogDebug("불리언 값: {Value}", boolValue);
                    return boolValue;

                case XLDataType.Number:
                    var numValue = cell.GetDouble();
                    Logger.LogDebug("숫자 값: {Value}", numValue);
                    return numValue;

                case XLDataType.Text:
                    var stringCellValue = cell.GetString();
                    Logger.LogDebug("문자열 값: {Value}", stringCellValue);
                    if (string.IsNullOrEmpty(stringCellValue))
                    {
                        Logger.LogDebug("빈 문자열 값");
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
                                Logger.LogDebug("문자열을 정수로 변환: {Value}", intValue);
                                return intValue;
                            }
                            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue))
                            {
                                Logger.LogDebug("문자열을 실수로 변환: {Value}", doubleValue);
                                return doubleValue;
                            }
                        }
                        
                        Logger.LogDebug("문자열 그대로 반환: {Value}", str);
                        return str;
                    }
                    catch (Exception ex)
                    {
                        Logger.LogWarning("문자열 변환 중 오류: {Error}", ex.Message);
                        return stringCellValue;
                    }

                case XLDataType.DateTime:
                    var dateValue = cell.GetDateTime();
                    Logger.LogDebug("날짜 값: {Value}", dateValue);
                    return dateValue;

                case XLDataType.TimeSpan:
                    var timeValue = cell.GetTimeSpan();
                    Logger.LogDebug("시간 값: {Value}", timeValue);
                    return timeValue;

                default:
                    // 그 외의 타입은 문자열로 반환
                    var value = cell.Value.ToString();
                    Logger.LogDebug("기타 타입 값: {Value}", value);
                    return value;
            }
        }
    }
}
