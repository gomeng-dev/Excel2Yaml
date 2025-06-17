using System;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using ExcelToYamlAddin.Infrastructure.Logging;

namespace ExcelToYamlAddin.Core
{
    /// <summary>
    /// Excel 시트를 HTML로 내보내는 디버깅용 클래스
    /// </summary>
    public class ExcelToHtmlExporter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ExcelToHtmlExporter>();

        /// <summary>
        /// Excel 워크시트를 HTML 파일로 내보냅니다.
        /// </summary>
        public static void ExportToHtml(Worksheet worksheet, string outputPath)
        {
            try
            {
                Logger.Information($"Excel 시트를 HTML로 내보내기 시작: {worksheet.Name}");

                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html>");
                html.AppendLine("<head>");
                html.AppendLine("<meta charset='UTF-8'>");
                html.AppendLine($"<title>{worksheet.Name}</title>");
                html.AppendLine("<style>");
                html.AppendLine("table { border-collapse: collapse; margin: 20px; }");
                html.AppendLine("td, th { border: 1px solid #999; padding: 8px; min-width: 80px; }");
                html.AppendLine("th { background-color: #f2f2f2; font-weight: bold; }");
                html.AppendLine(".merged { background-color: #e8f4fc; }");
                html.AppendLine(".scheme-end { background-color: #ff0000; color: white; text-align: center; }");
                html.AppendLine(".array-marker { background-color: #00CC00; }");
                html.AppendLine(".object-marker { background-color: #CCFFCC; }");
                html.AppendLine(".empty { background-color: #f9f9f9; }");
                html.AppendLine(".row-header { background-color: #ddd; font-weight: bold; width: 50px; }");
                html.AppendLine(".col-header { background-color: #ddd; font-weight: bold; height: 30px; text-align: center; }");
                html.AppendLine("</style>");
                html.AppendLine("</head>");
                html.AppendLine("<body>");
                html.AppendLine($"<h1>시트: {worksheet.Name}</h1>");
                
                // 사용된 범위 가져오기
                Range usedRange = worksheet.UsedRange;
                if (usedRange == null)
                {
                    html.AppendLine("<p>시트에 데이터가 없습니다.</p>");
                }
                else
                {
                    int rowCount = usedRange.Rows.Count;
                    int colCount = usedRange.Columns.Count;
                    int startRow = usedRange.Row;
                    int startCol = usedRange.Column;

                    html.AppendLine("<table>");
                    
                    // 열 헤더 추가 (A, B, C...)
                    html.AppendLine("<tr>");
                    html.AppendLine("<td class='col-header'></td>"); // 빈 셀
                    for (int c = 0; c < colCount; c++)
                    {
                        string colLetter = GetColumnLetter(startCol + c);
                        html.AppendLine($"<td class='col-header'>{colLetter}</td>");
                    }
                    html.AppendLine("</tr>");

                    // 데이터 행 추가
                    for (int r = 0; r < rowCount; r++)
                    {
                        html.AppendLine("<tr>");
                        
                        // 행 번호 추가
                        html.AppendLine($"<td class='row-header'>{startRow + r}</td>");
                        
                        for (int c = 0; c < colCount; c++)
                        {
                            Range cell = (Range)usedRange.Cells[r + 1, c + 1];
                            string value = cell.Value2?.ToString() ?? "";
                            string cssClass = GetCellClass(value, cell);
                            
                            // 병합된 셀 확인
                            Range mergeArea = cell.MergeArea;
                            int colspan = mergeArea.Columns.Count;
                            int rowspan = mergeArea.Rows.Count;
                            
                            // 병합된 셀의 첫 번째 셀만 표시
                            if (cell.Row == mergeArea.Row && cell.Column == mergeArea.Column)
                            {
                                if (colspan > 1 || rowspan > 1)
                                {
                                    cssClass += " merged";
                                    html.Append($"<td class='{cssClass}'");
                                    if (colspan > 1) html.Append($" colspan='{colspan}'");
                                    if (rowspan > 1) html.Append($" rowspan='{rowspan}'");
                                    html.AppendLine($">{HtmlEncode(value)}</td>");
                                }
                                else
                                {
                                    html.AppendLine($"<td class='{cssClass}'>{HtmlEncode(value)}</td>");
                                }
                                
                                // 병합된 열 건너뛰기
                                if (colspan > 1) c += colspan - 1;
                            }
                            else if (IsCellInMergeArea(cell, mergeArea))
                            {
                                // 병합된 영역의 다른 셀들은 건너뛰기
                                continue;
                            }
                            else
                            {
                                html.AppendLine($"<td class='{cssClass}'>{HtmlEncode(value)}</td>");
                            }
                        }
                        html.AppendLine("</tr>");
                    }
                    
                    html.AppendLine("</table>");
                }
                
                html.AppendLine("<div style='margin: 20px;'>");
                html.AppendLine("<h3>범례:</h3>");
                html.AppendLine("<p><span style='background-color: #00CC00; padding: 5px;'>$[]</span> - 배열 마커</p>");
                html.AppendLine("<p><span style='background-color: #CCFFCC; padding: 5px;'>${}</span> - 객체 마커</p>");
                html.AppendLine("<p><span style='background-color: #ff0000; color: white; padding: 5px;'>$scheme_end</span> - 스키마 종료</p>");
                html.AppendLine("<p><span style='background-color: #e8f4fc; padding: 5px;'>병합된 셀</span></p>");
                html.AppendLine("</div>");
                
                html.AppendLine("</body>");
                html.AppendLine("</html>");

                // HTML 파일 저장
                File.WriteAllText(outputPath, html.ToString(), Encoding.UTF8);
                Logger.Information($"HTML 파일 생성 완료: {outputPath}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Excel을 HTML로 내보내는 중 오류 발생");
                throw;
            }
        }

        /// <summary>
        /// 열 번호를 문자로 변환 (1 -> A, 2 -> B, 27 -> AA)
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// 셀의 CSS 클래스를 결정합니다.
        /// </summary>
        private static string GetCellClass(string value, Range cell)
        {
            if (string.IsNullOrEmpty(value))
                return "empty";
            
            if (value.Contains("$scheme_end"))
                return "scheme-end";
            
            if (value.Contains("$[]"))
                return "array-marker";
            
            if (value.Contains("${}"))
                return "object-marker";
            
            // 배경색 확인
            try
            {
                var interior = cell.Interior;
                if (interior.Color != null)
                {
                    int colorValue = Convert.ToInt32(interior.Color);
                    // Excel의 색상 값을 RGB로 변환
                    int red = colorValue & 0xFF;
                    int green = (colorValue >> 8) & 0xFF;
                    int blue = (colorValue >> 16) & 0xFF;
                    
                    // 빨간색 배경 확인
                    if (red > 200 && green < 100 && blue < 100)
                        return "scheme-end";
                    
                    // 초록색 배경 확인
                    if (green > 200 && red < 100 && blue < 100)
                        return "array-marker";
                    
                    // 연한 초록색 배경 확인
                    if (green > 200 && red > 200 && blue < 100)
                        return "object-marker";
                }
            }
            catch { }
            
            return "";
        }

        /// <summary>
        /// 셀이 병합 영역에 속하는지 확인합니다.
        /// </summary>
        private static bool IsCellInMergeArea(Range cell, Range mergeArea)
        {
            return cell.Row >= mergeArea.Row && 
                   cell.Row < mergeArea.Row + mergeArea.Rows.Count &&
                   cell.Column >= mergeArea.Column && 
                   cell.Column < mergeArea.Column + mergeArea.Columns.Count &&
                   !(cell.Row == mergeArea.Row && cell.Column == mergeArea.Column);
        }

        /// <summary>
        /// 문자열을 HTML 인코딩합니다.
        /// </summary>
        private static string HtmlEncode(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }
    }
}