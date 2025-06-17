using System.Collections.Generic;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Infrastructure.Interfaces
{
    /// <summary>
    /// Excel 처리 서비스 인터페이스
    /// </summary>
    public interface IExcelService
    {
        /// <summary>
        /// Excel 파일을 엽니다.
        /// </summary>
        /// <param name="filePath">파일 경로</param>
        /// <returns>워크북</returns>
        Task<IXLWorkbook> OpenWorkbookAsync(string filePath);

        /// <summary>
        /// 워크북을 저장합니다.
        /// </summary>
        /// <param name="workbook">워크북</param>
        /// <param name="filePath">저장 경로</param>
        /// <returns>비동기 작업</returns>
        Task SaveWorkbookAsync(IXLWorkbook workbook, string filePath);

        /// <summary>
        /// 새 워크북을 생성합니다.
        /// </summary>
        /// <returns>새 워크북</returns>
        IXLWorkbook CreateWorkbook();

        /// <summary>
        /// 워크시트를 추가합니다.
        /// </summary>
        /// <param name="workbook">워크북</param>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>추가된 워크시트</returns>
        IXLWorksheet AddWorksheet(IXLWorkbook workbook, string sheetName);

        /// <summary>
        /// 자동 생성 대상 시트를 추출합니다.
        /// </summary>
        /// <param name="workbook">워크북</param>
        /// <returns>대상 시트 목록</returns>
        IEnumerable<IXLWorksheet> ExtractAutoGenTargetSheets(IXLWorkbook workbook);

        /// <summary>
        /// 셀 값을 가져옵니다.
        /// </summary>
        /// <param name="cell">셀</param>
        /// <returns>셀 값</returns>
        object GetCellValue(IXLCell cell);

        /// <summary>
        /// 셀 값을 설정합니다.
        /// </summary>
        /// <param name="cell">셀</param>
        /// <param name="value">값</param>
        void SetCellValue(IXLCell cell, object value);

        /// <summary>
        /// 셀 범위를 병합합니다.
        /// </summary>
        /// <param name="range">셀 범위</param>
        void MergeCells(IXLRange range);
    }
}