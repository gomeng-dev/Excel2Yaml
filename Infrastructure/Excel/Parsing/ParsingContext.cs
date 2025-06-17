using ClosedXML.Excel;
using System;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 스키마 파싱 컨텍스트 정보를 담는 클래스
    /// </summary>
    public class ParsingContext
    {
        /// <summary>
        /// 파싱 중인 워크시트
        /// </summary>
        public IXLWorksheet Worksheet { get; }

        /// <summary>
        /// 스키마 시작 행
        /// </summary>
        public IXLRow SchemeStartRow { get; }

        /// <summary>
        /// 스키마 끝 행 번호
        /// </summary>
        public int SchemeEndRowNumber { get; }

        /// <summary>
        /// 데이터 시작 행 번호
        /// </summary>
        public int DataStartRowNumber { get; }

        /// <summary>
        /// 첫 번째 셀 열 번호
        /// </summary>
        public int FirstCellNumber { get; }

        /// <summary>
        /// 마지막 셀 열 번호
        /// </summary>
        public int LastCellNumber { get; }

        public ParsingContext(
            IXLWorksheet worksheet,
            IXLRow schemeStartRow,
            int schemeEndRowNumber,
            int dataStartRowNumber,
            int firstCellNumber,
            int lastCellNumber)
        {
            Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            SchemeStartRow = schemeStartRow ?? throw new ArgumentNullException(nameof(schemeStartRow));
            SchemeEndRowNumber = schemeEndRowNumber;
            DataStartRowNumber = dataStartRowNumber;
            FirstCellNumber = firstCellNumber;
            LastCellNumber = lastCellNumber;

            Validate();
        }

        private void Validate()
        {
            if (SchemeEndRowNumber <= 0)
                throw new ArgumentException("스키마 끝 행 번호는 0보다 커야 합니다.", nameof(SchemeEndRowNumber));

            if (DataStartRowNumber <= SchemeEndRowNumber)
                throw new ArgumentException("데이터 시작 행은 스키마 끝 행 다음에 있어야 합니다.", nameof(DataStartRowNumber));

            if (FirstCellNumber <= 0)
                throw new ArgumentException("첫 번째 셀 번호는 0보다 커야 합니다.", nameof(FirstCellNumber));

            if (LastCellNumber < FirstCellNumber)
                throw new ArgumentException("마지막 셀 번호는 첫 번째 셀 번호보다 크거나 같아야 합니다.", nameof(LastCellNumber));
        }

        /// <summary>
        /// 데이터 끝 행 번호를 가져옵니다.
        /// </summary>
        public int GetDataEndRowNumber()
        {
            return Worksheet.LastRowUsed()?.RowNumber() ?? DataStartRowNumber;
        }

        /// <summary>
        /// 파싱 컨텍스트의 요약 정보를 반환합니다.
        /// </summary>
        public override string ToString()
        {
            return $"ParsingContext: Sheet={Worksheet.Name}, " +
                   $"SchemeRows={SchemeStartRow.RowNumber()}-{SchemeEndRowNumber}, " +
                   $"DataStartRow={DataStartRowNumber}, " +
                   $"Columns={FirstCellNumber}-{LastCellNumber}";
        }
    }
}