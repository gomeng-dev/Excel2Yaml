using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 스키마 끝 마커를 찾는 서비스 구현
    /// </summary>
    public class SchemeEndMarkerFinder : ISchemeEndMarkerFinder
    {
        private readonly ISimpleLogger _logger;

        public SchemeEndMarkerFinder(ISimpleLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public int FindSchemeEndRow(IXLWorksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            _logger.Debug($"스키마 끝 마커 검색 시작: 시트={worksheet.Name}");

            foreach (var row in worksheet.Rows())
            {
                var rowNumber = row.RowNumber();
                
                // 주석 행은 건너뛰기
                if (rowNumber == (SchemeConstants.RowNumbers.CommentRow + 1))
                {
                    _logger.Debug($"주석 행 건너뛰기: 행={rowNumber}");
                    continue;
                }

                if (ContainsEndMarker(row))
                {
                    _logger.Information($"스키마 끝 마커 발견: 행={rowNumber}");
                    return rowNumber;
                }
            }

            _logger.Warning("스키마 끝 마커를 찾을 수 없음");
            return SchemeConstants.RowNumbers.IllegalRow;
        }

        /// <inheritdoc/>
        public bool ContainsEndMarker(IXLRow row)
        {
            if (row == null)
                return false;

            var firstCell = row.Cell(SchemeConstants.Position.FirstColumnIndex);
            if (firstCell == null || firstCell.IsEmpty())
                return false;

            var cellValue = firstCell.GetString();
            var containsMarker = !string.IsNullOrEmpty(cellValue) &&
                               cellValue.Equals(SchemeConstants.Markers.SchemeEnd, StringComparison.OrdinalIgnoreCase);

            if (containsMarker)
            {
                _logger.Debug($"스키마 끝 마커 확인됨: 행={row.RowNumber()}, 값={cellValue}");
            }

            return containsMarker;
        }
    }
}