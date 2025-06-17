using ClosedXML.Excel;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 병합된 셀을 처리하는 서비스 구현
    /// </summary>
    public class MergedCellHandler : IMergedCellHandler
    {
        private readonly ISimpleLogger _logger;

        public MergedCellHandler(ISimpleLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public List<IXLRange> GetMergedRegionsInRow(IXLWorksheet worksheet, int rowNumber)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            var regions = new List<IXLRange>();
            
            foreach (var range in worksheet.MergedRanges)
            {
                if (range.FirstRow().RowNumber() <= rowNumber && 
                    range.LastRow().RowNumber() >= rowNumber)
                {
                    regions.Add(range);
                    _logger.Debug($"병합 영역 발견: {range.RangeAddress}, 행={rowNumber}");
                }
            }

            _logger.Debug($"행 {rowNumber}에서 {regions.Count}개의 병합 영역 발견");
            return regions;
        }

        /// <inheritdoc/>
        public (int startColumn, int endColumn) GetMergedCellRange(IXLCell cell, List<IXLRange> mergedRegions)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));
            
            if (mergedRegions == null)
                throw new ArgumentNullException(nameof(mergedRegions));

            var cellColumn = cell.Address.ColumnNumber;
            
            var containingRegion = mergedRegions.FirstOrDefault(region => region.Contains(cell));
            
            if (containingRegion != null)
            {
                var startColumn = containingRegion.FirstColumn().ColumnNumber();
                var endColumn = containingRegion.LastColumn().ColumnNumber();
                
                _logger.Debug($"셀 {cell.Address}는 병합 영역 {containingRegion.RangeAddress}에 포함됨: " +
                            $"열 범위 {startColumn}-{endColumn}");
                
                return (startColumn, endColumn);
            }

            return (cellColumn, cellColumn);
        }
    }
}