using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using ClosedXML.Excel;
using System;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 스키마 노드를 생성하는 빌더 구현
    /// </summary>
    public class SchemeNodeBuilder : ISchemeNodeBuilder
    {
        private readonly ISimpleLogger _logger;

        public SchemeNodeBuilder(ISimpleLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public SchemeNode BuildFromCell(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return null;

            var value = cell.GetString();
            var row = cell.Address.RowNumber;
            var column = cell.Address.ColumnNumber;

            if (ShouldIgnore(value))
            {
                _logger.Debug($"셀 값 무시됨: 행={row}, 열={column}, 값={value}");
                return null;
            }

            return Build(value, row, column);
        }

        /// <inheritdoc/>
        public SchemeNode Build(string value, int row, int column)
        {
            if (string.IsNullOrEmpty(value))
                throw new ArgumentException("노드 값은 비어있을 수 없습니다.", nameof(value));

            try
            {
                var node = SchemeNode.Create(value, row, column);
                _logger.Debug($"스키마 노드 생성됨: 키={node.Key}, 타입={node.NodeType}, 위치=({row},{column})");
                return node;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"스키마 노드 생성 실패: 값={value}, 위치=({row},{column})");
                throw;
            }
        }

        /// <inheritdoc/>
        public bool ShouldIgnore(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return true;

            return value.Equals(SchemeConstants.Markers.Ignore, StringComparison.OrdinalIgnoreCase);
        }
    }
}