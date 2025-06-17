using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 노드 프로세서의 기본 추상 클래스
    /// </summary>
    public abstract class NodeProcessorBase : INodeProcessor
    {
        protected static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger(nameof(NodeProcessorBase));

        /// <summary>
        /// 이 프로세서가 처리할 수 있는 노드 타입
        /// </summary>
        protected abstract SchemeNodeType SupportedNodeType { get; }

        /// <summary>
        /// 노드 처리 가능 여부 확인
        /// </summary>
        public virtual bool CanProcess(SchemeNode node)
        {
            return node != null && node.NodeType.Equals(SupportedNodeType);
        }

        /// <summary>
        /// 노드 처리 (동기)
        /// </summary>
        public abstract object Process(SchemeNode node, GenerationContext context);

        /// <summary>
        /// 노드 처리 (비동기)
        /// </summary>
        public abstract Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default);

        /// <summary>
        /// 셀 값을 가져오고 포맷팅합니다.
        /// </summary>
        protected object GetFormattedCellValue(GenerationContext context, int column)
        {
            var cellValue = context.GetCellValue(column);
            if (cellValue == null)
                return null;

            // ExcelCellValueResolver를 사용하여 적절한 타입으로 변환
            var cell = context.Worksheet.Cell(context.CurrentRow, column);
            return ExcelCellValueResolver.GetCellValue(cell);
        }

        /// <summary>
        /// 빈 값인지 확인
        /// </summary>
        protected bool IsEmptyValue(object value)
        {
            if (value == null)
                return true;

            if (value is string str)
                return string.IsNullOrWhiteSpace(str);

            return false;
        }

        /// <summary>
        /// 빈 필드를 포함해야 하는지 확인
        /// </summary>
        protected bool ShouldIncludeEmpty(GenerationContext context, object value)
        {
            if (!IsEmptyValue(value))
                return true;

            return context.Options.ShowEmptyFields;
        }

        /// <summary>
        /// 현재 행이 완전히 비어있는지 확인
        /// </summary>
        protected bool IsRowEmpty(GenerationContext context, SchemeNode node)
        {
            // 노드의 시작 열부터 자식 노드들의 끝 열까지 확인
            int startCol = node.Position.Column;
            int endCol = GetNodeEndColumn(node);

            return context.IsCurrentRowEmpty(startCol, endCol);
        }

        /// <summary>
        /// 노드의 끝 열 위치를 계산
        /// </summary>
        protected int GetNodeEndColumn(SchemeNode node)
        {
            if (node.Children == null || node.Children.Count == 0)
                return node.Position.Column;

            int maxColumn = node.Position.Column;
            foreach (var child in node.Children)
            {
                maxColumn = Math.Max(maxColumn, GetNodeEndColumn(child));
            }
            return maxColumn;
        }

        /// <summary>
        /// 로그 메시지에 노드 정보 포함
        /// </summary>
        protected void LogNodeProcessing(string message, SchemeNode node, GenerationContext context)
        {
            Logger.Debug($"{message} - 노드: {node.Key}, 타입: {node.NodeType}, 행: {context.CurrentRow}, 열: {node.Position.Column}");
        }
    }
}