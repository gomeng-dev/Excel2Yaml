using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 속성(PROPERTY) 노드를 처리하는 프로세서
    /// </summary>
    public class PropertyNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Property;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            // 현재 행의 셀 값 가져오기
            var cellValue = GetFormattedCellValue(context, node.Position.Column);

            // 빈 값 처리
            if (!ShouldIncludeEmpty(context, cellValue))
            {
                return null;
            }

            // 키-값 쌍 반환
            return new KeyValuePair<string, object>(node.Key, cellValue);
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            LogNodeProcessing("속성 노드 처리 시작", node, context);

            try
            {
                // 현재 행의 셀 값 가져오기
                var cellValue = GetFormattedCellValue(context, node.Position.Column);

                // 빈 값 처리
                if (!ShouldIncludeEmpty(context, cellValue))
                {
                    LogNodeProcessing("빈 값 스킵", node, context);
                    return NodeProcessResult.Skip();
                }

                // 속성 데이터 생성
                var propertyData = new PropertyData
                {
                    Name = node.Key,
                    Value = cellValue
                };

                LogNodeProcessing($"속성 노드 처리 완료: 값={cellValue}", node, context);
                return await Task.FromResult(NodeProcessResult.Ok(propertyData, 1));
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, $"속성 노드 처리 중 오류: {node.Key}");
                return NodeProcessResult.Error($"속성 노드 처리 실패: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 속성 데이터
    /// </summary>
    public class PropertyData
    {
        public string Name { get; set; }
        public object Value { get; set; }
    }
}