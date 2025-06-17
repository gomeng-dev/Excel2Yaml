using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 동적 키-값 쌍을 처리하는 프로세서
    /// $key와 $value 노드를 함께 처리합니다.
    /// </summary>
    public class KeyValueNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Key;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            if (node.NodeType == SchemeNodeType.Key)
            {
                // KEY 노드 처리 - VALUE 노드 찾기
                var valueNode = node.Children?.FirstOrDefault(c => c.NodeType == SchemeNodeType.Value) ??
                               node.Parent?.Children?.FirstOrDefault(c => c.NodeType == SchemeNodeType.Value);
                
                if (valueNode != null)
                {
                    var key = GetFormattedCellValue(context, node.Position.Column);
                    var value = GetFormattedCellValue(context, valueNode.Position.Column);
                    
                    if (!ShouldIncludeEmpty(context, value))
                        return null;
                    
                    return new KeyValuePair<string, object>(key?.ToString() ?? "", value);
                }
            }
            else if (node.NodeType == SchemeNodeType.Value)
            {
                // 단독 VALUE 노드 처리
                var value = GetFormattedCellValue(context, node.Position.Column);
                return value;
            }
            
            return null;
        }

        public override bool CanProcess(SchemeNode node)
        {
            // KEY 또는 VALUE 노드인 경우 처리 가능
            return node != null && 
                   (node.NodeType == SchemeNodeType.Key || node.NodeType == SchemeNodeType.Value);
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            LogNodeProcessing("키-값 노드 처리 시작", node, context);

            try
            {
                var parent = node.Parent;
                if (parent == null || parent.Children == null)
                {
                    return NodeProcessResult.Error("부모 노드를 찾을 수 없습니다.");
                }

                // VALUE 노드 찾기
                var valueNode = parent.Children.FirstOrDefault(child => child.NodeType.Equals(SchemeNodeType.Value));
                if (valueNode == null)
                {
                    return NodeProcessResult.Error("VALUE 노드를 찾을 수 없습니다.");
                }

                // 키와 값 가져오기
                var key = GetFormattedCellValue(context, node.Position.Column);
                var value = GetFormattedCellValue(context, valueNode.Position.Column);

                // 빈 키는 스킵
                if (IsEmptyValue(key))
                {
                    LogNodeProcessing("빈 키 스킵", node, context);
                    return NodeProcessResult.Skip();
                }

                // 빈 값 처리
                if (!ShouldIncludeEmpty(context, value))
                {
                    LogNodeProcessing($"빈 값 스킵: 키={key}", node, context);
                    return NodeProcessResult.Skip();
                }

                var kvData = new KeyValueData
                {
                    Key = key.ToString(),
                    Value = value
                };

                LogNodeProcessing($"키-값 노드 처리 완료: 키={key}, 값={value}", node, context);
                return await Task.FromResult(NodeProcessResult.Ok(kvData, 1));
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, "키-값 노드 처리 중 오류");
                return NodeProcessResult.Error($"키-값 노드 처리 실패: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// VALUE 노드를 처리하는 프로세서
    /// KEY 노드와 함께 처리되므로 별도 처리는 하지 않습니다.
    /// </summary>
    public class ValueNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Value;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            // VALUE 노드는 KEY 노드에서 함께 처리되므로 null 반환
            return null;
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            // VALUE 노드는 KEY 노드에서 함께 처리되므로 스킵
            LogNodeProcessing("VALUE 노드 스킵 (KEY 노드에서 처리됨)", node, context);
            return await Task.FromResult(NodeProcessResult.Skip());
        }
    }

    /// <summary>
    /// 키-값 데이터
    /// </summary>
    public class KeyValueData
    {
        public string Key { get; set; }
        public object Value { get; set; }
    }
}