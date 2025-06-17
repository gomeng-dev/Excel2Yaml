using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.FileSystem;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 배열(ARRAY) 노드를 처리하는 프로세서
    /// </summary>
    public class ArrayNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Array;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            // 배열 객체 생성
            return OrderedYamlFactory.CreateArray();
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            LogNodeProcessing("배열 노드 처리 시작", node, context);

            try
            {
                var arrayData = new ArrayData { Name = node.Key };
                var items = new List<object>();
                int rowsConsumed = 0;

                // 배열의 각 항목 처리
                while (context.CurrentRow <= context.EndRow)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 현재 행이 비어있는지 확인
                    if (IsRowEmpty(context, node))
                    {
                        // 빈 행은 배열의 끝을 의미
                        LogNodeProcessing("빈 행 발견 - 배열 종료", node, context);
                        break;
                    }

                    // 배열 항목 처리
                    var itemData = await ProcessArrayItem(node, context, traverser, cancellationToken);
                    
                    if (itemData != null)
                    {
                        items.Add(itemData);
                    }

                    rowsConsumed++;
                    
                    // 다음 행으로 이동
                    if (context.CurrentRow < context.EndRow)
                    {
                        context.MoveToNextRow();
                    }
                    else
                    {
                        break;
                    }
                }

                // 빈 배열 처리
                if (items.Count == 0 && !context.Options.ShowEmptyFields)
                {
                    LogNodeProcessing("빈 배열 스킵", node, context);
                    return NodeProcessResult.Skip();
                }

                arrayData.Items = items;
                LogNodeProcessing($"배열 노드 처리 완료: 항목 수={items.Count}, 소비된 행={rowsConsumed}", node, context);
                
                return NodeProcessResult.Ok(arrayData, rowsConsumed);
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, $"배열 노드 처리 중 오류: {node.Key}");
                return NodeProcessResult.Error($"배열 노드 처리 실패: {ex.Message}");
            }
        }

        private async Task<object> ProcessArrayItem(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken)
        {
            // 배열의 자식 노드가 하나인 경우 (단순 배열)
            if (node.Children != null && node.Children.Count == 1)
            {
                var child = node.Children[0];
                
                // PROPERTY 노드인 경우 값만 추출
                if (child.NodeType.Equals(SchemeNodeType.Property))
                {
                    var value = GetFormattedCellValue(context, child.Position.Column);
                    if (ShouldIncludeEmpty(context, value))
                    {
                        return value;
                    }
                    return null;
                }
                else
                {
                    // 다른 타입의 노드는 정상 처리
                    var result = await traverser.TraverseAsync(child, context, cancellationToken);
                    return result.Success ? result.Data : null;
                }
            }
            // 배열의 자식 노드가 여러 개인 경우 (복잡한 배열)
            else if (node.Children != null && node.Children.Count > 1)
            {
                var itemProperties = new Dictionary<string, object>();
                
                foreach (var child in node.Children)
                {
                    var result = await traverser.TraverseAsync(child, context, cancellationToken);
                    if (result.Success && result.Data != null)
                    {
                        AddToItem(itemProperties, result.Data);
                    }
                }

                return itemProperties.Count > 0 ? itemProperties : null;
            }

            return null;
        }

        private void AddToItem(Dictionary<string, object> item, object data)
        {
            switch (data)
            {
                case PropertyData property:
                    item[property.Name] = property.Value;
                    break;
                case MapData map:
                    if (!string.IsNullOrEmpty(map.Name))
                    {
                        item[map.Name] = map.Properties;
                    }
                    break;
                case ArrayData array:
                    item[array.Name] = array.Items;
                    break;
            }
        }
    }

    /// <summary>
    /// 배열 데이터
    /// </summary>
    public class ArrayData
    {
        public string Name { get; set; }
        public List<object> Items { get; set; } = new List<object>();
    }
}