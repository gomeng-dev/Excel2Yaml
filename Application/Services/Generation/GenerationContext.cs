using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation
{
    /// <summary>
    /// YAML 생성 과정에서 필요한 컨텍스트 정보를 관리합니다.
    /// 스택 기반 접근을 대체하여 더 명확한 상태 관리를 제공합니다.
    /// </summary>
    public class GenerationContext
    {
        private readonly IXLWorksheet _worksheet;
        private readonly Scheme _scheme;
        private readonly YamlGenerationOptions _options;
        private readonly Stack<NodeContext> _nodeStack;
        private int _currentRow;
        private readonly Dictionary<string, object> _metadata;

        public GenerationContext(
            IXLWorksheet worksheet,
            Scheme scheme,
            YamlGenerationOptions options)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _scheme = scheme ?? throw new ArgumentNullException(nameof(scheme));
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _nodeStack = new Stack<NodeContext>();
            _currentRow = scheme.ContentStartRow;
            _metadata = new Dictionary<string, object>();
        }

        /// <summary>
        /// 현재 워크시트
        /// </summary>
        public IXLWorksheet Worksheet => _worksheet;

        /// <summary>
        /// 스키마 정보
        /// </summary>
        public Scheme Scheme => _scheme;

        /// <summary>
        /// 생성 옵션
        /// </summary>
        public YamlGenerationOptions Options => _options;

        /// <summary>
        /// 현재 처리 중인 행 번호
        /// </summary>
        public int CurrentRow
        {
            get => _currentRow;
            set
            {
                if (value < 1 || value > _scheme.EndRow)
                {
                    throw new ArgumentOutOfRangeException(nameof(value), 
                        $"행 번호는 1과 {_scheme.EndRow} 사이여야 합니다.");
                }
                _currentRow = value;
            }
        }

        /// <summary>
        /// 스키마 종료 행
        /// </summary>
        public int EndRow => _scheme.EndRow;

        /// <summary>
        /// 현재 처리 중인 노드의 깊이
        /// </summary>
        public int CurrentDepth => _nodeStack.Count;

        /// <summary>
        /// 최대 깊이 초과 여부 확인
        /// </summary>
        public bool IsMaxDepthExceeded => CurrentDepth >= _options.MaxDepth;

        /// <summary>
        /// 현재 노드 컨텍스트
        /// </summary>
        public NodeContext CurrentNodeContext => _nodeStack.Count > 0 ? _nodeStack.Peek() : null;

        /// <summary>
        /// 부모 노드 컨텍스트
        /// </summary>
        public NodeContext ParentNodeContext
        {
            get
            {
                if (_nodeStack.Count < 2)
                    return null;

                var contexts = _nodeStack.ToArray();
                return contexts[1]; // 0은 현재, 1은 부모
            }
        }

        /// <summary>
        /// 노드 컨텍스트를 스택에 추가
        /// </summary>
        public void PushNodeContext(SchemeNode node, object data = null)
        {
            var context = new NodeContext
            {
                Node = node,
                Data = data,
                StartRow = _currentRow,
                Depth = CurrentDepth + 1
            };
            _nodeStack.Push(context);
        }

        /// <summary>
        /// 노드 컨텍스트를 스택에서 제거
        /// </summary>
        public NodeContext PopNodeContext()
        {
            if (_nodeStack.Count == 0)
            {
                throw new InvalidOperationException("노드 컨텍스트 스택이 비어있습니다.");
            }
            return _nodeStack.Pop();
        }

        /// <summary>
        /// 다음 행으로 이동
        /// </summary>
        public void MoveToNextRow()
        {
            if (_currentRow < _scheme.EndRow)
            {
                _currentRow++;
            }
        }

        /// <summary>
        /// 특정 행으로 이동
        /// </summary>
        public void MoveToRow(int row)
        {
            CurrentRow = row;
        }

        /// <summary>
        /// 메타데이터 추가
        /// </summary>
        public void AddMetadata(string key, object value)
        {
            _metadata[key] = value;
        }

        /// <summary>
        /// 메타데이터 가져오기
        /// </summary>
        public T GetMetadata<T>(string key, T defaultValue = default)
        {
            if (_metadata.TryGetValue(key, out var value) && value is T typedValue)
            {
                return typedValue;
            }
            return defaultValue;
        }

        /// <summary>
        /// 현재 행이 데이터 범위 내에 있는지 확인
        /// </summary>
        public bool IsInDataRange => _currentRow >= _scheme.ContentStartRow && _currentRow <= _scheme.EndRow;

        /// <summary>
        /// 현재 셀 값 가져오기
        /// </summary>
        public object GetCellValue(int column)
        {
            if (!IsInDataRange)
                return null;

            var cell = _worksheet.Cell(_currentRow, column);
            return cell?.Value;
        }

        /// <summary>
        /// 현재 행의 모든 값이 비어있는지 확인
        /// </summary>
        public bool IsCurrentRowEmpty(int startColumn, int endColumn)
        {
            for (int col = startColumn; col <= endColumn; col++)
            {
                var value = GetCellValue(col);
                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                {
                    return false;
                }
            }
            return true;
        }
    }

    /// <summary>
    /// 노드 처리 컨텍스트
    /// </summary>
    public class NodeContext
    {
        /// <summary>
        /// 처리 중인 스키마 노드
        /// </summary>
        public SchemeNode Node { get; set; }

        /// <summary>
        /// 노드에 연결된 데이터
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// 노드 처리 시작 행
        /// </summary>
        public int StartRow { get; set; }

        /// <summary>
        /// 노드의 깊이
        /// </summary>
        public int Depth { get; set; }

        /// <summary>
        /// 추가 속성들
        /// </summary>
        public Dictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();
    }
}