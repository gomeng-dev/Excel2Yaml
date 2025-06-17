using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Domain.Common;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Infrastructure.Logging;

namespace ExcelToYamlAddin.Domain.Entities
{
    /// <summary>
    /// Excel 시트의 스키마 노드를 나타내는 엔티티
    /// </summary>
    public class SchemeNode
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<SchemeNode>();
        private readonly List<SchemeNode> _children;
        private SchemeNode _parent;

        /// <summary>
        /// 노드 식별자
        /// </summary>
        public Guid Id { get; }

        /// <summary>
        /// 노드 키
        /// </summary>
        public string Key { get; }

        /// <summary>
        /// 노드 타입
        /// </summary>
        public SchemeNodeType NodeType { get; }

        /// <summary>
        /// 스키마 이름 (원본)
        /// </summary>
        public string SchemeName { get; }

        /// <summary>
        /// 셀 위치
        /// </summary>
        public CellPosition Position { get; }

        /// <summary>
        /// 부모 노드
        /// </summary>
        public SchemeNode Parent => _parent;

        /// <summary>
        /// 자식 노드 목록
        /// </summary>
        public IReadOnlyList<SchemeNode> Children => _children.AsReadOnly();

        /// <summary>
        /// 노드 깊이 (루트는 0)
        /// </summary>
        public int Depth => _parent?.Depth + 1 ?? 0;

        /// <summary>
        /// 루트 노드인지 여부
        /// </summary>
        public bool IsRoot => _parent == null;

        /// <summary>
        /// 컨테이너 노드인지 여부
        /// </summary>
        public bool IsContainer => NodeType.IsContainer;

        /// <summary>
        /// 키를 제공할 수 있는지 여부
        /// </summary>
        public bool IsKeyProvidable => NodeType == SchemeNodeType.Key || NodeType == SchemeNodeType.Property;

        /// <summary>
        /// 자식 노드 개수
        /// </summary>
        public int ChildCount => _children.Count;

        /// <summary>
        /// 노드가 값을 가질 수 있는지 여부
        /// </summary>
        public bool CanHaveValue => NodeType.CanHaveData;

        private SchemeNode(
            string key,
            SchemeNodeType nodeType,
            string schemeName,
            CellPosition position)
        {
            Id = Guid.NewGuid();
            Key = key ?? string.Empty;
            NodeType = nodeType ?? throw new ArgumentNullException(nameof(nodeType));
            SchemeName = schemeName ?? string.Empty;
            Position = position ?? throw new ArgumentNullException(nameof(position));
            _children = new List<SchemeNode>();
        }

        /// <summary>
        /// 스키마 노드 생성 팩토리 메서드
        /// </summary>
        public static SchemeNode Create(
            string schemeName,
            int row,
            int column)
        {
            if (string.IsNullOrWhiteSpace(schemeName))
                throw new ArgumentException(ErrorMessages.Schema.SchemeNameIsEmpty, nameof(schemeName));

            var position = new CellPosition(row, column);
            var nodeType = SchemeNodeType.FromSchemeName(schemeName);
            var key = ExtractKeyFromSchemeName(schemeName, nodeType, position);

            return new SchemeNode(key, nodeType, schemeName, position);
        }

        /// <summary>
        /// 루트 노드 생성
        /// </summary>
        public static SchemeNode CreateRoot()
        {
            return new SchemeNode(
                string.Empty,
                SchemeNodeType.Map,
                SchemeConstants.NodeTypes.Map,
                new CellPosition(SchemeConstants.Position.RootNodeRow, SchemeConstants.Position.RootNodeColumn));
        }

        /// <summary>
        /// 스키마 이름에서 키 추출
        /// </summary>
        private static string ExtractKeyFromSchemeName(
            string schemeName,
            SchemeNodeType nodeType,
            CellPosition position)
        {
            // 마커가 없으면 전체 이름이 키
            if (!schemeName.Contains(SchemeConstants.Markers.MarkerPrefix))
                return schemeName;

            // 마커로 분리
            var parts = schemeName.Split(
                new[] { SchemeConstants.Markers.MarkerPrefix[0] },
                StringSplitOptions.RemoveEmptyEntries);

            var key = parts.Length > 0 ? parts[0] : string.Empty;

            // 루트 컨테이너는 키가 없음 (YAML/JSON 표준)
            if (position.Column == SchemeConstants.Position.RootContainerColumn && nodeType.IsContainer)
                return string.Empty;

            return key;
        }

        /// <summary>
        /// 부모 노드 설정 (내부용)
        /// </summary>
        private void SetParent(SchemeNode parent)
        {
            _parent = parent;
        }

        /// <summary>
        /// 자식 노드 추가
        /// </summary>
        public void AddChild(SchemeNode child)
        {
            if (child == null)
                throw new ArgumentNullException(nameof(child));

            // 타입별 자식 노드 추가 검증
            var validationResult = ValidateChildAddition(child);
            if (!validationResult.IsValid)
                throw new InvalidOperationException(validationResult.ErrorMessage);

            child.SetParent(this);
            _children.Add(child);
        }

        /// <summary>
        /// 자식 노드 추가 검증
        /// </summary>
        private ValidationResult ValidateChildAddition(SchemeNode child)
        {
            // VALUE와 IGNORE 노드는 자식을 가질 수 없음
            if (NodeType == SchemeNodeType.Value)
                return ValidationResult.Failure(string.Format(ErrorMessages.Validation.CannotAddChildToValueNode, Key));

            if (NodeType == SchemeNodeType.Ignore)
                return ValidationResult.Failure(string.Format(ErrorMessages.Validation.CannotAddChildToIgnoreNode, Key));

            // KEY와 PROPERTY 노드는 KEY/PROPERTY 자식을 가질 수 없음
            if ((NodeType == SchemeNodeType.Key || NodeType == SchemeNodeType.Property) &&
                (child.NodeType == SchemeNodeType.Key || child.NodeType == SchemeNodeType.Property))
            {
                return ValidationResult.Failure(
                    string.Format(ErrorMessages.Validation.InvalidChildNodeType, NodeType, child.NodeType, Key, child.Key));
            }

            return ValidationResult.Success();
        }

        /// <summary>
        /// 이 노드와 모든 자식 노드를 평면화하여 반환
        /// </summary>
        public LinkedList<SchemeNode> Linear()
        {
            var result = new LinkedList<SchemeNode>();
            LinearizeRecursive(result);
            return result;
        }

        /// <summary>
        /// 재귀적으로 노드를 평면화
        /// </summary>
        private void LinearizeRecursive(LinkedList<SchemeNode> result)
        {
            result.AddLast(this);
            foreach (var child in _children)
            {
                child.LinearizeRecursive(result);
            }
        }

        /// <summary>
        /// 노드의 전체 경로 반환 (예: root/parent/child)
        /// </summary>
        public string GetFullPath()
        {
            var path = new List<string>();
            var current = this;

            while (current != null && !current.IsRoot)
            {
                if (!string.IsNullOrEmpty(current.Key))
                    path.Insert(0, current.Key);
                current = current._parent;
            }

            return path.Count > 0 ? string.Join("/", path) : string.Empty;
        }

        /// <summary>
        /// 특정 노드 타입의 자식 찾기
        /// </summary>
        public SchemeNode FindChildByType(SchemeNodeType nodeType)
        {
            return _children.FirstOrDefault(c => c.NodeType == nodeType);
        }

        /// <summary>
        /// 특정 키를 가진 자식 찾기
        /// </summary>
        public SchemeNode FindChildByKey(string key)
        {
            return _children.FirstOrDefault(c => c.Key == key);
        }

        /// <summary>
        /// 노드 복사
        /// </summary>
        public SchemeNode Clone()
        {
            var cloned = new SchemeNode(Key, NodeType, SchemeName, Position);
            
            foreach (var child in _children)
            {
                cloned.AddChild(child.Clone());
            }

            return cloned;
        }

        /// <summary>
        /// 노드 검증
        /// </summary>
        public NodeValidationResult Validate()
        {
            var errors = new List<string>();

            // 기본 검증
            if (NodeType == null)
                errors.Add(ErrorMessages.Validation.NodeTypeIsNull);

            if (Position == null)
                errors.Add(ErrorMessages.Validation.NodePositionIsNull);

            // 컨테이너 노드 검증
            if (IsContainer && _children.Count == 0)
                errors.Add(string.Format(ErrorMessages.Validation.ContainerNodeHasNoChildren, Key));

            // 자식 노드 검증
            foreach (var child in _children)
            {
                var childValidation = child.Validate();
                if (!childValidation.IsValid)
                    errors.AddRange(childValidation.Errors);
            }

            return new NodeValidationResult(errors.Count == 0, errors);
        }

        public override string ToString()
        {
            return $"{Key}:{NodeType.Code} [{Position}]";
        }

        /// <summary>
        /// 현재 행에서 노드의 값을 가져옵니다.
        /// </summary>
        public object GetValue(IXLRow row)
        {
            if (row == null)
            {
                Logger.Warning($"행이 null임: {Key}");
                return string.Empty;
            }

            var cell = row.Cell(Position.Column);
            if (cell == null || cell.IsEmpty())
            {
                Logger.Debug($"셀이 비어있음: 행={row.RowNumber()}, 열={Position.Column}");
                return string.Empty;
            }

            return ExcelCellValueResolver.GetCellValue(cell);
        }

        /// <summary>
        /// 현재 행에서 노드의 키를 가져옵니다.
        /// </summary>
        public string GetKey(IXLRow row)
        {
            // 기본 검증
            if (!IsKeyProvidable || row == null)
            {
                int rowNumber = row != null ? row.RowNumber() : -1;
                Logger.Debug($"키를 가져올 수 없음: 타입={NodeType}, 행={rowNumber}");
                return string.Empty;
            }

            // 1. $key 이름을 가진 노드는 항상 셀 값을 키로 사용
            if (NodeType == SchemeNodeType.Key && !string.IsNullOrEmpty(SchemeName) && SchemeName.Contains("$key"))
            {
                var cell = row.Cell(Position.Column);
                if (cell != null && !cell.IsEmpty())
                {
                    object cellValue = ExcelCellValueResolver.GetCellValue(cell);
                    string cellValueStr = cellValue != null ? cellValue.ToString() : string.Empty;
                    Logger.Debug($"$key 노드의 실제 셀 값: {cellValueStr}");
                    return cellValueStr;
                }
            }

            // 2. 키가 이미 있으면 그대로 사용
            if (!string.IsNullOrEmpty(Key))
            {
                return Key;
            }

            // 3. KEY 노드인 경우 셀 값 또는 자식 노드 값 사용
            if (NodeType == SchemeNodeType.Key)
            {
                // 값 노드가 있는 경우 해당 값을 사용
                var valueNode = Children.FirstOrDefault(c => c.NodeType == SchemeNodeType.Value);
                if (valueNode != null)
                {
                    object value = valueNode.GetValue(row);
                    string valueStr = value != null ? value.ToString() : string.Empty;
                    Logger.Debug($"KEY 노드의 값 노드 값: {valueStr}");
                    return valueStr;
                }

                // 값 노드가 없는 경우 직접 셀 값 사용
                var cell = row.Cell(Position.Column);
                if (cell != null && !cell.IsEmpty())
                {
                    object cellValue = ExcelCellValueResolver.GetCellValue(cell);
                    string cellValueStr = cellValue != null ? cellValue.ToString() : string.Empty;
                    Logger.Debug($"KEY 노드의 셀 값: {cellValueStr}");
                    return cellValueStr;
                }
            }

            // 4. 부모가 있고 부모가 키를 제공할 수 있는 경우 부모의 키 사용
            if (Parent != null && Parent.IsKeyProvidable)
            {
                // 부모의 키가 비어있는 경우 부모 셀의 값 사용
                if (string.IsNullOrEmpty(Parent.Key))
                {
                    var parentCell = row.Cell(Parent.Position.Column);
                    if (parentCell != null && !parentCell.IsEmpty())
                    {
                        object parentCellValue = ExcelCellValueResolver.GetCellValue(parentCell);
                        string parentCellValueStr = parentCellValue != null ? parentCellValue.ToString() : string.Empty;
                        Logger.Debug($"부모 노드의 셀 값: {parentCellValueStr}");
                        return parentCellValueStr;
                    }
                }

                // 부모의 키를 사용
                return Parent.GetKey(row);
            }

            // 5. 기본값
            Logger.Warning($"키를 결정할 수 없음: {this}, 부모={Parent}");
            return string.Empty;
        }

        public override bool Equals(object obj)
        {
            if (obj is SchemeNode other)
            {
                return Id == other.Id;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }
    }

    /// <summary>
    /// 검증 결과
    /// </summary>
    internal class ValidationResult
    {
        public bool IsValid { get; }
        public string ErrorMessage { get; }

        private ValidationResult(bool isValid, string errorMessage)
        {
            IsValid = isValid;
            ErrorMessage = errorMessage;
        }

        public static ValidationResult Success() => new ValidationResult(true, null);
        public static ValidationResult Failure(string message) => new ValidationResult(false, message);
    }

    /// <summary>
    /// 노드 검증 결과
    /// </summary>
    public class NodeValidationResult
    {
        public bool IsValid { get; }
        public IReadOnlyList<string> Errors { get; }

        public NodeValidationResult(bool isValid, IEnumerable<string> errors)
        {
            IsValid = isValid;
            Errors = (errors ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
        }
    }
}
