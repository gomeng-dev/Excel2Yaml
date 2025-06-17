using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Domain.Common;

namespace ExcelToYamlAddin.Domain.Entities
{
    /// <summary>
    /// Excel 시트의 스키마 구조를 나타내는 엔티티
    /// </summary>
    public class Scheme : IEnumerable<SchemeNode>
    {
        private readonly LinkedList<SchemeNode> _linearNodes;

        /// <summary>
        /// 루트 노드
        /// </summary>
        public SchemeNode Root { get; }

        /// <summary>
        /// 시트 이름
        /// </summary>
        public string SheetName { get; }

        /// <summary>
        /// 컨텐츠 시작 행 번호 (1-based)
        /// </summary>
        public int ContentStartRow { get; }

        /// <summary>
        /// 스키마 종료 행 번호 (1-based)
        /// </summary>
        public int EndRow { get; }

        /// <summary>
        /// 스키마 시작 위치
        /// </summary>
        public CellPosition StartPosition { get; }

        /// <summary>
        /// 스키마 종료 위치
        /// </summary>
        public CellPosition EndPosition { get; }

        /// <summary>
        /// 스키마의 노드 개수
        /// </summary>
        public int NodeCount => _linearNodes?.Count ?? 0;

        /// <summary>
        /// 스키마가 유효한지 여부
        /// </summary>
        public bool IsValid => Root != null && EndRow > 0 && ContentStartRow > 0;

        /// <summary>
        /// 데이터 행 개수 (스키마 종료 행 - 컨텐츠 시작 행 + 1)
        /// </summary>
        public int DataRowCount => Math.Max(0, EndRow - ContentStartRow + 1);

        /// <summary>
        /// 스키마 생성 시간
        /// </summary>
        public DateTime CreatedAt { get; }

        /// <summary>
        /// 스키마 메타데이터
        /// </summary>
        public SchemeMetadata Metadata { get; }

        private Scheme()
        {
            _linearNodes = new LinkedList<SchemeNode>();
            CreatedAt = DateTime.UtcNow;
        }

        private Scheme(
            string sheetName,
            SchemeNode root,
            int contentStartRow,
            int endRow,
            CellPosition startPosition = null,
            CellPosition endPosition = null,
            SchemeMetadata metadata = null) : this()
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException(ErrorMessages.Validation.InvalidSheetName, nameof(sheetName));

            if (root == null)
                throw new ArgumentNullException(nameof(root), ErrorMessages.Schema.RootNodeIsNull);

            if (contentStartRow < 1)
                throw new ArgumentException(ErrorMessages.Schema.InvalidContentStartRow, nameof(contentStartRow));

            if (endRow < contentStartRow)
                throw new ArgumentException(ErrorMessages.Schema.EndRowLessThanStartRow, nameof(endRow));

            SheetName = sheetName;
            Root = root;
            ContentStartRow = contentStartRow;
            EndRow = endRow;
            StartPosition = startPosition ?? new CellPosition(SchemeConstants.Sheet.SchemaStartRow, 1);
            EndPosition = endPosition ?? new CellPosition(endRow, 1);
            Metadata = metadata ?? SchemeMetadata.Default();

            // 선형 노드 목록 생성
            _linearNodes = root.Linear() ?? new LinkedList<SchemeNode>();
        }

        /// <summary>
        /// 스키마 생성 팩토리 메서드
        /// </summary>
        public static Scheme Create(
            string sheetName,
            SchemeNode root,
            int contentStartRow,
            int endRow,
            CellPosition startPosition = null,
            CellPosition endPosition = null,
            SchemeMetadata metadata = null)
        {
            return new Scheme(
                sheetName,
                root,
                contentStartRow,
                endRow,
                startPosition,
                endPosition,
                metadata);
        }

        /// <summary>
        /// 빈 스키마 생성
        /// </summary>
        public static Scheme Empty(string sheetName)
        {
            var emptyRoot = SchemeNode.CreateRoot();
            return new Scheme(
                sheetName,
                emptyRoot,
                SchemeConstants.Sheet.DataStartRow,
                SchemeConstants.Sheet.DataStartRow);
        }

        /// <summary>
        /// 모든 스키마 노드를 선형 순서로 반환
        /// </summary>
        public IEnumerable<SchemeNode> GetLinearNodes()
        {
            return _linearNodes.ToList();
        }

        /// <summary>
        /// 특정 타입의 노드만 필터링
        /// </summary>
        public IEnumerable<SchemeNode> GetNodesByType(SchemeNodeType nodeType)
        {
            return _linearNodes.Where(node => node.NodeType == nodeType);
        }

        /// <summary>
        /// 특정 깊이의 노드만 필터링
        /// </summary>
        public IEnumerable<SchemeNode> GetNodesByDepth(int depth)
        {
            return _linearNodes.Where(node => node.Depth == depth);
        }

        /// <summary>
        /// 특정 경로의 노드 찾기
        /// </summary>
        public SchemeNode FindNodeByPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;

            var parts = path.Split(new[] { '/', '.' }, StringSplitOptions.RemoveEmptyEntries);
            var current = Root;

            foreach (var part in parts)
            {
                if (current == null)
                    return null;

                current = current.Children.FirstOrDefault(c => c.Key == part);
            }

            return current;
        }

        /// <summary>
        /// 스키마 검증
        /// </summary>
        public SchemeValidationResult Validate()
        {
            var errors = new List<string>();

            // 기본 검증
            if (!IsValid)
            {
                errors.Add(ErrorMessages.Schema.SchemaIsInvalid);
            }

            // 루트 노드 검증
            if (Root == null)
            {
                errors.Add(ErrorMessages.Schema.RootNodeIsNull);
            }

            // 노드 구조 검증
            if (_linearNodes.Count == 0)
            {
                errors.Add(ErrorMessages.Schema.NoNodesInSchema);
            }

            // 순환 참조 검증
            if (HasCircularReference())
            {
                errors.Add(ErrorMessages.Schema.CircularReferenceFound);
            }

            // 중복 키 검증
            var duplicateKeys = GetDuplicateKeys();
            if (duplicateKeys.Any())
            {
                errors.Add(string.Format(ErrorMessages.Schema.DuplicateKeysFound, string.Join(", ", duplicateKeys)));
            }

            return new SchemeValidationResult(errors.Count == 0, errors);
        }

        /// <summary>
        /// 순환 참조 확인
        /// </summary>
        private bool HasCircularReference()
        {
            var visited = new HashSet<SchemeNode>();
            var recursionStack = new HashSet<SchemeNode>();

            return HasCircularReferenceHelper(Root, visited, recursionStack);
        }

        private bool HasCircularReferenceHelper(
            SchemeNode node,
            HashSet<SchemeNode> visited,
            HashSet<SchemeNode> recursionStack)
        {
            if (node == null)
                return false;

            visited.Add(node);
            recursionStack.Add(node);

            foreach (var child in node.Children)
            {
                if (!visited.Contains(child))
                {
                    if (HasCircularReferenceHelper(child, visited, recursionStack))
                        return true;
                }
                else if (recursionStack.Contains(child))
                {
                    return true;
                }
            }

            recursionStack.Remove(node);
            return false;
        }

        /// <summary>
        /// 중복 키 찾기
        /// </summary>
        private IEnumerable<string> GetDuplicateKeys()
        {
            var keyGroups = _linearNodes
                .Where(n => !string.IsNullOrEmpty(n.Key))
                .GroupBy(n => n.GetFullPath())
                .Where(g => g.Count() > 1)
                .Select(g => g.Key);

            return keyGroups;
        }

        /// <summary>
        /// 스키마 복사
        /// </summary>
        public Scheme Clone()
        {
            var clonedRoot = Root.Clone();
            return new Scheme(
                SheetName,
                clonedRoot,
                ContentStartRow,
                EndRow,
                StartPosition,
                EndPosition,
                Metadata.Clone());
        }

        /// <summary>
        /// 스키마 병합
        /// </summary>
        public static Scheme Merge(Scheme primary, Scheme secondary)
        {
            if (primary == null)
                return secondary;
            if (secondary == null)
                return primary;

            // TODO: 구현 필요
            throw new NotImplementedException("스키마 병합은 아직 구현되지 않았습니다.");
        }

        public IEnumerator<SchemeNode> GetEnumerator()
        {
            return _linearNodes.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public override string ToString()
        {
            return $"Scheme[Sheet={SheetName}, Nodes={NodeCount}, Rows={ContentStartRow}-{EndRow}]";
        }
    }

    /// <summary>
    /// 스키마 메타데이터
    /// </summary>
    public class SchemeMetadata : ValueObject
    {
        public string Version { get; }
        public string Author { get; }
        public string Description { get; }
        public Dictionary<string, string> CustomProperties { get; }

        private SchemeMetadata()
        {
            Version = "1.0";
            Author = "";
            Description = "";
            CustomProperties = new Dictionary<string, string>();
        }

        private SchemeMetadata(
            string version,
            string author,
            string description,
            Dictionary<string, string> customProperties)
        {
            Version = version ?? "1.0";
            Author = author ?? "";
            Description = description ?? "";
            CustomProperties = customProperties ?? new Dictionary<string, string>();
        }

        public static SchemeMetadata Default() => new SchemeMetadata();

        public static SchemeMetadata Create(
            string version,
            string author,
            string description,
            Dictionary<string, string> customProperties = null)
        {
            return new SchemeMetadata(version, author, description, customProperties);
        }

        public SchemeMetadata Clone()
        {
            return new SchemeMetadata(
                Version,
                Author,
                Description,
                new Dictionary<string, string>(CustomProperties));
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Version;
            yield return Author;
            yield return Description;
            foreach (var kvp in CustomProperties)
            {
                yield return kvp.Key;
                yield return kvp.Value;
            }
        }
    }

    /// <summary>
    /// 스키마 검증 결과
    /// </summary>
    public class SchemeValidationResult
    {
        public bool IsValid { get; }
        public IReadOnlyList<string> Errors { get; }
        public bool HasErrors => Errors?.Count > 0;

        public SchemeValidationResult(bool isValid, IEnumerable<string> errors)
        {
            IsValid = isValid;
            Errors = (errors ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
        }

        public static SchemeValidationResult Success()
        {
            return new SchemeValidationResult(true, Enumerable.Empty<string>());
        }

        public static SchemeValidationResult Failure(params string[] errors)
        {
            return new SchemeValidationResult(false, errors);
        }
    }
}