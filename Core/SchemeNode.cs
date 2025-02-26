using Microsoft.Extensions.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public class SchemeNode
    {
        public enum SchemeNodeType
        {
            PROPERTY,
            KEY,
            VALUE,
            MAP,
            ARRAY,
            IGNORE
        }
        
        // 노드 타입 구분을 위한 상수
        private const string TYPE_MAP = "{}";
        private const string TYPE_ARRAY = "[]";
        private const string TYPE_KEY = "key";
        private const string TYPE_VALUE = "value";
        private const string TYPE_IGNORE = "^";

        // 로깅 방식 변경
        private static readonly ILogger<SchemeNode> Logger = CreateLogger();

        private static ILogger<SchemeNode> CreateLogger()
        {
            // 간단한 로거 팩토리 생성
            var loggerFactory = LoggerFactory.Create(builder => 
            {
                // AddConsole 대신 디버그 로깅만 사용
                builder.SetMinimumLevel(LogLevel.Debug);
            });
            
            return loggerFactory.CreateLogger<SchemeNode>();
        }

        private string key = "";
        private SchemeNodeType type = SchemeNodeType.PROPERTY;
        private SchemeNode parent = null;
        private readonly LinkedList<SchemeNode> children = new LinkedList<SchemeNode>();
        private readonly int schemeRowNum;
        private readonly int schemeCellNum;
        private readonly IXLWorksheet sheet;

        public SchemeNode(IXLWorksheet sheet, int rowNum, int cellNum, string schemeName)
        {
            this.sheet = sheet;
            this.schemeRowNum = rowNum;
            this.schemeCellNum = cellNum;
            
            Logger.LogDebug("SchemeNode 생성: 이름={Name}, 행={Row}, 열={Cell}", schemeName, rowNum, cellNum);

            if (!schemeName.Contains("$"))
            {
                this.key = schemeName;
                this.type = SchemeNodeType.PROPERTY;
                Logger.LogDebug("PROPERTY 노드 생성: {Name}", key);
            }
            else
            {
                string[] splitted = schemeName.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries);
                // 키와 타입을 분리
                if (splitted.Length > 0)
                {
                    this.key = splitted[0];
                }
                else
                {
                    this.key = "";
                }
                
                string typeString = splitted.Length > 0 ? splitted[splitted.Length - 1] : "";

                switch (typeString)
                {
                    case TYPE_MAP:
                        this.type = SchemeNodeType.MAP;
                        // 루트 MAP 노드에 대한 처리
                        if (string.IsNullOrEmpty(key) || this.sheet.Cell(this.schemeRowNum, this.schemeCellNum).Address.ColumnNumber == 1)
                        {
                            this.key = "";  // 명시적으로 빈 문자열 설정
                        }
                        Logger.LogDebug("MAP 노드 생성: {Name}", key);
                        break;
                    case TYPE_ARRAY:
                        this.type = SchemeNodeType.ARRAY;
                        // 루트 ARRAY 노드에 대한 처리
                        if (string.IsNullOrEmpty(key) || this.sheet.Cell(this.schemeRowNum, this.schemeCellNum).Address.ColumnNumber == 1)
                        {
                            this.key = "";  // 명시적으로 빈 문자열 설정
                        }
                        Logger.LogDebug("ARRAY 노드 생성: {Name}", key);
                        break;
                    case TYPE_KEY:
                        this.type = SchemeNodeType.KEY;
                        Logger.LogDebug("KEY 노드 생성: {Name}", key);
                        break;
                    case TYPE_VALUE:
                        this.type = SchemeNodeType.VALUE;
                        Logger.LogDebug("VALUE 노드 생성: {Name}", key);
                        break;
                    case TYPE_IGNORE:
                        this.type = SchemeNodeType.IGNORE;
                        Logger.LogDebug("IGNORE 노드 생성: {Name}", key);
                        break;
                    default:
                        throw new InvalidOperationException($"알 수 없는 노드 유형: {typeString}");
                }
            }
        }

        public void SetParent(SchemeNode parent)
        {
            string parentKey = parent != null ? parent.key : "null";
            Logger.LogDebug("부모 설정: {ChildName} -> {ParentName}", this.key, parentKey);
            this.parent = parent;
        }

        public void AddChild(SchemeNode child)
        {
            if (child == null)
            {
                Logger.LogWarning("null 자식 추가 시도 무시");
                return;
            }
            
            // Java 코드와 유사한 검증 로직 추가
            switch (this.type)
            {
                case SchemeNodeType.KEY:
                case SchemeNodeType.PROPERTY:
                    if (child.NodeType == SchemeNodeType.KEY || child.NodeType == SchemeNodeType.PROPERTY)
                    {
                        Logger.LogWarning("PROPERTY 또는 KEY 노드에 다른 PROPERTY 또는 KEY 노드 추가 시도 무시: {Parent} -> {Child}", this.key, child.key);
                        return;
                    }
                    break;
                case SchemeNodeType.IGNORE:
                    Logger.LogWarning("IGNORE 노드에 자식 추가 시도 무시: {Child}", child.key);
                    return;
            }
            
            child.SetParent(this);
            children.AddLast(child);
            Logger.LogDebug("자식 노드 추가됨: {Parent} -> {Child}", this.key, child.key);
        }

        /// <summary>
        /// 이 노드와 모든 자식 노드를 포함하는 평면화된 목록을 반환합니다.
        /// </summary>
        /// <returns>노드 구조를 평면화한 목록</returns>
        public LinkedList<SchemeNode> Linear()
        {
            Logger.LogDebug("Linear() 호출: {Node}", key);
            var result = new LinkedList<SchemeNode>();
            result.AddLast(this);
            
            foreach (var child in children)
            {
                foreach (var node in child.Linear())
                {
                    result.AddLast(node);
                }
            }
            
            return result;
        }

        public object GetValue(IXLRow row)
        {
            if (sheet == null || row == null)
            {
                Logger.LogWarning("시트 또는 행이 null임: {Key}", key);
                return string.Empty;
            }

            IXLCell cell = row.Cell(schemeCellNum);
            if (cell == null || cell.IsEmpty()) 
            {
                Logger.LogDebug("셀이 비어있음: 행={Row}, 열={Column}", row.RowNumber(), schemeCellNum);
                return string.Empty;
            }

            return ExcelCellValueResolver.GetCellValue(cell);
        }

        public string GetKey(IXLRow row)
        {
            // 기본 검증
            if (!IsKeyProvidable || sheet == null || row == null)
            {
                int rowNumber = row != null ? row.RowNumber() : -1;
                Logger.LogDebug("키를 가져올 수 없음: 타입={Type}, 행={Row}", type, rowNumber);
                return string.Empty;
            }

            // 1. 키가 이미 있으면 그대로 사용 (Java와 동일)
            if (!string.IsNullOrEmpty(key))
            {
                return key;
            }

            // 2. KEY 노드인 경우 셀 값 또는 자식 노드 값 사용 (Java와 동일)
            if (type == SchemeNodeType.KEY)
            {
                // 값 노드가 있는 경우 해당 값을 사용
                SchemeNode valueNode = children.FirstOrDefault(c => c.NodeType == SchemeNodeType.VALUE);
                if (valueNode != null)
                {
                    object value = valueNode.GetValue(row);
                    string valueStr = value != null ? value.ToString() : string.Empty;
                    Logger.LogDebug("KEY 노드의 값 노드 값: {Value}", valueStr);
                    return valueStr;
                }
                
                // 값 노드가 없는 경우 직접 셀 값 사용
                IXLCell cell = row.Cell(schemeCellNum);
                if (cell != null && !cell.IsEmpty())
                {
                    object cellValue = ExcelCellValueResolver.GetCellValue(cell);
                    string cellValueStr = cellValue != null ? cellValue.ToString() : string.Empty;
                    Logger.LogDebug("KEY 노드의 셀 값: {Value}", cellValueStr);
                    return cellValueStr;
                }
            }

            // 3. 부모가 있고 부모가 키를 제공할 수 있는 경우 부모의 키 사용 (Java와 동일)
            if (parent != null && parent.IsKeyProvidable)
            {
                // 부모의 키가 비어있는 경우 부모 셀의 값 사용
                if (string.IsNullOrEmpty(parent.key))
                {
                    IXLCell parentCell = row.Cell(parent.schemeCellNum);
                    if (parentCell != null && !parentCell.IsEmpty())
                    {
                        object parentCellValue = ExcelCellValueResolver.GetCellValue(parentCell);
                        string parentCellValueStr = parentCellValue != null ? parentCellValue.ToString() : string.Empty;
                        Logger.LogDebug("부모 노드의 셀 값: {Value}", parentCellValueStr);
                        return parentCellValueStr;
                    }
                }
                
                // 부모의 키를 사용
                return parent.GetKey(row);
            }

            // 4. 기본값
            Logger.LogWarning("키를 결정할 수 없음: {Self}, 부모={Parent}", this, parent);
            return string.Empty;
        }

        public override string ToString()
        {
            return $"{key}:{type}";
        }

        public bool IsRoot => parent == null;
        public SchemeNode Parent => parent;
        public int SchemeRowNum => schemeRowNum;
        public int CellNum => schemeCellNum;
        public string Key => key;
        public SchemeNodeType NodeType => type;
        public IEnumerable<SchemeNode> Children => children;
        public int ChildCount => children.Count;

        public bool IsContainer => 
            type == SchemeNodeType.MAP ||
            type == SchemeNodeType.ARRAY ||
            type == SchemeNodeType.KEY;
        
        public bool IsKeyProvidable => 
            type == SchemeNodeType.KEY ||
            type == SchemeNodeType.PROPERTY;
    }
}
