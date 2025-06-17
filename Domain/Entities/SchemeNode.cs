using ClosedXML.Excel;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToYamlAddin.Domain.Entities
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
        private const string TYPE_MAP = SchemeConstants.NodeTypes.Map;
        private const string TYPE_ARRAY = SchemeConstants.NodeTypes.Array;
        private const string TYPE_KEY = SchemeConstants.NodeTypes.Key;
        private const string TYPE_VALUE = SchemeConstants.NodeTypes.Value;
        private const string TYPE_IGNORE = SchemeConstants.NodeTypes.Ignore;

        // 로깅 방식 변경
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<SchemeNode>();

        private string key = "";
        private SchemeNodeType type = SchemeNodeType.PROPERTY;
        private SchemeNode parent = null;
        private readonly LinkedList<SchemeNode> children = new LinkedList<SchemeNode>();
        private readonly int schemeRowNum;
        private readonly int schemeCellNum;
        private readonly IXLWorksheet sheet;
        private readonly string schemeName;

        public SchemeNode(IXLWorksheet sheet, int rowNum, int cellNum, string schemeName)
        {
            this.sheet = sheet;
            this.schemeRowNum = rowNum;
            this.schemeCellNum = cellNum;
            this.schemeName = schemeName;

            Logger.Debug("SchemeNode 생성: 이름=" + schemeName + ", 행=" + rowNum + ", 열=" + cellNum);

            if (!schemeName.Contains(SchemeConstants.Markers.MarkerPrefix))
            {
                this.key = schemeName;
                this.type = SchemeNodeType.PROPERTY;
                Logger.Debug("PROPERTY 노드 생성: " + key);
            }
            else
            {
                // 원본 CS 코드와 동일하게 구현
                string[] splitted = schemeName.Split(new char[] { SchemeConstants.Markers.MarkerPrefix[0] }, StringSplitOptions.RemoveEmptyEntries);

                // 키와 타입을 분리
                if (splitted.Length > 0)
                {
                    this.key = splitted[0];
                }
                else
                {
                    this.key = "";
                }

                // ARRAY 타입($[])인 경우 특별 처리
                if (schemeName.Contains(SchemeConstants.Markers.ArrayStart))
                {
                    Logger.Debug("ARRAY 형식 감지: " + schemeName);
                    this.type = SchemeNodeType.ARRAY;
                    // A열(루트)에 있는 배열은 항상 이름 없음 (YAML/JSON 표준)
                    if (this.sheet.Cell(this.schemeRowNum, this.schemeCellNum).Address.ColumnNumber == 1)
                    {
                        this.key = "";  // 루트 배열은 이름 없음
                    }
                    Logger.Debug("ARRAY 노드 생성: " + key);
                }
                // KEY 타입($key)인 경우 특별 처리
                else if (schemeName.Contains(SchemeConstants.Markers.DynamicKey))
                {
                    Logger.Debug("KEY 형식 감지: " + schemeName);
                    this.type = SchemeNodeType.KEY;
                    Logger.Debug("KEY 노드 생성: " + key);
                }
                // VALUE 타입($value)인 경우 특별 처리  
                else if (schemeName.Contains(SchemeConstants.Markers.DynamicValue))
                {
                    Logger.Debug("VALUE 형식 감지: " + schemeName);
                    this.type = SchemeNodeType.VALUE;
                    Logger.Debug("VALUE 노드 생성: " + key);
                }
                else
                {
                    // 타입 문자열 추출 - 원본 CS 코드처럼 마지막 요소 사용
                    string typeString = splitted.Length > 0 ? splitted[splitted.Length - 1] : "";

                    Logger.Debug("스키마 문자열 분석: 원본='" + schemeName + "', 키='" + key + "', 타입 문자열='" + typeString + "'");

                    switch (typeString)
                    {
                        case TYPE_MAP:
                            this.type = SchemeNodeType.MAP;
                            // A열(루트)에 있는 맵은 항상 이름 없음 (YAML/JSON 표준)
                            if (this.sheet.Cell(this.schemeRowNum, this.schemeCellNum).Address.ColumnNumber == 1)
                            {
                                this.key = "";  // 루트 맵은 이름 없음
                            }
                            Logger.Debug("MAP 노드 생성: " + key);
                            break;
                        case TYPE_ARRAY:
                            this.type = SchemeNodeType.ARRAY;
                            // A열(루트)에 있는 배열은 항상 이름 없음 (YAML/JSON 표준)
                            if (this.sheet.Cell(this.schemeRowNum, this.schemeCellNum).Address.ColumnNumber == 1)
                            {
                                this.key = "";  // 루트 배열은 이름 없음
                            }
                            Logger.Debug("ARRAY 노드 생성: " + key);
                            break;
                        case TYPE_KEY:
                            this.type = SchemeNodeType.KEY;
                            Logger.Debug("KEY 노드 생성: " + key);
                            break;
                        case TYPE_VALUE:
                            this.type = SchemeNodeType.VALUE;
                            Logger.Debug("VALUE 노드 생성: " + key);
                            break;
                        case TYPE_IGNORE:
                            this.type = SchemeNodeType.IGNORE;
                            Logger.Debug("IGNORE 노드 생성: " + key);
                            break;
                        default:
                            throw new InvalidOperationException(ErrorMessages.Schema.UnknownNodeType + typeString);
                    }
                }
            }
        }

        public void SetParent(SchemeNode parent)
        {
            string parentKey = parent != null ? parent.key : "null";
            Logger.Debug("부모 설정: " + this.key + " -> " + parentKey);
            this.parent = parent;
        }

        public void AddChild(SchemeNode child)
        {
            if (child == null)
            {
                Logger.Warning("null 자식 추가 시도 무시");
                return;
            }

            // 자바 코드 기반으로 타입별 자식 노드 추가 제약 조건 구현
            switch (this.type)
            {
                case SchemeNodeType.KEY:
                    if (child.NodeType == SchemeNodeType.KEY || child.NodeType == SchemeNodeType.PROPERTY)
                    {
                        Logger.Warning($"KEY 노드에 KEY/PROPERTY 추가 불가: {this.key} -> {child.key}");
                        return;
                    }
                    break;
                case SchemeNodeType.PROPERTY:
                    if (child.NodeType == SchemeNodeType.KEY || child.NodeType == SchemeNodeType.PROPERTY)
                    {
                        Logger.Warning($"PROPERTY 노드에 KEY/PROPERTY 추가 불가: {this.key} -> {child.key}");
                        return;
                    }
                    break;
                case SchemeNodeType.ARRAY:
                    // 자바 코드에서는 ARRAY 노드에 PROPERTY나 KEY 노드를 추가할 수 없었으나,
                    // 현재 케이스에서는 이런 제한이 문제를 일으킬 수 있으므로 로깅만 하고 계속 진행합니다.
                    if (child.NodeType == SchemeNodeType.PROPERTY || child.NodeType == SchemeNodeType.KEY)
                    {
                        Logger.Debug($"ARRAY 노드에 {child.NodeType} 노드 추가: {this.key}$[] -> {child.key}");
                    }
                    break;
                case SchemeNodeType.MAP:
                    // 자바 코드에서는 MAP 노드에 VALUE 노드를 추가할 수 없었으나,
                    // 현재 케이스에서는 이런 제한이 문제를 일으킬 수 있으므로 로깅만 하고 계속 진행합니다.
                    if (child.NodeType == SchemeNodeType.VALUE)
                    {
                        Logger.Debug($"MAP 노드에 VALUE 노드 추가: {this.key}${{}} -> {child.key}");
                    }
                    break;
                case SchemeNodeType.VALUE:
                    Logger.Warning($"VALUE 노드에 자식 추가 불가: {this.key} -> {child.key}");
                    return;
                case SchemeNodeType.IGNORE:
                    Logger.Warning($"IGNORE 노드에 자식 추가 시도 무시: {child.key}");
                    return;
            }

            child.SetParent(this);
            children.AddLast(child);
            Logger.Debug($"자식 노드 추가됨: {this.key} ({this.type}) -> {child.key} ({child.NodeType})");
        }

        /// <summary>
        /// 이 노드와 모든 자식 노드를 포함하는 평면화된 목록을 반환합니다.
        /// </summary>
        /// <returns>노드 구조를 평면화한 목록</returns>
        public LinkedList<SchemeNode> Linear()
        {
            Logger.Debug("Linear() 호출: " + key);
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
                Logger.Warning("시트 또는 행이 null임: " + key);
                return string.Empty;
            }

            IXLCell cell = row.Cell(schemeCellNum);
            if (cell == null || cell.IsEmpty())
            {
                Logger.Debug("셀이 비어있음: 행=" + row.RowNumber() + ", 열=" + schemeCellNum);
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
                Logger.Debug("키를 가져올 수 없음: 타입=" + type + ", 행=" + rowNumber);
                return string.Empty;
            }

            // 1. $key 이름을 가진 노드는 항상 셀 값을 키로 사용
            if (type == SchemeNodeType.KEY && !string.IsNullOrEmpty(this.schemeName) && this.schemeName.Contains("$key"))
            {
                IXLCell cell = row.Cell(schemeCellNum);
                if (cell != null && !cell.IsEmpty())
                {
                    object cellValue = ExcelCellValueResolver.GetCellValue(cell);
                    string cellValueStr = cellValue != null ? cellValue.ToString() : string.Empty;
                    Logger.Debug("$key 노드의 실제 셀 값: " + cellValueStr);
                    return cellValueStr;
                }
            }

            // 2. 키가 이미 있으면 그대로 사용 (Java와 동일)
            if (!string.IsNullOrEmpty(key))
            {
                return key;
            }

            // 3. KEY 노드인 경우 셀 값 또는 자식 노드 값 사용 (Java와 동일)
            if (type == SchemeNodeType.KEY)
            {
                // 값 노드가 있는 경우 해당 값을 사용
                SchemeNode valueNode = children.FirstOrDefault(c => c.NodeType == SchemeNodeType.VALUE);
                if (valueNode != null)
                {
                    object value = valueNode.GetValue(row);
                    string valueStr = value != null ? value.ToString() : string.Empty;
                    Logger.Debug("KEY 노드의 값 노드 값: " + valueStr);
                    return valueStr;
                }

                // 값 노드가 없는 경우 직접 셀 값 사용
                IXLCell cell = row.Cell(schemeCellNum);
                if (cell != null && !cell.IsEmpty())
                {
                    object cellValue = ExcelCellValueResolver.GetCellValue(cell);
                    string cellValueStr = cellValue != null ? cellValue.ToString() : string.Empty;
                    Logger.Debug("KEY 노드의 셀 값: " + cellValueStr);
                    return cellValueStr;
                }
            }

            // 4. 부모가 있고 부모가 키를 제공할 수 있는 경우 부모의 키 사용 (Java와 동일)
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
                        Logger.Debug("부모 노드의 셀 값: " + parentCellValueStr);
                        return parentCellValueStr;
                    }
                }

                // 부모의 키를 사용
                return parent.GetKey(row);
            }

            // 5. 기본값
            Logger.Warning("키를 결정할 수 없음: " + this + ", 부모=" + parent);
            return string.Empty;
        }

        public override string ToString()
        {
            return key + ":" + type;
        }

        public bool IsRoot => parent == null;
        public SchemeNode Parent => parent;
        public int SchemeRowNum => schemeRowNum;
        public int CellNum => schemeCellNum;
        public string Key => key;
        public SchemeNodeType NodeType => type;
        public IEnumerable<SchemeNode> Children => children;
        public int ChildCount => children.Count;
        public string SchemeName => schemeName;

        public bool IsContainer =>
            type == SchemeNodeType.MAP ||
            type == SchemeNodeType.ARRAY ||
            type == SchemeNodeType.KEY;

        public bool IsKeyProvidable =>
            type == SchemeNodeType.KEY ||
            type == SchemeNodeType.PROPERTY;
    }
}
