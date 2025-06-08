using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 데이터를 Excel 행으로 매핑하는 클래스
    /// </summary>
    public class DynamicDataMapper
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicDataMapper>();

        /// <summary>
        /// Excel 행 데이터를 표현하는 클래스
        /// </summary>
        public class ExcelRow
        {
            private readonly Dictionary<int, object> cells = new Dictionary<int, object>();

            public void SetCell(int column, object value)
            {
                cells[column] = value;
            }

            public object GetCell(int column)
            {
                return cells.ContainsKey(column) ? cells[column] : null;
            }

            public int GetMaxColumn()
            {
                return cells.Keys.DefaultIfEmpty(0).Max();
            }

            public void WriteToWorksheet(IXLWorksheet worksheet, int rowNumber)
            {
                foreach (var cell in cells)
                {
                    var xlCell = worksheet.Cell(rowNumber, cell.Key);
                    
                    if (cell.Value == null)
                    {
                        xlCell.Value = "";
                    }
                    else if (cell.Value is bool boolValue)
                    {
                        xlCell.Value = boolValue;
                    }
                    else if (cell.Value is int intValue)
                    {
                        xlCell.Value = intValue;
                    }
                    else if (cell.Value is double doubleValue)
                    {
                        xlCell.Value = doubleValue;
                    }
                    else if (cell.Value is DateTime dateValue)
                    {
                        xlCell.Value = dateValue;
                    }
                    else
                    {
                        xlCell.Value = cell.Value.ToString();
                    }
                }
            }

            public Dictionary<int, object> GetAllCells()
            {
                return new Dictionary<int, object>(cells);
            }
        }

        public List<ExcelRow> MapToExcelRows(
            YamlNode data,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            Logger.Information("YAML 데이터를 Excel 행으로 매핑 시작");
            var rows = new List<ExcelRow>();

            if (data is YamlSequenceNode sequence)
            {
                Logger.Debug($"시퀀스 노드 처리: {sequence.Children.Count}개 항목");
                foreach (var item in sequence.Children)
                {
                    var mappedRows = MapItem(item, scheme, pattern);
                    rows.AddRange(mappedRows);
                }
            }
            else if (data is YamlMappingNode mapping)
            {
                Logger.Debug("단일 매핑 노드 처리");
                var row = MapSingleItem(mapping, scheme, pattern);
                rows.Add(row);
            }

            Logger.Information($"매핑 완료: {rows.Count}개 행 생성");
            return rows;
        }

        private List<ExcelRow> MapItem(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var rows = new List<ExcelRow>();

            // Weapons.yaml의 경우 수평 확장만 사용
            // 각 무기 항목은 하나의 행으로 매핑되어야 함
            rows.Add(MapHorizontally(item, scheme, pattern));

            return rows;
        }

        private ExcelRow MapHorizontally(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var row = new ExcelRow();

            if (item is YamlMappingNode mapping)
            {
                // ^ 마커 (무시 마커)를 첫 번째 컬럼에 설정
                // 주의: 데이터 행의 첫 번째 컬럼은 빈 셀로 남겨둬야 함
                // row.SetCell(1, "^"); // 이 줄을 제거하고 빈 셀로 남김

                // 속성 매핑
                Logger.Information("========== MapHorizontally 속성 매핑 시작 ==========");
                foreach (var prop in mapping.Children)
                {
                    var key = prop.Key.ToString();
                    Logger.Debug($"속성 처리: {key}, 타입: {prop.Value.GetType().Name}");
                    
                    if (prop.Value is YamlSequenceNode nestedArray)
                    {
                        Logger.Debug($"→ {key}는 배열 타입, MapNestedArray로 처리");
                        // 중첩 배열은 별도 처리 - 배열 자체의 컬럼 인덱스는 찾지 않음
                        MapNestedArray(row, nestedArray, key, scheme, pattern);
                    }
                    else if (prop.Value is YamlMappingNode nestedObject)
                    {
                        Logger.Debug($"→ {key}는 객체 타입, MapNestedObject로 처리");
                        // 중첩 객체는 속성들을 확장하여 매핑
                        MapNestedObject(row, nestedObject, key, scheme, pattern);
                    }
                    else
                    {
                        // 단순 값만 직접 매핑
                        var columnIndex = scheme.GetColumnIndex(key);
                        Logger.Information($"🔍 단순 속성 '{key}' 컬럼 인덱스 조회 결과: {columnIndex}");
                        
                        if (columnIndex > 0)
                        {
                            var value = ConvertValue(prop.Value);
                            row.SetCell(columnIndex, value);
                            Logger.Information($"✓ 단순 속성 매핑: {key} -> 컬럼 {columnIndex}: {value}");
                        }
                        else
                        {
                            // 컬럼 인덱스를 찾을 수 없는 경우, 혼재 구조인지 확인
                            Logger.Warning($"✗ {key} 속성의 컬럼 인덱스를 찾을 수 없음 - 혼재 구조 확인 중...");
                            
                            // XML에서 변환된 혼재 구조(_Step, _SubStep, __text 등)인지 확인
                            if (prop.Value is YamlScalarNode scalar)
                            {
                                // 스칼라 값인데 컬럼을 찾을 수 없는 경우
                                // {key}__{textKey} 형태의 복합 속성일 가능성 체크
                                string textContentKey = $"{key}.__text";
                                var textColumnIndex = scheme.GetColumnIndex(textContentKey);
                                
                                if (textColumnIndex > 0)
                                {
                                    // XML 혼재 구조에서 텍스트 내용만 있는 경우
                                    var value = ConvertValue(prop.Value);
                                    row.SetCell(textColumnIndex, value);
                                    Logger.Information($"✓ 혼재 구조 텍스트 매핑: {textContentKey} -> 컬럼 {textColumnIndex}: {value}");
                                }
                                else
                                {
                                    Logger.Warning($"✗ {key} 속성과 관련된 어떤 컬럼도 찾을 수 없음 (혼재 구조 포함)");
                                }
                            }
                        }
                    }
                }
                Logger.Information("========== MapHorizontally 속성 매핑 완료 ==========");
            }
            else if (item is YamlScalarNode scalar)
            {
                // 스칼라 값인 경우
                row.SetCell(2, ConvertValue(scalar));
            }

            return row;
        }

        private void MapNestedArray(
            ExcelRow row,
            YamlSequenceNode array,
            string arrayName,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            Logger.Information($"========== MapNestedArray 시작: {arrayName} ==========");
            Logger.Debug($"배열 크기: {array.Children.Count}개 요소");
            
            // 배열을 수평으로 확장하여 매핑
            int elementIndex = 0;
            foreach (var element in array.Children)
            {
                Logger.Debug($"--- 배열 요소 [{elementIndex}] 처리 시작 ---");
                
                if (element is YamlMappingNode elementMapping)
                {
                    Logger.Debug($"요소 타입: YamlMappingNode, 속성 개수: {elementMapping.Children.Count}");
                    
                    // 각 요소의 속성들을 개별 컬럼으로 매핑
                    foreach (var prop in elementMapping.Children)
                    {
                        var propKey = prop.Key.ToString();
                        var fullKey = $"{arrayName}[{elementIndex}].{propKey}";
                        
                        Logger.Debug($"속성 매핑 시도: {fullKey}");
                        
                        // 재귀적으로 속성 매핑
                        MapPropertyRecursively(row, prop.Key.ToString(), prop.Value, fullKey, scheme);
                    }
                }
                else
                {
                    Logger.Debug($"요소 타입: {element.GetType().Name}");
                    
                    // 단순 값인 경우
                    var fullKey = $"{arrayName}[{elementIndex}]";
                    Logger.Debug($"단순 값 매핑 시도: {fullKey}");
                    
                    var columnIndex = scheme.GetColumnIndex(fullKey);
                    Logger.Debug($"scheme.GetColumnIndex('{fullKey}') = {columnIndex}");
                    
                    if (columnIndex > 0)
                    {
                        var value = ConvertValue(element);
                        row.SetCell(columnIndex, value);
                        Logger.Information($"✓ 매핑 성공: {fullKey} -> 컬럼 {columnIndex}: {value}");
                    }
                    else
                    {
                        Logger.Warning($"✗ 매핑 실패: {fullKey} - 컬럼 인덱스를 찾을 수 없음");
                    }
                }
                
                elementIndex++;
            }
            
            Logger.Information($"========== MapNestedArray 완료: {arrayName} ==========");
        }

        private List<ExcelRow> ExpandVertically(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var rows = new List<ExcelRow>();

            if (item is YamlMappingNode mapping)
            {
                // 기본 행 생성
                var baseRow = new ExcelRow();
                // 첫 번째 컬럼은 빈 셀로 남김 (데이터 행에서는 ^ 마커를 사용하지 않음)
                // baseRow.SetCell(1, "^"); // 제거

                // 단순 속성들 먼저 처리
                var simpleProps = new Dictionary<string, object>();
                var arrayProps = new Dictionary<string, YamlSequenceNode>();

                foreach (var prop in mapping.Children)
                {
                    var key = prop.Key.ToString();
                    if (prop.Value is YamlSequenceNode array)
                    {
                        arrayProps[key] = array;
                    }
                    else
                    {
                        var columnIndex = scheme.GetColumnIndex(key);
                        if (columnIndex > 0)
                        {
                            baseRow.SetCell(columnIndex, ConvertValue(prop.Value));
                        }
                    }
                }

                // 수직 확장이 필요한 배열 찾기
                var verticalArrays = arrayProps.Where(a => 
                    pattern.Arrays.ContainsKey(a.Key) && 
                    pattern.Arrays[a.Key].RequiresMultipleRows).ToList();

                if (verticalArrays.Any())
                {
                    // 가장 긴 배열 기준으로 행 생성
                    int maxRows = verticalArrays.Max(a => a.Value.Children.Count);
                    
                    for (int i = 0; i < maxRows; i++)
                    {
                        var newRow = new ExcelRow();
                        
                        // 기본 행의 데이터 복사
                        foreach (var cell in baseRow.GetAllCells())
                        {
                            newRow.SetCell(cell.Key, cell.Value);
                        }

                        // 각 배열의 i번째 요소 추가
                        foreach (var array in verticalArrays)
                        {
                            if (i < array.Value.Children.Count)
                            {
                                var element = array.Value.Children[i];
                                var columnIndex = scheme.GetColumnIndex(array.Key);
                                if (columnIndex > 0)
                                {
                                    newRow.SetCell(columnIndex, ConvertValue(element));
                                }
                            }
                        }

                        rows.Add(newRow);
                    }
                }
                else
                {
                    // 수직 확장이 필요 없으면 기본 행만 반환
                    rows.Add(baseRow);
                }
            }

            return rows;
        }

        private ExcelRow MapSingleItem(
            YamlMappingNode mapping,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            return MapHorizontally(mapping, scheme, pattern);
        }

        private void MapNestedObject(
            ExcelRow row,
            YamlMappingNode obj,
            string objectName,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            Logger.Information($"========== MapNestedObject 시작: {objectName} ==========");
            Logger.Debug($"객체 속성 개수: {obj.Children.Count}");
            
            // XML 혼재 구조 감지: _속성과 __text가 함께 있는지 확인
            bool hasXmlAttributes = obj.Children.Any(p => p.Key.ToString().StartsWith("_") && p.Key.ToString() != "__text");
            bool hasTextContent = obj.Children.Any(p => p.Key.ToString() == "__text");
            bool isXmlMixedContent = hasXmlAttributes && hasTextContent;
            
            if (isXmlMixedContent)
            {
                Logger.Information($"🔍 XML 혼재 구조 감지: '{objectName}' - 속성과 텍스트 내용이 혼재됨");
                Logger.Information($"  - XML 속성들: [{string.Join(", ", obj.Children.Where(p => p.Key.ToString().StartsWith("_")).Select(p => p.Key.ToString()))}]");
                Logger.Information($"  - 텍스트 내용: {obj.Children.Where(p => p.Key.ToString() == "__text").Select(p => p.Value.ToString()).FirstOrDefault()}");
            }
            
            // 중첩 객체의 각 속성을 개별 컬럼으로 매핑
            foreach (var prop in obj.Children)
            {
                var propKey = prop.Key.ToString();
                var fullKey = $"{objectName}.{propKey}";
                
                Logger.Debug($"속성 매핑 시도: {fullKey}");
                Logger.Debug($"속성 값 타입: {prop.Value.GetType().Name}");
                
                var columnIndex = scheme.GetColumnIndex(fullKey);
                Logger.Debug($"scheme.GetColumnIndex('{fullKey}') = {columnIndex}");
                
                // XML 속성 특별 디버깅
                if (propKey.StartsWith("_Arg") || propKey == "__text")
                {
                    Logger.Information($"🔍🔍🔍 XML 속성 매핑 시도: '{fullKey}' -> 컬럼 인덱스 = {columnIndex}");
                    scheme.DebugAllMappings(); // 모든 매핑 상황 출력
                }
                
                if (columnIndex > 0)
                {
                    var value = ConvertValue(prop.Value);
                    row.SetCell(columnIndex, value);
                    Logger.Information($"✓ 매핑 성공: {fullKey} -> 컬럼 {columnIndex}: {value}");
                    
                    if (propKey.StartsWith("_Arg") || propKey == "__text")
                    {
                        Logger.Information($"  ★★★ XML 속성 매핑 성공: {fullKey} -> 컬럼 {columnIndex}: {value}");
                    }
                }
                else
                {
                    Logger.Warning($"✗ 매핑 실패: {fullKey} - 컬럼 인덱스를 찾을 수 없음");
                    
                    // XML 속성인 경우 특별 경고
                    if (propKey.StartsWith("_Arg") || propKey == "__text")
                    {
                        Logger.Warning($"  ⚠️⚠️⚠️ XML 속성 매핑 실패: {fullKey} - 스키마에서 누락된 것으로 보임");
                    }
                    
                    // XML 혼재 구조인 경우 특별 처리
                    if (isXmlMixedContent)
                    {
                        Logger.Information($"  🔧 XML 혼재 구조 특별 처리 시도 중...");
                        
                        // Step="1" SubStep="0">Starter 형태의 구조인 경우
                        // 단일 컬럼으로 합치거나 개별 속성으로 매핑 시도
                        if (propKey == "__text")
                        {
                            // 텍스트 내용을 원래 속성명으로 매핑 시도
                            var originalColumnIndex = scheme.GetColumnIndex(objectName);
                            if (originalColumnIndex > 0)
                            {
                                var value = ConvertValue(prop.Value);
                                row.SetCell(originalColumnIndex, value);
                                Logger.Information($"✓ XML 혼재 구조 텍스트 매핑: {objectName} -> 컬럼 {originalColumnIndex}: {value}");
                            }
                        }
                        else if (propKey.StartsWith("_"))
                        {
                            // 속성을 별도 컬럼에 매핑 시도 (예: Type_Step, Type_SubStep)
                            var attributeName = propKey.Substring(1); // _ 제거
                            var attributeColumnName = $"{objectName}_{attributeName}";
                            var attributeColumnIndex = scheme.GetColumnIndex(attributeColumnName);
                            
                            if (attributeColumnIndex > 0)
                            {
                                var value = ConvertValue(prop.Value);
                                row.SetCell(attributeColumnIndex, value);
                                Logger.Information($"✓ XML 속성 매핑: {attributeColumnName} -> 컬럼 {attributeColumnIndex}: {value}");
                            }
                        }
                    }
                }
            }
            
            Logger.Information($"========== MapNestedObject 완료: {objectName} ==========");
        }

        private void MapPropertyRecursively(
            ExcelRow row,
            string propKey,
            YamlNode propValue,
            string fullKey,
            DynamicSchemaBuilder.ExcelScheme scheme)
        {
            if (propValue is YamlMappingNode nestedObj)
            {
                // 중첩된 객체 재귀적 처리
                Logger.Debug($"→ {propKey}는 중첩된 객체");
                foreach (var nestedProp in nestedObj.Children)
                {
                    var nestedKey = nestedProp.Key.ToString();
                    var nestedFullKey = $"{fullKey}.{nestedKey}";
                    
                    // 재귀 호출
                    MapPropertyRecursively(row, nestedKey, nestedProp.Value, nestedFullKey, scheme);
                }
            }
            else if (propValue is YamlSequenceNode nestedArray)
            {
                // 중첩된 배열 처리
                Logger.Debug($"→ {propKey}는 중첩된 배열");
                
                // 배열의 각 요소 처리
                for (int i = 0; i < nestedArray.Children.Count; i++)
                {
                    var element = nestedArray.Children[i];
                    var elementFullKey = $"{fullKey}[{i}]";
                    
                    if (element is YamlMappingNode elemMapping)
                    {
                        // 배열 요소가 객체인 경우 재귀적 처리
                        foreach (var elemProp in elemMapping.Children)
                        {
                            var elemPropKey = elemProp.Key.ToString();
                            var elemPropFullKey = $"{elementFullKey}.{elemPropKey}";
                            
                            MapPropertyRecursively(row, elemPropKey, elemProp.Value, elemPropFullKey, scheme);
                        }
                    }
                    else
                    {
                        // 단순 배열 요소
                        var columnIndex = scheme.GetColumnIndex(elementFullKey);
                        if (columnIndex > 0)
                        {
                            var value = ConvertValue(element);
                            row.SetCell(columnIndex, value);
                            Logger.Information($"✓ 배열 요소 매핑: {elementFullKey} -> 컬럼 {columnIndex}: {value}");
                        }
                    }
                }
                
                // 배열 자체도 하나의 컬럼으로 직렬화
                var arrayColumnIndex = scheme.GetColumnIndex(fullKey);
                if (arrayColumnIndex > 0)
                {
                    var value = ConvertArray(nestedArray);
                    row.SetCell(arrayColumnIndex, value);
                    Logger.Information($"✓ 배열 전체 매핑: {fullKey} -> 컬럼 {arrayColumnIndex}");
                }
            }
            else
            {
                // 일반 속성 (리프 노드)
                var columnIndex = scheme.GetColumnIndex(fullKey);
                Logger.Debug($"scheme.GetColumnIndex('{fullKey}') = {columnIndex}");
                
                if (columnIndex > 0)
                {
                    var value = ConvertValue(propValue);
                    row.SetCell(columnIndex, value);
                    Logger.Information($"✓ 매핑 성공: {fullKey} -> 컬럼 {columnIndex}: {value}");
                }
                else
                {
                    Logger.Warning($"✗ 매핑 실패: {fullKey} - 컬럼 인덱스를 찾을 수 없음");
                }
            }
        }

        private object ConvertValue(YamlNode node)
        {
            if (node is YamlScalarNode scalar)
                return ConvertScalar(scalar);
            else if (node is YamlSequenceNode)
                return "[Array]";
            else if (node is YamlMappingNode)
                return "[Object]";
            else
                return null;
        }

        private object ConvertArray(YamlSequenceNode sequence)
        {
            // 배열을 문자열로 직렬화
            var items = new List<string>();
            foreach (var item in sequence.Children)
            {
                if (item is YamlScalarNode scalar)
                {
                    // 배열 요소의 개행문자도 이스케이프 처리
                    items.Add(EscapeNewlines(scalar.Value));
                }
                else if (item is YamlMappingNode mapping)
                {
                    // 중첩된 객체를 간단한 형식으로 변환
                    var objStr = ConvertObject(mapping);
                    items.Add(objStr.ToString());
                }
                else if (item is YamlSequenceNode nestedSeq)
                {
                    // 중첩된 배열
                    var arrStr = ConvertArray(nestedSeq);
                    items.Add(arrStr.ToString());
                }
            }
            return "[" + string.Join(", ", items) + "]";
        }

        private object ConvertObject(YamlMappingNode mapping)
        {
            // 객체를 문자열로 직렬화
            var props = new List<string>();
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var value = ConvertValue(kvp.Value);
                // 개행문자 이스케이프 처리
                if (value is string strValue)
                {
                    value = EscapeNewlines(strValue);
                }
                props.Add($"{key}: {value}");
            }
            return "{" + string.Join(", ", props) + "}";
        }

        private object ConvertScalar(YamlScalarNode scalar)
        {
            var value = scalar.Value;

            if (string.IsNullOrEmpty(value))
                return "";

            // 동적 타입 추론
            // bool
            if (bool.TryParse(value, out bool boolResult))
                return boolResult;

            // int
            if (int.TryParse(value, out int intResult))
                return intResult;

            // double (소수점이 있는 경우)
            if (value.Contains('.') && double.TryParse(value, NumberStyles.Any,
                CultureInfo.InvariantCulture, out double doubleResult))
                return doubleResult;

            // DateTime
            if (DateTime.TryParse(value, out DateTime dateResult))
                return dateResult;

            // 기본값은 문자열 - 개행문자 이스케이프 처리
            return EscapeNewlines(value);
        }

        private string EscapeNewlines(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;
            
            // \n을 리터럴 문자로 변환
            return value.Replace("\n", "\\n")
                        .Replace("\r\n", "\\r\\n")
                        .Replace("\r", "\\r")
                        .Replace("\t", "\\t");
        }
    }
}