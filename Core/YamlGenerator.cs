using ExcelToJsonAddin.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelToJsonAddin.Config;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public class YamlGenerator
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlGenerator>();

        private readonly Scheme _scheme;
        private readonly IXLWorksheet _sheet;

        public YamlGenerator(Scheme scheme)
        {
            _scheme = scheme;
            _sheet = scheme.Sheet;
            Logger.Debug("YamlGenerator 초기화: 스키마 노드 타입={0}", scheme.Root.NodeType);
        }

        // 외부에서 호출할 정적 메서드 추가
        public static string Generate(Scheme scheme, YamlStyle style = YamlStyle.Block, 
            int indentSize = 2, bool preserveQuotes = false, bool includeEmptyFields = false)
        {
            try 
            {
                var generator = new YamlGenerator(scheme);
                object rootObj = generator.ProcessRootNode();
                
                // 필요한 경우 빈 속성 제거
                if (!includeEmptyFields)
                {
                    OrderedYamlFactory.RemoveEmptyProperties(rootObj);
                }
                
                // YAML 문자열로 직렬화
                return OrderedYamlFactory.SerializeToYaml(rootObj, indentSize, style, preserveQuotes);
            }
            catch (Exception ex)
            {
                // 로그를 사용할 수 없으므로 디버그 출력
                Debug.WriteLine($"YAML 생성 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        public string Generate(YamlStyle style = YamlStyle.Block, int indentSize = 2, bool preserveQuotes = false)
        {
            try
            {
                Logger.Debug("YAML 생성 시작: 스타일={0}, 들여쓰기={1}", style, indentSize);
                
                // 루트 노드 처리
                object rootObj = ProcessRootNode();
                
                // YAML 문자열로 직렬화
                return OrderedYamlFactory.SerializeToYaml(rootObj, indentSize, style, preserveQuotes);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "YAML 생성 중 오류 발생");
                throw;
            }
        }
        
        private YamlObject ProcessMapNode(SchemeNode node)
        {
            YamlObject result = OrderedYamlFactory.CreateObject();
            
            // 모든 데이터 행에 대해 처리
            for (int rowNum = _scheme.ContentStartRowNum; rowNum <= _scheme.EndRowNum; rowNum++)
            {
                IXLRow row = _sheet.Row(rowNum);
                if (row == null) continue;
                
                // 각 자식 노드에 대해 처리
                foreach (var child in node.Children)
                {
                    string key = GetNodeKey(child, row);
                    if (string.IsNullOrEmpty(key)) continue;
                    
                    // PROPERTY 노드 처리
                    if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                    {
                        object value = child.GetValue(row);
                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            if (!result.ContainsKey(key))
                            {
                                result.Add(key, value);
                            }
                        }
                    }
                    // MAP 노드 처리
                    else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                    {
                        if (!result.ContainsKey(key))
                        {
                            YamlObject childMap = OrderedYamlFactory.CreateObject();
                            AddChildProperties(child, childMap, row);
                            if (childMap.HasValues)
                            {
                                result.Add(key, childMap);
                            }
                        }
                    }
                    // ARRAY 노드 처리
                    else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                    {
                        if (!result.ContainsKey(key))
                        {
                            YamlArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues)
                            {
                                result.Add(key, childArray);
                            }
                        }
                    }
                }
            }
            
            return result;
        }
        
        private YamlArray ProcessArrayNode(SchemeNode node)
        {
            YamlArray result = OrderedYamlFactory.CreateArray();
            
            // 모든 데이터 행에 대해 처리
            for (int rowNum = _scheme.ContentStartRowNum; rowNum <= _scheme.EndRowNum; rowNum++)
            {
                IXLRow row = _sheet.Row(rowNum);
                if (row == null) continue;
                
                // 행마다 새 객체 생성
                YamlObject rowObj = OrderedYamlFactory.CreateObject();
                bool hasValues = false;
                
                // 각 자식 노드에 대해 처리
                foreach (var child in node.Children)
                {
                    string key = GetNodeKey(child, row);
                    if (!string.IsNullOrEmpty(key))
                    {
                        // PROPERTY 노드 처리
                        if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                        {
                            object value = child.GetValue(row);
                            if (value != null && !string.IsNullOrEmpty(value.ToString()))
                            {
                                rowObj.Add(key, FormatStringValue(value));
                                hasValues = true;
                            }
                        }
                        // MAP 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                        {
                            YamlObject childMap = OrderedYamlFactory.CreateObject();
                            AddChildProperties(child, childMap, row);
                            if (childMap.HasValues)
                            {
                                rowObj.Add(key, childMap);
                                hasValues = true;
                            }
                        }
                        // ARRAY 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            YamlArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues)
                            {
                                rowObj.Add(key, childArray);
                                hasValues = true;
                            }
                        }
                        // KEY 노드 처리 추가
                        else if (child.NodeType == SchemeNode.SchemeNodeType.KEY)
                        {
                            Logger.Debug($"ProcessArrayNode: KEY 노드 처리 - {key} (스키마 이름: {child.SchemeName})");
                            string dynamicKey = child.GetKey(row);
                            Logger.Debug($"ProcessArrayNode: 동적 키 값: '{dynamicKey}'");
                            
                            // KEY 노드에 대응하는 VALUE 노드 찾기
                            SchemeNode valueNode = child.Children.FirstOrDefault(c => c.NodeType == SchemeNode.SchemeNodeType.VALUE);
                            if (valueNode != null)
                            {
                                object value = valueNode.GetValue(row);
                                if (!string.IsNullOrEmpty(dynamicKey) && value != null && !string.IsNullOrEmpty(value.ToString()))
                                {
                                    Logger.Debug($"ProcessArrayNode: KEY-VALUE 쌍 추가 - {dynamicKey}={value}");
                                    // 동적으로 생성된 키를 사용
                                    rowObj.Add(dynamicKey, FormatStringValue(value));
                                    hasValues = true;
                                }
                            }
                            else
                            {
                                // VALUE 노드가 없는 경우 KEY 자체를 키로 사용
                                if (!string.IsNullOrEmpty(dynamicKey))
                                {
                                    // 키만 있고 값이 없는 경우 빈 값을 추가
                                    Logger.Debug($"ProcessArrayNode: 키만 추가 - {dynamicKey}");
                                    rowObj.Add(dynamicKey, "");
                                    hasValues = true;
                                }
                            }
                        }
                        // VALUE 노드 처리 (KEY의 자식이 아닌 경우)
                        else if (child.NodeType == SchemeNode.SchemeNodeType.VALUE)
                        {
                            if (child.Parent == null || child.Parent.NodeType != SchemeNode.SchemeNodeType.KEY)
                            {
                                Logger.Debug($"ProcessArrayNode: 독립 VALUE 노드 처리 - {key}");
                                object value = child.GetValue(row);
                                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                {
                                    rowObj.Add(key, FormatStringValue(value));
                                    hasValues = true;
                                    Logger.Debug($"ProcessArrayNode: VALUE 값 추가 - {key}={value}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // 키가 없는 경우의 처리
                        
                        if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                        {
                            // MAP 노드의 모든 자식을 직접 rowObj에 추가
                            AddChildProperties(child, rowObj, row);
                            hasValues = rowObj.HasValues;
                        }
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            // ARRAY 노드의 처리
                            YamlArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues && childArray.Count > 0 && childArray[0] is YamlObject firstObj)
                            {
                                foreach (var property in firstObj.Properties)
                                {
                                    rowObj.Add(property.Key, property.Value);
                                    hasValues = true;
                                }
                            }
                        }
                        else if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                        {
                            // PROPERTY 노드의 값을 직접 추가
                            object value = child.GetValue(row);
                            if (value != null && !string.IsNullOrEmpty(value.ToString()))
                            {
                                // 값이 있지만 키가 없는 경우, 기본 키를 사용하거나 처리 방식 결정
                                // 여기서는 값 자체를 별도 객체로 추가
                                YamlObject valueObj = OrderedYamlFactory.CreateObject();
                                valueObj.Add("value", value); // 기본 키 사용
                                for (int i = 0; i < valueObj.Properties.Count(); i++)
                                {
                                    var prop = valueObj.Properties.ElementAt(i);
                                    rowObj.Add(prop.Key, prop.Value);
                                    hasValues = true;
                                }
                            }
                        }
                        // KEY 노드 처리 추가 (키가 없는 경우)
                        else if (child.NodeType == SchemeNode.SchemeNodeType.KEY)
                        {
                            Logger.Debug($"ProcessArrayNode(key없음): KEY 노드 처리 - 스키마 이름: {child.SchemeName}");
                            string dynamicKey = child.GetKey(row);
                            Logger.Debug($"ProcessArrayNode(key없음): 동적 키 값: '{dynamicKey}'");
                            
                            if (!string.IsNullOrEmpty(dynamicKey))
                            {
                                // KEY 노드에 대응하는 VALUE 노드 찾기
                                SchemeNode valueNode = child.Children.FirstOrDefault(c => c.NodeType == SchemeNode.SchemeNodeType.VALUE);
                                if (valueNode != null)
                                {
                                    object value = valueNode.GetValue(row);
                                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                    {
                                        Logger.Debug($"ProcessArrayNode(key없음): KEY-VALUE 쌍 추가 - {dynamicKey}={value}");
                                        rowObj.Add(dynamicKey, FormatStringValue(value));
                                        hasValues = true;
                                    }
                                }
                                else
                                {
                                    // VALUE 노드가 없는 경우 빈 값으로 처리
                                    Logger.Debug($"ProcessArrayNode(key없음): 키만 추가 - {dynamicKey}");
                                    rowObj.Add(dynamicKey, "");
                                    hasValues = true;
                                }
                            }
                        }
                        // VALUE 노드 처리 추가 (키가 없는 경우, KEY의 자식이 아닌 경우)
                        else if (child.NodeType == SchemeNode.SchemeNodeType.VALUE)
                        {
                            if (child.Parent == null || child.Parent.NodeType != SchemeNode.SchemeNodeType.KEY)
                            {
                                Logger.Debug($"ProcessArrayNode(key없음): VALUE 노드 처리");
                                object value = child.GetValue(row);
                                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                {
                                    // 키가 없는 VALUE는 기본 키와 함께 추가
                                    Logger.Debug($"ProcessArrayNode(key없음): VALUE 값 추가 - {value}");
                                    rowObj.Add("value", FormatStringValue(value));
                                    hasValues = true;
                                }
                            }
                        }
                    }
                }
                
                // 비어있지 않은 객체만 추가
                if (hasValues)
                {
                    result.Add(rowObj);
                }
            }
            
            return result;
        }
        
        private YamlArray ProcessArrayItems(SchemeNode node, IXLRow row)
        {
            YamlArray result = OrderedYamlFactory.CreateArray();
            Logger.Debug($"ProcessArrayItems: 노드={node.Key}, 타입={node.NodeType}");
            
            // 직접 자식 노드가 있는 경우 처리
            if (node.Children.Any())
            {
                foreach (var child in node.Children)
                {
                    Logger.Debug($"배열 자식 처리: 자식={child.Key}, 타입={child.NodeType}");
                    
                    if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                    {
                        // PROPERTY 노드 처리
                        object value = child.GetValue(row);
                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            // 키가 있는 경우 객체로, 없는 경우 값으로 추가
                            string childKey = GetNodeKey(child, row);
                            if (!string.IsNullOrEmpty(childKey))
                            {
                                YamlObject childObj = OrderedYamlFactory.CreateObject();
                                childObj.Add(childKey, FormatStringValue(value));
                                result.Add(childObj);
                            }
                            else
                            {
                                result.Add(FormatStringValue(value));
                            }
                        }
                    }
                    else if (child.NodeType == SchemeNode.SchemeNodeType.KEY)
                    {
                        // KEY 노드 처리
                        Logger.Debug($"KEY 타입 노드 처리: {child.Key} (스키마 이름: {child.SchemeName})");
                        string keyValue = child.GetKey(row);
                        Logger.Debug($"KEY 노드의 실제 키 값: '{keyValue}'");
                        
                        // KEY 노드의 자식 중 VALUE 노드 찾기
                        SchemeNode valueNode = child.Children.FirstOrDefault(c => c.NodeType == SchemeNode.SchemeNodeType.VALUE);
                        if (valueNode != null)
                        {
                            // VALUE 노드가 있으면 키-값 쌍으로 처리
                            object value = valueNode.GetValue(row);
                            if (!string.IsNullOrEmpty(keyValue) && value != null && !string.IsNullOrEmpty(value.ToString()))
                            {
                                Logger.Debug($"KEY-VALUE 쌍 생성: '{keyValue}'={value}");
                                YamlObject keyValueObj = OrderedYamlFactory.CreateObject();
                                keyValueObj.Add(keyValue, FormatStringValue(value));
                                result.Add(keyValueObj);
                            }
                        }
                        else if (!string.IsNullOrEmpty(keyValue))
                        {
                            // VALUE 노드가 없으면 KEY 자체 값을 사용
                            // 키 값이 있으면 빈 값과 함께 객체로 추가
                            Logger.Debug($"KEY 값만 사용: '{keyValue}'");
                            YamlObject keyObj = OrderedYamlFactory.CreateObject();
                            keyObj.Add(keyValue, "");
                            result.Add(keyObj);
                        }
                    }
                    else if (child.NodeType == SchemeNode.SchemeNodeType.VALUE)
                    {
                        // VALUE 노드 처리 (부모가 KEY가 아닌 경우)
                        if (child.Parent == null || child.Parent.NodeType != SchemeNode.SchemeNodeType.KEY)
                        {
                            Logger.Debug($"독립 VALUE 타입 노드 처리: {child.Key}");
                            object value = child.GetValue(row);
                            if (value != null && !string.IsNullOrEmpty(value.ToString()))
                            {
                                // VALUE 타입은 배열 항목에 직접 값을 추가
                                Logger.Debug($"VALUE 값 추가: {value}");
                                result.Add(FormatStringValue(value));
                            }
                        }
                    }
                    else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                    {
                        // MAP 노드 처리
                        YamlObject childObj = OrderedYamlFactory.CreateObject();
                        AddChildProperties(child, childObj, row);
                        if (childObj.HasValues)
                        {
                            result.Add(childObj);
                        }
                    }
                    else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                    {
                        // 배열 노드 처리
                        YamlArray childArray = ProcessArrayItems(child, row);
                        if (childArray.HasValues)
                        {
                            // 배열의 각 항목을 결과 배열에 추가
                            for (int i = 0; i < childArray.Count; i++)
                            {
                                result.Add(childArray[i]);
                            }
                        }
                    }
                }
            }
            else
            {
                // 자식 노드가 없는 경우 기본 객체 추가
                YamlObject obj = OrderedYamlFactory.CreateObject();
                string key = GetNodeKey(node, row);
                object value = node.GetValue(row);
                
                if (!string.IsNullOrEmpty(key) && value != null && !string.IsNullOrEmpty(value.ToString()))
                {
                    obj.Add(key, value);
                    if (obj.HasValues)
                    {
                        result.Add(obj);
                    }
                }
            }
            
            return result;
        }
        
        private void AddChildProperties(SchemeNode node, YamlObject parent, IXLRow row)
        {
            foreach (var child in node.Children)
            {
                string key = GetNodeKey(child, row);
                if (string.IsNullOrEmpty(key)) continue;
                
                // PROPERTY 노드 처리
                if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                {
                    object value = child.GetValue(row);
                    if (value != null && !string.IsNullOrEmpty(value.ToString()))
                    {
                        parent.Add(key, value);
                    }
                }
                // MAP 노드 처리
                else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                {
                    YamlObject childMap = OrderedYamlFactory.CreateObject();
                    AddChildProperties(child, childMap, row);
                    if (childMap.HasValues)
                    {
                        parent.Add(key, childMap);
                    }
                }
                // ARRAY 노드 처리
                else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                {
                    YamlArray childArray = ProcessArrayItems(child, row);
                    if (childArray.HasValues)
                    {
                        parent.Add(key, childArray);
                    }
                }
            }
        }
        
        private string GetNodeKey(SchemeNode node, IXLRow row)
        {
            string key = node.Key;
            if (node.IsKeyProvidable)
            {
                string rowKey = node.GetKey(row);
                if (!string.IsNullOrEmpty(rowKey))
                {
                    key = rowKey;
                }
            }
            return key;
        }
        
        private bool RemoveEmptyAttributes(object arg)
        {
            bool valueExist = false;
            
            if (arg is string str)
            {
                valueExist = !string.IsNullOrEmpty(str);
            }
            else if (arg is int || arg is long || arg is float || arg is double || arg is decimal)
            {
                valueExist = true;
            }
            else if (arg is bool)
            {
                valueExist = true;
            }
            else if (arg is YamlObject yamlObject)
            {
                var keysToRemove = new List<string>();
                
                foreach (var property in yamlObject.Properties)
                {
                    if (!RemoveEmptyAttributes(property.Value))
                    {
                        keysToRemove.Add(property.Key);
                    }
                    else
                    {
                        valueExist = true;
                    }
                }
                
                foreach (var key in keysToRemove)
                {
                    yamlObject.Remove(key);
                }
            }
            else if (arg is YamlArray yamlArray)
            {
                for (int i = 0; i < yamlArray.Count; i++)
                {
                    var item = yamlArray[i];
                    if (!RemoveEmptyAttributes(item))
                    {
                        yamlArray.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        valueExist = true;
                    }
                }
            }
            
            return valueExist;
        }

        /// <summary>
        /// 문자열 값을 YAML 형식에 맞게 처리합니다.
        /// 한글이 포함된 경우와 개행 문자가 있는 경우를 처리합니다.
        /// </summary>
        /// <param name="value">처리할 문자열 값</param>
        /// <returns>처리된 문자열 값</returns>
        private object FormatStringValue(object value)
        {
            if (value == null)
                return null;
                
            string strValue = value.ToString();
            if (string.IsNullOrEmpty(strValue))
                return strValue;
                
            // 개행 문자 포함 여부 확인 및 처리
            if (strValue.Contains('\n') || strValue.Contains('\r'))
            {
                // 개행 문자가 이미 이스케이프되어 있는지 확인
                if (!strValue.Contains("\\n") && !strValue.Contains("\\r"))
                {
                    // 새 줄 문자를 이스케이프 처리하지 않고 보존하면
                    // YAML에서 자동으로 블록 스타일을 적용하므로 그대로 반환
                    return strValue;
                }
            }
            
            return strValue;
        }

        // YAML 객체 생성을 위한 메서드
        public object ProcessRootNode()
        {
            SchemeNode rootNode = _scheme.Root;
            Logger.Debug("루트 노드 처리: 타입={0}", rootNode.NodeType);
            
            if (rootNode.NodeType == SchemeNode.SchemeNodeType.MAP)
            {
                return ProcessMapNode(rootNode);
            }
            else if (rootNode.NodeType == SchemeNode.SchemeNodeType.ARRAY)
            {
                return ProcessArrayNode(rootNode);
            }
            else
            {
                Logger.Warning("지원하지 않는 루트 노드 타입: {0}", rootNode.NodeType);
                return OrderedYamlFactory.CreateObject(); // 기본 빈 객체 반환
            }
        }
    }
} 