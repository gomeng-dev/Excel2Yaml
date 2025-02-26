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
                                rowObj.Add(key, value);
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
            
            // 직접 자식 노드가 있는 경우 처리
            if (node.Children.Any())
            {
                foreach (var child in node.Children)
                {
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
                                childObj.Add(childKey, value);
                                result.Add(childObj);
                            }
                            else
                            {
                                result.Add(value);
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