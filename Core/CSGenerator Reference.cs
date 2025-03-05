using Microsoft.Extensions.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelToJsonCSharp.Config;

namespace ExcelToJsonCSharp.Core
{
    public class YamlGenerator
    {
        private static readonly ILogger<YamlGenerator> Logger = LoggerFactory.Create(builder => builder.AddConsole()).CreateLogger<YamlGenerator>();

        private readonly Scheme scheme;
        private readonly IXLWorksheet sheet;

        public YamlGenerator(Scheme scheme)
        {
            this.scheme = scheme;
            this.sheet = scheme.Sheet;
            Logger.LogInformation("YamlGenerator 초기화: 내용 시작 행={ContentStartRow}", scheme.ContentStartRowNum);
        }

        public string Generate(YamlStyle style = YamlStyle.Block, int indentSize = 2, bool preserveQuotes = false)
        {
            SchemeNode rootNode = scheme.Root;
            Logger.LogInformation("YAML 생성 시작");
            Logger.LogInformation("루트 노드: {Key}, 타입={Type}", rootNode.Key, rootNode.NodeType);
            
            object rootYaml;
            if (rootNode.NodeType == SchemeNode.SchemeNodeType.MAP)
            {
                Logger.LogInformation("MAP 루트 노드 처리");
                // MAP 노드를 직접 처리하지 않고, rootNode의 자식들을 직접 처리하도록 수정
                var rootMapping = ProcessMapNode(rootNode);
                rootYaml = rootMapping;
            }
            else if (rootNode.NodeType == SchemeNode.SchemeNodeType.ARRAY)
            {
                Logger.LogInformation("ARRAY 루트 노드 처리");
                // 루트 배열 노드의 항목들을 직접 추출
                YamlArray array = OrderedYamlFactory.CreateArray();
                
                // 모든 데이터 행에 대해 처리
                for (int rowNum = scheme.ContentStartRowNum; rowNum <= scheme.EndRowNum; rowNum++)
                {
                    IXLRow row = sheet.Row(rowNum);
                    if (row == null) continue;
                    
                    // 행마다 새 객체 생성
                    YamlObject rowObj = OrderedYamlFactory.CreateObject();
                    bool hasValues = false;
                    
                    // 각 자식 노드에 대해 처리
                    foreach (var child in rootNode.Children)
                    {
                        string key = GetNodeKey(child, row);
                        
                        // PROPERTY 노드 처리
                        if (child.NodeType == SchemeNode.SchemeNodeType.PROPERTY)
                        {
                            object value = child.GetValue(row);
                            if (value != null && !string.IsNullOrEmpty(value.ToString()))
                            {
                                if (!string.IsNullOrEmpty(key))
                                {
                                    rowObj.Add(key, value);
                                    hasValues = true;
                                }
                            }
                        }
                        // MAP 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                        {
                            YamlObject childMap = OrderedYamlFactory.CreateObject();
                            AddChildProperties(child, childMap, row);
                            if (childMap.HasValues)
                            {
                                if (!string.IsNullOrEmpty(key))
                                {
                                    rowObj.Add(key, childMap);
                                    hasValues = true;
                                }
                                else
                                {
                                    // 키가 없는 경우 맵의 속성들을 직접 추가
                                    foreach (var property in childMap.Properties)
                                    {
                                        rowObj.Add(property.Key, property.Value);
                                        hasValues = true;
                                    }
                                }
                            }
                        }
                        // ARRAY 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            YamlArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues)
                            {
                                if (!string.IsNullOrEmpty(key))
                                {
                                    rowObj.Add(key, childArray);
                                    hasValues = true;
                                }
                                else
                                {
                                    // 키가 없는 경우 배열의 항목들을 처리
                                    for (int i = 0; i < childArray.Count; i++)
                                    {
                                        if (childArray[i] is YamlObject obj)
                                        {
                                            foreach (var property in obj.Properties)
                                            {
                                                rowObj.Add(property.Key, property.Value);
                                                hasValues = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    Logger.LogDebug("행 {RowNum} 처리 결과: 유효한 값={HasValues}", rowNum, hasValues);
                    
                    // 비어있지 않은 객체만 추가
                    if (hasValues)
                    {
                        array.Add(rowObj);
                    }
                }
                
                rootYaml = array;
            }
            else
            {
                Logger.LogError("지원되지 않는 루트 노드 타입: {Type}", rootNode.NodeType);
                throw new InvalidOperationException("Illegal root yaml node type. must be unnamed map or array");
            }
            
            RemoveEmptyAttributes(rootYaml);
            return OrderedYamlFactory.SerializeToYaml(rootYaml, indentSize, style, preserveQuotes);
        }
        
        private YamlObject ProcessMapNode(SchemeNode node)
        {
            YamlObject result = OrderedYamlFactory.CreateObject();
            
            // 모든 데이터 행에 대해 처리
            for (int rowNum = scheme.ContentStartRowNum; rowNum <= scheme.EndRowNum; rowNum++)
            {
                IXLRow row = sheet.Row(rowNum);
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
            for (int rowNum = scheme.ContentStartRowNum; rowNum <= scheme.EndRowNum; rowNum++)
            {
                IXLRow row = sheet.Row(rowNum);
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
    }
} 