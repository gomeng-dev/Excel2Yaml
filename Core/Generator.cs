using ExcelToJsonAddin.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.IO;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public class Generator
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<Generator>();

        private readonly Scheme _scheme;
        private readonly IXLWorksheet _sheet;

        public Generator(Scheme scheme)
        {
            _scheme = scheme;
            _sheet = scheme.Sheet;
            Logger.Debug("Generator 초기화: 스키마 노드 타입={0}", scheme.Root.NodeType);
        }

        // 외부에서 호출할 정적 메서드 추가
        public static string GenerateJson(Scheme scheme, bool includeEmptyOptionals)
        {
            var generator = new Generator(scheme);
            Logger.Debug("JSON 생성 시작: 스키마 노드 타입={0}", scheme.Root.NodeType);
            var result = generator.Generate();
            
            if (!includeEmptyOptionals)
            {
                try
                {
                    // 빈 속성 제거 로직 수정
                    // 직접 JsonObject를 생성하고 처리
                    JsonObject jsonObj = OrderedJsonFactory.NewJsonObject(result);
                    if (jsonObj != null)
                    {
                        OrderedJsonFactory.RemoveEmptyProperties(jsonObj);
                        result = OrderedJsonFactory.SerializeObject(jsonObj);
                    }
                    else
                    {
                        Logger.Warning("JSON 객체 생성 실패: 원본 JSON 반환");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "JSON 처리 중 오류 발생");
                    // 오류 발생 시 원본 결과 반환
                }
            }
            
            return result;
        }

        public string Generate()
        {
            SchemeNode rootNode = _scheme.Root;
            Logger.Debug("JSON 생성 시작");
            Logger.Information("루트 노드: {0}, 타입={1}", rootNode.Key, rootNode.NodeType);
            
            object rootJson;
            if (rootNode.NodeType == SchemeNode.SchemeNodeType.MAP)
            {
                Logger.Information("MAP 루트 노드 처리");
                // MAP 노드를 직접 처리하지 않고, rootNode의 자식들을 직접 처리하도록 수정
                var rootObject = ProcessMapNode(rootNode);
                rootJson = rootObject;
            }
            else if (rootNode.NodeType == SchemeNode.SchemeNodeType.ARRAY)
            {
                Logger.Information("ARRAY 루트 노드 처리");
                // 루트 배열 노드의 항목들을 직접 추출
                JsonArray array = OrderedJsonFactory.CreateArray();
                
                // 모든 데이터 행에 대해 처리
                for (int rowNum = _scheme.ContentStartRowNum; rowNum <= _scheme.EndRowNum; rowNum++)
                {
                    IXLRow row = _sheet.Row(rowNum);
                    if (row == null) continue;
                    
                    // 행마다 새 객체 생성
                    JsonObject rowObj = OrderedJsonFactory.CreateObject();
                    
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
                                }
                            }
                        }
                        // MAP 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                        {
                            JsonObject childMap = OrderedJsonFactory.CreateObject();
                            AddChildProperties(child, childMap, row);
                            if (childMap.HasValues)
                            {
                                if (!string.IsNullOrEmpty(key))
                                {
                                    rowObj.Add(key, childMap);
                                }
                                else
                                {
                                    // 키가 없는 경우 맵의 속성들을 직접 추가
                                    foreach (var property in childMap.Properties)
                                    {
                                        rowObj.Add(property.Key, property.Value);
                                    }
                                }
                            }
                        }
                        // ARRAY 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            JsonArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues)
                            {
                                if (!string.IsNullOrEmpty(key))
                                {
                                    rowObj.Add(key, childArray);
                                }
                                else
                                {
                                    // 키가 없는 경우 배열의 항목들을 처리
                                    for (int i = 0; i < childArray.Count; i++)
                                    {
                                        if (childArray[i] is JsonObject obj)
                                        {
                                            foreach (var property in obj.Properties)
                                            {
                                                rowObj.Add(property.Key, property.Value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    Logger.Debug("행 {0} 처리 결과: 유효한 값={1}", rowNum, rowObj.HasValues);
                    
                    // 비어있지 않은 객체만 추가
                    if (rowObj.HasValues)
                    {
                        array.Add(rowObj);
                    }
                }
                
                rootJson = array;
            }
            else
            {
                Logger.Error("지원되지 않는 루트 노드 타입: {0}", rootNode.NodeType);
                throw new InvalidOperationException("Illegal root json node type. must be unnamed map or array");
            }
            
            RemoveEmptyAttributes(rootJson);
            return OrderedJsonFactory.SerializeObject(rootJson, true);
        }
        
        private JsonObject ProcessMapNode(SchemeNode node)
        {
            JsonObject result = OrderedJsonFactory.CreateObject();
            
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
                            JsonObject childMap = OrderedJsonFactory.CreateObject();
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
                            JsonArray childArray = ProcessArrayItems(child, row);
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
        
        private JsonArray ProcessArrayNode(SchemeNode node)
        {
            JsonArray result = OrderedJsonFactory.CreateArray();
            
            // 모든 데이터 행에 대해 처리
            for (int rowNum = _scheme.ContentStartRowNum; rowNum <= _scheme.EndRowNum; rowNum++)
            {
                IXLRow row = _sheet.Row(rowNum);
                if (row == null) continue;
                
                // 행마다 새 객체 생성
                JsonObject rowObj = OrderedJsonFactory.CreateObject();
                
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
                            }
                        }
                        // MAP 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.MAP)
                        {
                            JsonObject childMap = OrderedJsonFactory.CreateObject();
                            AddChildProperties(child, childMap, row);
                            if (childMap.HasValues)
                            {
                                rowObj.Add(key, childMap);
                            }
                        }
                        // ARRAY 노드 처리
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            JsonArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues)
                            {
                                rowObj.Add(key, childArray);
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
                        }
                        else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                        {
                            // ARRAY 노드의 처리
                            JsonArray childArray = ProcessArrayItems(child, row);
                            if (childArray.HasValues && childArray.Count > 0 && childArray[0] is JsonObject firstObj)
                            {
                                foreach (var property in firstObj.Properties)
                                {
                                    rowObj.Add(property.Key, property.Value);
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
                                JsonObject valueObj = OrderedJsonFactory.CreateObject();
                                valueObj.Add("value", value); // 기본 키 사용
                                for (int i = 0; i < valueObj.Properties.Count(); i++)
                                {
                                    var prop = valueObj.Properties.ElementAt(i);
                                    rowObj.Add(prop.Key, prop.Value);
                                }
                            }
                        }
                    }
                }
                
                // 비어있지 않은 객체만 추가
                if (rowObj.HasValues)
                {
                    result.Add(rowObj);
                }
            }
            
            return result;
        }
        
        private JsonArray ProcessArrayItems(SchemeNode node, IXLRow row)
        {
            JsonArray result = OrderedJsonFactory.CreateArray();
            
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
                                JsonObject childObj = OrderedJsonFactory.CreateObject();
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
                        JsonObject childObj = OrderedJsonFactory.CreateObject();
                        AddChildProperties(child, childObj, row);
                        if (childObj.HasValues)
                        {
                            result.Add(childObj);
                        }
                    }
                    else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                    {
                        // 배열 노드 처리
                        JsonArray childArray = ProcessArrayItems(child, row);
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
                JsonObject obj = OrderedJsonFactory.CreateObject();
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
        
        private void AddChildProperties(SchemeNode node, JsonObject parent, IXLRow row)
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
                    JsonObject childMap = OrderedJsonFactory.CreateObject();
                    AddChildProperties(child, childMap, row);
                    if (childMap.HasValues)
                    {
                        parent.Add(key, childMap);
                    }
                }
                // ARRAY 노드 처리
                else if (child.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                {
                    JsonArray childArray = ProcessArrayItems(child, row);
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
            else if (arg is JsonObject jsonObject)
            {
                var keysToRemove = new List<string>();
                
                foreach (var property in jsonObject.Properties)
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
                    jsonObject.Remove(key);
                }
            }
            else if (arg is JsonArray jsonArray)
            {
                for (int i = 0; i < jsonArray.Count; i++)
                {
                    var item = jsonArray[i];
                    if (!RemoveEmptyAttributes(item))
                    {
                        jsonArray.RemoveAt(i);
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

        // 워크시트 변환 메서드 추가
        public string ConvertWorksheet(Microsoft.Office.Interop.Excel.Worksheet worksheet, Config.ExcelToJsonConfig config)
        {
            try
            {
                // 임시 파일로 저장
                string tempFile = Path.GetTempFileName() + ".xlsx";
                var app = worksheet.Application;
                var workbook = worksheet.Parent as Microsoft.Office.Interop.Excel.Workbook;
                
                // 현재 워크북을 임시 파일로 저장
                workbook.SaveCopyAs(tempFile);
                
                // ClosedXML로 워크시트 열기
                using (var wb = new XLWorkbook(tempFile))
                {
                    var sheet = wb.Worksheets.First();
                    var scheme = new Scheme(sheet);
                    
                    // 출력 형식에 따라 변환
                    string result;
                    if (config.OutputFormat == Config.OutputFormat.Json)
                    {
                        result = GenerateJson(scheme, config.IncludeEmptyFields);
                    }
                    else
                    {
                        result = YamlGenerator.Generate(
                            scheme, 
                            config.YamlStyle, 
                            config.YamlIndentSize, 
                            config.YamlPreserveQuotes, 
                            config.IncludeEmptyFields
                        );
                    }
                    
                    // 임시 파일 삭제
                    try { File.Delete(tempFile); } catch { /* 무시 */ }
                    
                    return result;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "워크시트 변환 중 오류 발생");
                throw;
            }
        }
    }
}
