using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelToJsonAddin.Config;
using System.Linq;

namespace ExcelToJsonAddin.Core
{
    public class YamlObject
    {
        private readonly Dictionary<string, object> properties = new Dictionary<string, object>();
        private readonly List<string> propertyOrder = new List<string>();

        public void Add(string name, object value)
        {
            if (properties.ContainsKey(name))
            {
                properties[name] = value;
            }
            else
            {
                properties.Add(name, value);
                propertyOrder.Add(name);
            }
        }

        public bool ContainsKey(string key)
        {
            return properties.ContainsKey(key);
        }

        public object this[string key]
        {
            get { return properties[key]; }
        }

        public void Remove(string key)
        {
            if (properties.ContainsKey(key))
            {
                properties.Remove(key);
                propertyOrder.Remove(key);
            }
        }

        public bool HasValues => properties.Count > 0;

        public IEnumerable<KeyValuePair<string, object>> Properties
        {
            get
            {
                foreach (var key in propertyOrder)
                {
                    yield return new KeyValuePair<string, object>(key, properties[key]);
                }
            }
        }
    }

    public class YamlArray
    {
        private readonly List<object> items = new List<object>();

        public void Add(object value)
        {
            items.Add(value);
        }

        public void RemoveAt(int index)
        {
            if (index >= 0 && index < items.Count)
            {
                items.RemoveAt(index);
            }
        }

        public object this[int index]
        {
            get { return items[index]; }
        }

        public int Count => items.Count;

        public bool HasValues => items.Count > 0;

        public IEnumerable<object> Items => items;
    }

    public static class OrderedYamlFactory
    {
        public static YamlObject CreateObject() => new YamlObject();
        public static YamlArray CreateArray() => new YamlArray();

        public static void RemoveEmptyProperties(object token)
        {
            if (token is YamlObject obj)
            {
                var propertiesToRemove = new List<string>();
                
                // 먼저 모든 자식 속성을 처리
                foreach (var prop in obj.Properties)
                {
                    RemoveEmptyProperties(prop.Value);
                }
                
                // 그 다음 빈 속성 확인 및 제거
                foreach (var prop in obj.Properties)
                {
                    if (IsEmpty(prop.Value))
                    {
                        propertiesToRemove.Add(prop.Key);
                    }
                }
                
                foreach (var propName in propertiesToRemove)
                {
                    obj.Remove(propName);
                }
            }
            else if (token is YamlArray array)
            {
                // 먼저 모든 배열 항목 처리
                for (int i = 0; i < array.Count; i++)
                {
                    RemoveEmptyProperties(array[i]);
                }
                
                // 그 다음 빈 항목 제거 (뒤에서부터 제거해야 인덱스가 유효함)
                for (int i = array.Count - 1; i >= 0; i--)
                {
                    if (IsEmpty(array[i]))
                    {
                        array.RemoveAt(i);
                    }
                }
            }
        }

        private static bool IsEmpty(object token)
        {
            if (token == null)
                return true;
                
            if (token is string str && string.IsNullOrEmpty(str))
                return true;
                
            if (token is YamlObject obj && !obj.HasValues)
                return true;
                
            if (token is YamlArray array)
            {
                if (!array.HasValues)
                    return true;
                    
                // 추가: 배열의 모든 항목이 빈 경우 전체 배열을 빈 것으로 간주
                bool allItemsEmpty = true;
                foreach (var item in array.Items)
                {
                    if (!IsEmpty(item))
                    {
                        allItemsEmpty = false;
                        break;
                    }
                }
                return allItemsEmpty;
            }
                
            return false;
        }

        public static string SerializeToYaml(object obj, int indentSize = 2, YamlStyle style = YamlStyle.Block, bool preserveQuotes = false, bool includeEmptyFields = false)
        {
            // includeEmptyFields가 true일 경우 빈 속성 유지, false일 경우 제거
            if (!includeEmptyFields)
            {
                RemoveEmptyProperties(obj);
            }
            
            var sb = new StringBuilder();
            SerializeObject(obj, sb, 0, indentSize, style, preserveQuotes);
            return sb.ToString();
        }
        
        public static void SaveToYaml(object obj, string filePath, int indentSize = 2, YamlStyle style = YamlStyle.Block, bool preserveQuotes = false, bool includeEmptyFields = false)
        {
            string yaml = SerializeToYaml(obj, indentSize, style, preserveQuotes, includeEmptyFields);
            File.WriteAllText(filePath, yaml);
        }

        private static void SerializeObject(object obj, StringBuilder sb, int level, int indentSize, YamlStyle style, bool preserveQuotes)
        {
            if (obj == null)
            {
                sb.Append("null");
                return;
            }
            
            if (obj is string s)
            {
                SerializeString(s, sb, preserveQuotes);
                return;
            }
            
            if (obj is int || obj is long || obj is float || obj is double || obj is decimal)
            {
                sb.Append(Convert.ToString(obj));
                return;
            }
            
            if (obj is bool b)
            {
                sb.Append(b ? "true" : "false");
                return;
            }
            
            if (obj is YamlObject yamlObj)
            {
                SerializeYamlObject(yamlObj, sb, level, indentSize, style, preserveQuotes);
                return;
            }
            
            if (obj is YamlArray yamlArray)
            {
                SerializeYamlArray(yamlArray, sb, level, indentSize, style, preserveQuotes);
                return;
            }
            
            // 기타 타입은 문자열로 변환
            SerializeString(obj.ToString(), sb, preserveQuotes);
        }
        
        private static void SerializeString(string value, StringBuilder sb, bool preserveQuotes)
        {
            if (string.IsNullOrEmpty(value))
            {
                sb.Append(preserveQuotes ? "\"\"" : "");
                return;
            }
            
            // 숫자로만 이루어진 문자열인지 확인
            bool isNumericString = !string.IsNullOrEmpty(value) && value.All(char.IsDigit);
            
            // 한글 문자 포함 여부 확인
            bool containsKorean = false;
            foreach (char c in value)
            {
                if (c >= '\uAC00' && c <= '\uD7A3')  // 한글 유니코드 범위 (가-힣)
                {
                    containsKorean = true;
                    break;
                }
            }
            
            bool needQuotes = preserveQuotes || 
                              containsKorean ||  // 한글 포함 시 따옴표 추가
                              value.Contains(':') || 
                              value.Contains('#') || 
                              value.Contains(',') ||
                              value.Contains('_') ||
                              value.StartsWith(" ") || 
                              value.EndsWith(" ") ||
                              value == "true" || 
                              value == "false" || 
                              value == "null" ||
                              (value.Length > 0 && char.IsDigit(value[0]) && !isNumericString);
                              
            // 개행 문자 포함 여부 확인
            bool containsNewline = value.Contains('\n') || value.Contains('\r');
            if (containsNewline)
            {
                needQuotes = true;  // 개행 포함 시 무조건 따옴표 필요
            }
            
            if (needQuotes)
            {
                sb.Append('"');
                foreach (char c in value)
                {
                    switch (c)
                    {
                        case '"':
                            sb.Append("\\\"");
                            break;
                        case '\\':
                            sb.Append("\\\\");
                            break;
                        case '\n': sb.Append("\\n"); break;  // 항상 개행을 \n으로 치환
                        case '\r': sb.Append("\\r"); break;  // 항상 캐리지리턴을 \r로 치환
                        case '\t': sb.Append("\\t"); break;
                        default: sb.Append(c.ToString()); break;
                    }
                }
                sb.Append('"');
            }
            else
            {
                sb.Append(value);
            }
        }
        
        private static void SerializeYamlObject(YamlObject obj, StringBuilder sb, int level, int indentSize, YamlStyle style, bool preserveQuotes)
        {
            if (!obj.HasValues)
            {
                sb.Append("{}");
                if (style == YamlStyle.Block)
                {
                    sb.AppendLine();
                }
                return;
            }
            
            bool isFirst = true;
            
            if (style == YamlStyle.Flow)
            {
                sb.Append('{');
                
                foreach (var kvp in obj.Properties)
                {
                    if (!isFirst)
                    {
                        sb.Append(", ");
                    }
                    
                    sb.Append(kvp.Key).Append(": ");
                    SerializeObject(kvp.Value, sb, level + 1, indentSize, style, preserveQuotes);
                    isFirst = false;
                }
                
                sb.Append('}');
            }
            else // Block 스타일
            {
                if (level > 0)
                {
                    sb.AppendLine();
                }
                
                foreach (var kvp in obj.Properties)
                {
                    if (!isFirst || level > 0)
                    {
                        Indent(sb, level, indentSize);
                    }
                    
                    sb.Append(kvp.Key).Append(": ");
                    
                    // 빈 배열인 경우 바로 개행 처리
                    if (kvp.Value is YamlArray yamlArray && !yamlArray.HasValues)
                    {
                        sb.AppendLine("[]");
                    }
                    else if (kvp.Value is YamlObject || kvp.Value is YamlArray)
                    {
                        // MAP 노드의 자식들은 2 레벨 더 들여쓰기
                        SerializeObject(kvp.Value, sb, level + 2, indentSize, style, preserveQuotes);
                    }
                    else
                    {
                        SerializeObject(kvp.Value, sb, level, indentSize, style, preserveQuotes);
                        sb.AppendLine();
                    }
                    
                    isFirst = false;
                }
            }
        }
        
        private static void SerializeYamlArray(YamlArray array, StringBuilder sb, int level, int indentSize, YamlStyle style, bool preserveQuotes)
        {
            if (!array.HasValues)
            {
                sb.Append("[]");
                if (style == YamlStyle.Block)
                {
                    sb.AppendLine();
                }
                return;
            }
            
            if (style == YamlStyle.Flow)
            {
                sb.Append('[');
                bool isFirst = true;
                
                foreach (var item in array.Items)
                {
                    if (!isFirst)
                    {
                        sb.Append(", ");
                    }
                    
                    SerializeObject(item, sb, level + 1, indentSize, style, preserveQuotes);
                    isFirst = false;
                }
                
                sb.Append(']');
            }
            else // Block 스타일
            {
                if (level > 0)
                {
                    sb.AppendLine();
                }
                
                foreach (var item in array.Items)
                {
                    Indent(sb, level, indentSize);
                    sb.Append("- ");
                    
                    if (item is YamlObject yamlObj)
                    {
                        // 객체 내 속성들 처리
                        if (!yamlObj.HasValues)
                        {
                            sb.AppendLine("{}");
                        }
                        else
                        {
                            // 첫 번째 속성을 "- " 다음에 바로 표시
                            bool isFirstProperty = true;
                            
                            foreach (var prop in yamlObj.Properties)
                            {
                                if (isFirstProperty)
                                {
                                    // 첫 번째 속성은 같은 줄에 표시
                                    sb.Append(prop.Key).Append(": ");
                                    
                                    if (prop.Value is YamlObject || prop.Value is YamlArray)
                                    {
                                        // 복잡한 값은 다음 줄에 표시
                                        SerializeObject(prop.Value, sb, level + 2, indentSize, style, preserveQuotes);
                                    }
                                    else
                                    {
                                        // 단순 값은 같은 줄에 표시
                                        SerializeObject(prop.Value, sb, level, indentSize, style, preserveQuotes);
                                        sb.AppendLine();
                                    }
                                    
                                    isFirstProperty = false;
                                }
                                else
                                {
                                    // 두 번째 이후 속성은 새 줄에 들여쓰기 적용하여 표시
                                    Indent(sb, level + 1, indentSize);
                                    sb.Append(prop.Key).Append(": ");
                                    
                                    if (prop.Value is YamlObject || prop.Value is YamlArray)
                                    {
                                        // MAP 노드의 자식들은 2 레벨 더 들여쓰기
                                        SerializeObject(prop.Value, sb, level + 2, indentSize, style, preserveQuotes);
                                    }
                                    else
                                    {
                                        SerializeObject(prop.Value, sb, level + 1, indentSize, style, preserveQuotes);
                                        sb.AppendLine();
                                    }
                                }
                            }
                        }
                    }
                    else if (item is YamlArray)
                    {
                        // 배열의 자식들도 2 레벨 더 들여쓰기
                        SerializeObject(item, sb, level + 2, indentSize, style, preserveQuotes);
                    }
                    else
                    {
                        SerializeObject(item, sb, level, indentSize, style, preserveQuotes);
                        sb.AppendLine();
                    }
                }
            }
        }
        
        private static void Indent(StringBuilder sb, int level, int indentSize)
        {
            for (int i = 0; i < level * indentSize; i++)
            {
                sb.Append(' ');
            }
        }
    }
}