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
                
                foreach (var prop in obj.Properties)
                {
                    if (IsEmpty(prop.Value))
                    {
                        propertiesToRemove.Add(prop.Key);
                    }
                    else
                    {
                        RemoveEmptyProperties(prop.Value);
                    }
                }
                
                foreach (var propName in propertiesToRemove)
                {
                    obj.Remove(propName);
                }
            }
            else if (token is YamlArray array)
            {
                for (int i = array.Count - 1; i >= 0; i--)
                {
                    if (IsEmpty(array[i]))
                    {
                        array.RemoveAt(i);
                    }
                    else
                    {
                        RemoveEmptyProperties(array[i]);
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
                
            if (token is YamlArray array && !array.HasValues)
                return true;
                
            return false;
        }

        public static string SerializeToYaml(object obj, int indentSize = 2, YamlStyle style = YamlStyle.Block, bool preserveQuotes = false)
        {
            var sb = new StringBuilder();
            SerializeObject(obj, sb, 0, indentSize, style, preserveQuotes);
            return sb.ToString();
        }
        
        public static void SaveToYaml(object obj, string filePath, int indentSize = 2, YamlStyle style = YamlStyle.Block, bool preserveQuotes = false)
        {
            string yaml = SerializeToYaml(obj, indentSize, style, preserveQuotes);
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
            
            bool needQuotes = preserveQuotes || 
                              value.Contains(':') || 
                              value.Contains('#') || 
                              value.Contains(',') ||
                              value.StartsWith(" ") || 
                              value.EndsWith(" ") ||
                              value == "true" || 
                              value == "false" || 
                              value == "null" ||
                              (value.Length > 0 && char.IsDigit(value[0]));
                              
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
                        case '\n': sb.Append("\\n"); break;
                        case '\r': sb.Append("\\r"); break;
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
                    
                    if (kvp.Value is YamlObject || kvp.Value is YamlArray)
                    {
                        SerializeObject(kvp.Value, sb, level + 1, indentSize, style, preserveQuotes);
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
                        // 객체의 속성들을 현재 줄에서 시작
                        bool isFirst = true;
                        bool hasProcessedFirstProperty = false;
                        
                        foreach (var prop in yamlObj.Properties)
                        {
                            if (isFirst)
                            {
                                // 첫 번째 속성은 현재 줄에 표시
                                sb.Append(prop.Key).Append(": ");
                                
                                if (prop.Value is YamlObject || prop.Value is YamlArray)
                                {
                                    // 복잡한 값의 경우 다음 줄에 시작
                                    SerializeObject(prop.Value, sb, level + 1, indentSize, style, preserveQuotes);
                                }
                                else
                                {
                                    // 단순 값은 현재 줄에 표시
                                    SerializeObject(prop.Value, sb, level, indentSize, style, preserveQuotes);
                                    sb.AppendLine();
                                }
                                
                                isFirst = false;
                                hasProcessedFirstProperty = true;
                            }
                            else
                            {
                                // 두 번째 이후 속성은 새 줄에 표시
                                Indent(sb, level + 1, indentSize);
                                sb.Append(prop.Key).Append(": ");
                                
                                if (prop.Value is YamlObject || prop.Value is YamlArray)
                                {
                                    SerializeObject(prop.Value, sb, level + 2, indentSize, style, preserveQuotes);
                                }
                                else
                                {
                                    SerializeObject(prop.Value, sb, level + 1, indentSize, style, preserveQuotes);
                                    sb.AppendLine();
                                }
                            }
                        }
                        
                        // 속성이 없는 경우 빈 객체로 표시
                        if (!hasProcessedFirstProperty)
                        {
                            sb.AppendLine("{}");
                        }
                    }
                    else if (item is YamlArray)
                    {
                        SerializeObject(item, sb, level + 1, indentSize, style, preserveQuotes);
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