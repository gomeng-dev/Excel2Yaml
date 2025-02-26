using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToJsonAddin.Core
{
    public class JsonObject
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

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append('{');
            
            bool first = true;
            foreach (var key in propertyOrder)
            {
                if (properties.ContainsKey(key))
                {
                    if (!first)
                    {
                        sb.Append(',');
                    }
                    first = false;
                    
                    sb.Append('"').Append(OrderedJsonFactory.EscapeJsonString(key)).Append('"').Append(':');
                    OrderedJsonFactory.AppendValue(sb, properties[key]);
                }
            }
            
            sb.Append('}');
            return sb.ToString();
        }
    }

    public class JsonArray
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

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append('[');
            
            bool first = true;
            foreach (var item in items)
            {
                if (!first)
                {
                    sb.Append(',');
                }
                first = false;
                
                OrderedJsonFactory.AppendValue(sb, item);
            }
            
            sb.Append(']');
            return sb.ToString();
        }
    }

    public static class OrderedJsonFactory
    {
        public static JsonObject CreateObject() => new JsonObject();
        public static JsonArray CreateArray() => new JsonArray();

        public static void AddProperty(JsonObject obj, string name, object value)
        {
            obj.Add(name, value);
        }

        public static void RemoveEmptyProperties(object token)
        {
            if (token is JsonObject obj)
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
            else if (token is JsonArray array)
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
                
            if (token is JsonObject obj && !obj.HasValues)
                return true;
                
            if (token is JsonArray array && !array.HasValues)
                return true;
                
            return false;
        }

        public static void AppendValue(StringBuilder sb, object value)
        {
            if (value == null)
            {
                sb.Append("null");
            }
            else if (value is string str)
            {
                sb.Append('"').Append(EscapeJsonString(str)).Append('"');
            }
            else if (value is bool b)
            {
                sb.Append(b ? "true" : "false");
            }
            else if (value is int || value is long || value is float || value is double || value is decimal)
            {
                sb.Append(Convert.ToString(value));
            }
            else if (value is JsonObject obj)
            {
                sb.Append(obj.ToString());
            }
            else if (value is JsonArray array)
            {
                sb.Append(array.ToString());
            }
            else
            {
                sb.Append('"').Append(EscapeJsonString(value.ToString())).Append('"');
            }
        }

        public static string EscapeJsonString(string str)
        {
            if (string.IsNullOrEmpty(str))
                return string.Empty;

            var sb = new StringBuilder();
            foreach (char c in str)
            {
                switch (c)
                {
                    case '\\': sb.Append("\\\\"); break;
                    case '\"': sb.Append("\\\""); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    default:
                        if (c < 32)
                        {
                            sb.Append($"\\u{(int)c:X4}");
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }

        public static string SerializeObject(object obj, bool indented = true)
        {
            if (!indented)
                return obj.ToString();

            return PrettyPrint(obj.ToString());
        }

        private static string PrettyPrint(string json)
        {
            var sb = new StringBuilder();
            int indentLevel = 0;
            bool inQuotes = false;
            bool escapeNext = false;

            foreach (char c in json)
            {
                if (escapeNext)
                {
                    sb.Append(c);
                    escapeNext = false;
                    continue;
                }

                if (c == '\\')
                {
                    sb.Append(c);
                    escapeNext = true;
                    continue;
                }

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    sb.Append(c);
                    continue;
                }

                if (!inQuotes)
                {
                    if (c == '{' || c == '[')
                    {
                        sb.Append(c);
                        sb.Append(Environment.NewLine);
                        indentLevel++;
                        AddIndentation(sb, indentLevel);
                        continue;
                    }

                    if (c == '}' || c == ']')
                    {
                        sb.Append(Environment.NewLine);
                        indentLevel--;
                        AddIndentation(sb, indentLevel);
                        sb.Append(c);
                        continue;
                    }

                    if (c == ',')
                    {
                        sb.Append(c);
                        sb.Append(Environment.NewLine);
                        AddIndentation(sb, indentLevel);
                        continue;
                    }

                    if (c == ':')
                    {
                        sb.Append(c).Append(' ');
                        continue;
                    }

                    if (!char.IsWhiteSpace(c))
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private static void AddIndentation(StringBuilder sb, int indentLevel)
        {
            for (int i = 0; i < indentLevel; i++)
            {
                sb.Append("  ");
            }
        }

        public static JsonObject NewJsonObject() => new JsonObject();
    }
}
