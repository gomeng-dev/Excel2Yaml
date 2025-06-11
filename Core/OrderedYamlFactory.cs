using ExcelToYamlAddin.Config;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToYamlAddin.Core
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
            // 디버깅 로그 추가
            Debug.WriteLine($"[OrderedYamlFactory] SerializeToYaml 시작, includeEmptyFields: {includeEmptyFields}, 객체 타입: {(obj != null ? obj.GetType().Name : "null")}");
            Debug.WriteLine($"[OrderedYamlFactory] *** 중요 *** includeEmptyFields 값 확인: {includeEmptyFields}, 스택 트레이스: {Environment.StackTrace}");

            // includeEmptyFields가 true일 경우 빈 속성 유지, false일 경우 제거
            if (!includeEmptyFields)
            {
                Debug.WriteLine($"[OrderedYamlFactory] 빈 속성 제거 시작 (includeEmptyFields: {includeEmptyFields})");
                RemoveEmptyProperties(obj);
                Debug.WriteLine($"[OrderedYamlFactory] 빈 속성 제거 완료");
            }
            else
            {
                Debug.WriteLine($"[OrderedYamlFactory] 빈 속성 유지 모드 (includeEmptyFields: {includeEmptyFields})");
                // 빈 값을 명시적으로 표현하도록 처리 (null 값도 유지)
                PreserveEmptyValues(obj);
                
                // 추가: null 값이 있는 필드가 YAML에 포함되는지 디버깅 로그 출력
                if (obj is YamlObject yamlObj)
                {
                    Debug.WriteLine($"[OrderedYamlFactory] 객체 속성 수: {yamlObj.Properties.Count()}");
                    foreach (var prop in yamlObj.Properties)
                    {
                        Debug.WriteLine($"[OrderedYamlFactory] 속성: {prop.Key}, 값: {(prop.Value == null ? "null" : prop.Value.ToString())}");
                    }
                }
                
                Debug.WriteLine($"[OrderedYamlFactory] 빈 값 보존 처리 완료");
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
                // null 값은 빈 값으로 출력
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

            // 라인 엔딩 정규화 (\r\n 또는 \r -> \n)
            string normalizedValue = value.Replace("\r\n", "\n").Replace("\r", "\n");

            // 기존 needQuotes 결정 로직은 normalizedValue를 기준으로 수행되어야 일관성 있음
            bool isNumericString = !string.IsNullOrEmpty(normalizedValue) && normalizedValue.All(char.IsDigit);
            bool containsKorean = false;
            foreach (char c_iterator in normalizedValue) 
            {
                if (c_iterator >= '\uAC00' && c_iterator <= '\uD7A3')
                {
                    containsKorean = true;
                    break;
                }
            }

            bool containsNewline = normalizedValue.Contains("\n"); // 정규화된 값에서 개행 확인 (.NET Framework 4.7 이하 호환성)

            bool needQuotes = preserveQuotes ||
                              containsKorean ||
                              normalizedValue.Contains(":") ||
                              normalizedValue.Contains("%") ||
                              normalizedValue.Contains("#") ||
                              normalizedValue.Contains(",") ||
                              normalizedValue.StartsWith(" ") ||
                              normalizedValue.EndsWith(" ") ||
                              normalizedValue == "true" ||
                              normalizedValue == "false" ||
                              normalizedValue == "null" ||
                              (normalizedValue.Length > 0 && char.IsDigit(normalizedValue[0]) && !isNumericString);

            if (containsNewline) // 정규화된 값에 개행이 있으면 무조건 따옴표 필요
            {
                needQuotes = true;
            }

            if (needQuotes)
            {
                sb.Append('"');
                // 이스케이프 및 문자열 구성은 normalizedValue 사용
                foreach (char c in normalizedValue)
                {
                    switch (c)
                    {
                        case '"':  sb.Append("\\\""); break; // 큰따옴표 이스케이프
                        case '\\': sb.Append("\\\\"); break; // 백슬래시 이스케이프
                        case '\n': sb.Append("\\n");  break; // 개행은 \\n으로 이스케이프
                        // \r은 normalizedValue에 없을 것이므로 case '\r' 불필요
                        case '\t': sb.Append("\\t");  break; // 탭도 이스케이프 (표준 YAML/JSON 문자열 규칙)
                        default:
                            sb.Append(c); // 현재는 다른 문자는 그대로 추가
                            break;
                    }
                }
                sb.Append('"');
            }
            else
            {
                // 따옴표가 필요 없는 경우. 이 경우 normalizedValue는 개행이나 YAML 특수문자를 포함하지 않아야 함.
                sb.Append(normalizedValue);
            }
        }

        private static bool IsEffectivelyEmpty(object value)
        {
            if (value == null)
                return true;

            if (value is string str && string.IsNullOrEmpty(str))
                return true;

            if (value is YamlObject obj)
            {
                if (!obj.HasValues)
                    return true;

                // 모든 속성이 효과적으로 비어있는지 확인
                foreach (var prop in obj.Properties)
                {
                    if (!IsEffectivelyEmpty(prop.Value))
                        return false;
                }
                return true;
            }

            if (value is YamlArray arr)
            {
                if (!arr.HasValues)
                    return true;

                // 모든 배열 항목이 효과적으로 비어있는지 확인
                foreach (var item in arr.Items)
                {
                    if (!IsEffectivelyEmpty(item))
                        return false;
                }
                return true;
            }

            return false;
        }

        private static void SerializeYamlObject(YamlObject obj, StringBuilder sb, int level, int indentSize, YamlStyle style, bool preserveQuotes)
        {
            // 빈 객체 처리
            if (!obj.HasValues)
            {
                if (style == YamlStyle.Flow)
                {
                    sb.Append("{}");
                }
                if (style == YamlStyle.Block)
                {
                    // 빈 객체를 Block 스타일에서 "빈 객체 표시"로 수정
                    sb.Append("{}");
                    sb.AppendLine();
                }
                return;
            }

            // 모든 속성이 효과적으로 비어있는지 확인
            bool allPropertiesEffectivelyEmpty = true;
            foreach (var prop in obj.Properties)
            {
                if (!IsEffectivelyEmpty(prop.Value))
                {
                    allPropertiesEffectivelyEmpty = false;
                    break;
                }
            }

            // 모든 속성이 효과적으로 비어있으면 빈 객체로 출력
            if (allPropertiesEffectivelyEmpty)
            {
                if (style == YamlStyle.Flow)
                {
                    sb.Append("{}");
                }
                else
                {
                    sb.Append("{}");
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
                    SerializeObject(kvp.Value, sb, level + 2, indentSize, style, preserveQuotes);
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

                    // null 값 처리
                    if (kvp.Value == null)
                    {
                        // null 값은 빈 값으로 출력
                        sb.AppendLine();
                    }
                    // 빈 배열인 경우 바로 개행 처리
                    else if (kvp.Value is YamlArray yamlArray && !yamlArray.HasValues)
                    {
                        // 빈 배열은 [] 표기로 출력
                        sb.Append("[]");
                        sb.AppendLine();
                    }
                    // 빈 객체인 경우 바로 개행 처리
                    else if (kvp.Value is YamlObject yamlObj && !yamlObj.HasValues)
                    {
                        // 빈 객체는 {} 표기로 출력
                        sb.Append("{}");
                        sb.AppendLine();
                    }
                    // 빈 문자열 처리
                    else if (kvp.Value is string str && string.IsNullOrEmpty(str))
                    {
                        sb.Append(preserveQuotes ? "\"\"" : "");
                        sb.AppendLine();
                    }
                    else if (kvp.Value is YamlObject || kvp.Value is YamlArray)
                    {
                        // MAP 노드의 자식들은 2 레벨 더 들여쓰기 (일관성 유지)
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
            // 빈 배열 처리
            if (!array.HasValues)
            {
                // 빈 배열은 [] 표기로 출력
                sb.Append("[]");
                sb.AppendLine();
                return;
            }

            // 모든 항목이 효과적으로 비어있는지 확인
            bool allItemsEffectivelyEmpty = true;
            foreach (var item in array.Items)
            {
                if (!IsEffectivelyEmpty(item))
                {
                    allItemsEffectivelyEmpty = false;
                    break;
                }
            }

            // 모든 항목이 효과적으로 비어있으면 빈 배열로 출력
            if (allItemsEffectivelyEmpty)
            {
                sb.Append("[]");
                sb.AppendLine();
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

                    if (item == null)
                    {
                        // null 값은 빈 값으로 출력
                        sb.AppendLine();
                    }
                    else if (item is string str && string.IsNullOrEmpty(str))
                    {
                        sb.Append(preserveQuotes ? "\"\"" : "");
                        sb.AppendLine();
                    }
                    else if (item is YamlObject yamlObj)
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

                                    if (prop.Value == null)
                                    {
                                        // null 값은 빈 값으로 출력
                                        sb.AppendLine();
                                    }
                                    else if (prop.Value is string propStr && string.IsNullOrEmpty(propStr))
                                    {
                                        sb.Append(preserveQuotes ? "\"\"" : "");
                                        sb.AppendLine();
                                    }
                                    else if (prop.Value is YamlObject || prop.Value is YamlArray)
                                    {
                                        // MAP 노드의 자식들은 2 레벨 더 들여쓰기 (일관성 유지)
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

                                    if (prop.Value == null)
                                    {
                                        // null 값은 빈 값으로 출력
                                        sb.AppendLine();
                                    }
                                    else if (prop.Value is string propStr && string.IsNullOrEmpty(propStr))
                                    {
                                        sb.Append(preserveQuotes ? "\"\"" : "");
                                        sb.AppendLine();
                                    }
                                    else if (prop.Value is YamlObject || prop.Value is YamlArray)
                                    {
                                        // MAP 노드의 자식들은 2 레벨 더 들여쓰기 (일관성 유지)
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

        // 빈 값(스칼라, 배열, 객체 등)을 모두 보존하도록 처리하는 메서드
        public static void PreserveEmptyValues(object token, bool isRoot = true)
        {
            if (token is YamlArray array)
            {
                // 루트 배열인 경우, 빈 항목 제거 처리 (체크박스 상태와 무관하게)
                if (isRoot)
                {
                    Debug.WriteLine("[OrderedYamlFactory] 루트 배열 처리 중 (PreserveEmptyValues)");

                    // 빈 항목 확인 및 제거
                    bool hasEmptyItems = false;
                    foreach (var item in array.Items)
                    {
                        if (IsEmptyItem(item))
                        {
                            hasEmptyItems = true;
                            break;
                        }
                    }

                    // 빈 항목이 있으면 제거
                    if (hasEmptyItems)
                    {
                        Debug.WriteLine("[OrderedYamlFactory] 루트 배열에서 빈 항목 제거 (PreserveEmptyValues)");

                        // 뒤에서부터 제거 (인덱스 변경 방지)
                        for (int i = array.Count - 1; i >= 0; i--)
                        {
                            if (IsEmptyItem(array[i]))
                            {
                                array.RemoveAt(i);
                            }
                        }
                    }
                }

                // 배열 내의 항목들을 재귀적으로 처리 (루트가 아닌 것으로 처리)
                foreach (var item in array.Items)
                {
                    if (item != null)  // null이 아닌 항목만 처리
                    {
                        PreserveEmptyValues(item, false);
                    }
                }
            }
            else if (token is YamlObject obj)
            {
                // 객체 내의 모든 속성을 재귀적으로 처리
                foreach (var prop in obj.Properties)
                {
                    if (prop.Value != null)  // null이 아닌 값만 처리
                    {
                        PreserveEmptyValues(prop.Value, false);
                    }
                }
            }
            // 스칼라 값은 그대로 유지 (처리할 필요 없음)
        }

        // 빈 객체만 포함하는 배열을 빈 배열로 변환 및 루트 배열의 빈 항목 제거
        public static void CleanEmptyArrays(object token, bool isRoot = true)
        {
            if (token is YamlArray array)
            {
                // 루트 배열인 경우, 빈 항목 제거 처리
                if (isRoot)
                {
                    Debug.WriteLine("[OrderedYamlFactory] 루트 배열 처리 중");

                    // 먼저 빈 항목이 있는지 확인
                    bool hasEmptyItems = false;
                    foreach (var item in array.Items)
                    {
                        if (IsEmptyItem(item))
                        {
                            hasEmptyItems = true;
                            break;
                        }
                    }

                    // 빈 항목이 있으면 제거
                    if (hasEmptyItems)
                    {
                        Debug.WriteLine("[OrderedYamlFactory] 루트 배열에서 빈 항목 제거");

                        // 뒤에서부터 제거 (인덱스 변경 방지)
                        for (int i = array.Count - 1; i >= 0; i--)
                        {
                            if (IsEmptyItem(array[i]))
                            {
                                array.RemoveAt(i);
                            }
                        }
                    }
                }

                // 배열 내의 항목을 재귀적으로 처리 (루트가 아닌 것으로 처리)
                foreach (var item in array.Items)
                {
                    CleanEmptyArrays(item, false);
                }

                // 모든 항목이 빈 객체인지 확인 (빈 배열로 만들지 판단)
                bool allItemsAreEmptyObjects = true;
                foreach (var item in array.Items)
                {
                    if (!(item is YamlObject obj) || obj.HasValues)
                    {
                        allItemsAreEmptyObjects = false;
                        break;
                    }
                }

                // 모든 항목이 빈 객체라면 빈 배열로 변환 (루트가 아닌 경우에만)
                if (allItemsAreEmptyObjects && array.Count > 0 && !isRoot)
                {
                    Debug.WriteLine("[OrderedYamlFactory] 빈 객체들만 포함한 배열을 빈 배열로 변환");

                    // 배열 내 모든 항목 제거
                    for (int i = array.Count - 1; i >= 0; i--)
                    {
                        array.RemoveAt(i);
                    }
                }
            }
            else if (token is YamlObject obj)
            {
                // 모든 속성에 대해 재귀적으로 처리 (루트가 아닌 것으로 처리)
                foreach (var prop in obj.Properties)
                {
                    CleanEmptyArrays(prop.Value, false);
                }
            }
        }

        // 항목이 빈 항목인지 확인 (null, 빈 문자열, 빈 객체, 빈 배열)
        private static bool IsEmptyItem(object item)
        {
            if (item == null)
                return true;

            if (item is string str && string.IsNullOrEmpty(str))
                return true;

            if (item is YamlObject obj && !obj.HasValues)
                return true;

            // 여기서는 빈 배열도 빈 항목으로 간주 (루트 배열에서 제거)
            if (item is YamlArray arr && !arr.HasValues)
                return true;

            return false;
        }
    }
}