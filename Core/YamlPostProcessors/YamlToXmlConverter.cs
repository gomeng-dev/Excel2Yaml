using System;
using System.Collections; // Required for non-generic IDictionary and DictionaryEntry
using System.Collections.Generic;
using System.Diagnostics;
using System.IO; // Retained as it was present, though not strictly used by this logic
using System.Linq; // Added for the new line
using System.Xml.Linq;

namespace ExcelToYamlAddin.Core.YamlPostProcessors
{
    public static class YamlToXmlConverter
    {
        /// <summary>
        /// Converts YAML data (parsed as IDictionary<string, object>) into an XML string.
        /// Keys starting with '_' are treated as attributes.
        /// </summary>
        /// <param name="yamlData">The YAML data, typically the content for the root element.</param>
        /// <param name="rootElementName">The name for the root XML element.</param>
        /// <returns>An XML string representation of the YAML data.</returns>
        public static string ConvertToXmlString(IDictionary<string, object> yamlData, string rootElementName)
        {
            XElement actualRootElement;
            string chosenRootNameForDebug; // 디버깅 메시지에 사용할 최종 루트 이름

            // 시나리오 1: YAML 데이터가 비어있거나 null인 경우
            if (yamlData == null || !yamlData.Any()) // System.Linq.Any()
            {
                Debug.WriteLine("[YamlToXmlConverter.ConvertToXmlString] YAML data is null or empty.");
                if (string.IsNullOrEmpty(rootElementName))
                {
                    // 데이터도 없고, 명시적 루트 이름도 없으면 XML 생성 불가
                    Debug.WriteLine("[YamlToXmlConverter.ConvertToXmlString] Error: Both YAML data and rootElementName are null/empty. Cannot create XML.");
                    throw new ArgumentException("Cannot create XML: YAML data is null/empty and no rootElementName is provided.", nameof(rootElementName));
                }
                // 데이터는 없지만 명시적 루트 이름이 있으면, 해당 이름으로 빈 루트 요소 생성
                chosenRootNameForDebug = rootElementName;
                Debug.WriteLine($"[YamlToXmlConverter.ConvertToXmlString] Creating empty XML with root '{chosenRootNameForDebug}'.");
                actualRootElement = new XElement(chosenRootNameForDebug);
            }
            // 시나리오 2: YAML 데이터에 키가 하나만 있고, 그 값의 타입이 IDictionary (string 키 또는 object 키)인 경우
            // 이 경우 YAML의 첫 번째 키를 실제 루트로 사용하고, rootElementName은 무시.
            else if (yamlData.Count == 1 &&
                     (yamlData.First().Value is IDictionary<string, object> || yamlData.First().Value is IDictionary))
            {
                var firstEntry = yamlData.First();
                chosenRootNameForDebug = firstEntry.Key;
                Debug.WriteLine($"[YamlToXmlConverter.ConvertToXmlString] YAML data has a single dictionary entry. Using its key '{chosenRootNameForDebug}' as XML root. Provided rootElementName '{rootElementName}' is ignored.");
                actualRootElement = new XElement(chosenRootNameForDebug);

                if (firstEntry.Value is IDictionary<string, object> dictStringKey)
                {
                    AddNodes(actualRootElement, dictStringKey);
                }
                else if (firstEntry.Value is IDictionary dictObjectKey) // 이것은 IDictionary<string, object>가 아닌 IDictionary를 처리
                {
                    ConvertAndAddNonGenericDictionary(actualRootElement, dictObjectKey);
                }
            }
            // 시나리오 3: YAML 데이터가 다른 구조를 가지는 경우 (여러 키, 또는 단일 키지만 값이 딕셔너리가 아님)
            // 이 경우 rootElementName을 사용해야 함. 만약 rootElementName이 없다면 에러.
            else
            {
                if (string.IsNullOrEmpty(rootElementName))
                {
                    Debug.WriteLine("[YamlToXmlConverter.ConvertToXmlString] Error: YAML data structure requires a rootElementName, but it's null or empty.");
                    throw new ArgumentException("rootElementName must be provided when YAML data is multi-keyed or has a single non-dictionary entry.", nameof(rootElementName));
                }
                chosenRootNameForDebug = rootElementName;
                Debug.WriteLine($"[YamlToXmlConverter.ConvertToXmlString] YAML data is multi-keyed or single non-dictionary. Using provided rootElementName '{chosenRootNameForDebug}' as XML root.");
                actualRootElement = new XElement(chosenRootNameForDebug);
                AddNodes(actualRootElement, yamlData); // 전체 yamlData를 이 루트의 자식으로 추가
            }

            Debug.WriteLine($"[YamlToXmlConverter.ConvertToXmlString] Final XML document will be based on root '{chosenRootNameForDebug}'.");
            XDocument xmlDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), actualRootElement);
            return xmlDoc.ToString();
        }

        /// <summary>
        /// Recursively populates an XElement based on the provided dictionary data.
        /// </summary>
        private static void AddNodes(XElement parentElement, IDictionary<string, object> data)
        {
            if (data == null) return;

            // 1. __text 키가 현재 data 딕셔너리에 있으면 parentElement의 값으로 설정
            if (data.TryGetValue("__text", out object textContentValue))
            {
                if (textContentValue != null)
                {
                    parentElement.Value = textContentValue.ToString();
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Element '{parentElement.Name}' set text from '__text' key: '{textContentValue}'.");
                }
                else
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Element '{parentElement.Name}' found '__text' key, but its value is null. Text not set.");
                }
            }

            foreach (var kvp in data)
            {
                string key = kvp.Key;
                object value = kvp.Value;
                string valueTypeString = value?.GetType().ToString() ?? "null";

                Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Processing Key: '{key}', Value Type: '{valueTypeString}' for Parent: '{parentElement.Name}'");

                if (string.IsNullOrEmpty(key))
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Warning: Key is null or empty for value type '{valueTypeString}'. Skipping this entry.");
                    continue;
                }

                // 2. __text 키는 이미 위에서 parentElement의 값으로 처리했으므로 건너뜀
                if (key == "__text")
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '__text' for parent '{parentElement.Name}' skipped as it has been handled for the element's direct value.");
                    continue;
                }

                if (key.StartsWith("_")) // 규칙: '_'로 시작하는 키는 XML 속성으로 처리 (parentElement에 대한 속성)
                {
                    string attributeName = key.Substring(1);
                    if (string.IsNullOrEmpty(attributeName))
                    {
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Warning: Attribute name derived from key '{key}' for parent '{parentElement.Name}' is empty. Skipping attribute.");
                        continue;
                    }
                    if (value != null)
                    {
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Adding attribute '{attributeName}' with value '{value}' to element '{parentElement.Name}'.");
                        parentElement.SetAttributeValue(attributeName, value.ToString());
                    }
                    else
                    {
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Attribute '{attributeName}' for '{parentElement.Name}' has null value. Skipping attribute.");
                    }
                }
                else if (value is IList listValue)
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' is IList with {listValue.Count} items.");
                    foreach (var item in listValue)
                    {
                        // For list items, the element name is the key of the list itself (e.g., <WinReward> for items in WinReward list)
                        XElement listElement = new XElement(key);
                        if (item is IDictionary<string, object> itemDictStringKey)
                        {
                            Debug.WriteLine($"[YamlToXmlConverter.AddNodes] List item for key '{key}' is IDictionary<string, object>. Recursively calling AddNodes.");
                            AddNodes(listElement, itemDictStringKey);
                        }
                        else if (item is IDictionary itemDictObjectKey) // Handles cases where YamlDotNet returns Dictionary<object, object>
                        {
                            Debug.WriteLine($"[YamlToXmlConverter.AddNodes] List item for key '{key}' is IDictionary (object key, type: {item.GetType()}). Attempting conversion for list item.");
                            ConvertAndAddNonGenericDictionary(listElement, itemDictObjectKey);
                        }
                        else if (item != null)
                        {
                            Debug.WriteLine($"[YamlToXmlConverter.AddNodes] List item for key '{key}' is a scalar value: '{item}'.");
                            listElement.Value = item.ToString();
                        }
                        else
                        {
                            Debug.WriteLine($"[YamlToXmlConverter.AddNodes] List item for key '{key}' is null. Adding empty element.");
                            // listElement.Value is already empty by default
                        }
                        parentElement.Add(listElement);
                    }
                }
                else if (value is IDictionary<string, object> dictValueStringKey)
                {
                    XElement childElement = new XElement(key);
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' is IDictionary<string, object> (type: {dictValueStringKey.GetType()}). Recursively calling AddNodes.");
                    AddNodes(childElement, dictValueStringKey);
                    parentElement.Add(childElement);
                }
                else if (value is IDictionary dictValueObjectKey) // Handles cases where YamlDotNet returns Dictionary<object, object>
                {
                    XElement childElement = new XElement(key);
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' is IDictionary (object key, type: {dictValueObjectKey.GetType()}). Attempting conversion.");
                    ConvertAndAddNonGenericDictionary(childElement, dictValueObjectKey);
                    parentElement.Add(childElement);
                }
                else if (value != null) // Scalar value
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' has a non-null scalar value (type: {value.GetType()}). Adding as element value: '{value}'.");
                    parentElement.Add(new XElement(key, value.ToString()));
                }
                else
                {
                    // Value is null, and it's not an attribute, list, or dictionary.
                    // Depending on desired XML for nulls, either add an empty element or skip.
                    // Current behavior: Add empty element if key is present but value is null.
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' has a null scalar value. Adding empty element <{key}/>.");
                    parentElement.Add(new XElement(key));
                }
            }
        }

        /// <summary>
        /// Helper method to convert a non-generic IDictionary (potentially with object keys)
        /// to an IDictionary<string, object> and then process it using AddNodes.
        /// </summary>
        private static void ConvertAndAddNonGenericDictionary(XElement parentElement, IDictionary data)
        {
            if (data == null) return;

            var stringKeyData = new Dictionary<string, object>();
            foreach (DictionaryEntry entry in data)
            {
                if (entry.Key == null)
                {
                    Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Warning: Null key found in non-generic IDictionary for parent '{parentElement.Name}'. Skipping entry.");
                    continue;
                }
                string keyString = entry.Key.ToString();
                if (string.IsNullOrEmpty(keyString))
                {
                    Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Warning: Key '{entry.Key}' converted to empty string. Skipping entry for parent '{parentElement.Name}'.");
                    continue;
                }
                stringKeyData[keyString] = entry.Value;
            }
            AddNodes(parentElement, stringKeyData); // Call the main AddNodes with the converted dictionary
        }
    }
}
