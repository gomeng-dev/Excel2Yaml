using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics; // Debug.WriteLine을 위해 추가
using System.IO;
using System.Xml.Linq;

namespace ExcelToYamlAddin.Core.YamlPostProcessors // 네임스페이스를 Core로 변경 (Ribbon.cs의 호출 경로와 일치시키기 위해)
{
    public static class YamlToXmlConverter
    {
        /// <summary>
        /// YAML에서 파싱된 IDictionary<string, object> 데이터를 XML 문자열로 변환합니다.
        /// ExcelToXmlConversionRules.md의 규칙을 따릅니다 (키가 '_'로 시작하면 속성으로 처리).
        /// </summary>
        /// <param name="yamlData">YAML에서 파싱된 데이터입니다.</param>
        /// <param name="rootElementName">XML 문서의 루트 요소 이름입니다.</param>
        /// <returns>생성된 XML 문자열입니다.</returns>
        public static string ConvertToXmlString(IDictionary<string, object> yamlData, string rootElementName)
        {
            if (string.IsNullOrEmpty(rootElementName))
            {
                // 루트 요소 이름이 없으면 기본값 또는 예외 처리
                Debug.WriteLine("[YamlToXmlConverter.ConvertToXmlString] Error: rootElementName cannot be null or empty.");
                throw new ArgumentException("Root element name cannot be null or empty.", nameof(rootElementName));
            }

            if (yamlData == null)
            {
                Debug.WriteLine($"[YamlToXmlConverter.ConvertToXmlString] Input yamlData is null for root '{rootElementName}'. Returning empty root.");
                // 빈 루트 요소만 반환
                return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), new XElement(rootElementName)).ToString();
            }

            XElement rootElement = new XElement(rootElementName);
            AddNodes(rootElement, yamlData);

            XDocument xmlDoc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), rootElement);
            return xmlDoc.ToString();
        }

        private static void AddNodes(XElement parentElement, IDictionary<string, object> data)
        {
            if (data == null) return;

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

                if (key.StartsWith("_")) // 규칙: '_'로 시작하는 키는 XML 속성으로 처리
                {
                    string attributeName = key.Substring(1);
                    if (string.IsNullOrEmpty(attributeName))
                    {
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Warning: Attribute name derived from key '{key}' is empty. Skipping attribute.");
                        continue;
                    }
                    if (value != null)
                    {
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Adding attribute '{attributeName}' with value '{value}' to element '{parentElement.Name}'.");
                        parentElement.SetAttributeValue(attributeName, value.ToString());
                    }
                    else
                    {
                        // null 값 속성은 추가하지 않거나, 빈 문자열로 추가할 수 있습니다. 여기서는 추가하지 않음.
                        Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Attribute '{attributeName}' has null value. Skipping attribute for element '{parentElement.Name}'.");
                    }
                }
                else if (value is IList listValue)
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' is IList with {listValue.Count} items.");
                    foreach (var item in listValue)
                    {
                        XElement listElement = new XElement(key);
                        if (item is IDictionary<string, object> itemDictStringKey)
                        {
                            Debug.WriteLine($"[YamlToXmlConverter.AddNodes] List item for key '{key}' is IDictionary<string, object>. Recursively calling AddNodes.");
                            AddNodes(listElement, itemDictStringKey);
                        }
                        else if (item is IDictionary itemDictObjectKey) // YamlDotNet이 Dictionary<object, object> 반환 시
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
                            // listElement.Value는 기본적으로 비어있음
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
                else if (value is IDictionary dictValueObjectKey) // YamlDotNet이 Dictionary<object, object> 반환 시
                {
                    XElement childElement = new XElement(key);
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' is IDictionary (object key, type: {dictValueObjectKey.GetType()}). Attempting conversion.");
                    ConvertAndAddNonGenericDictionary(childElement, dictValueObjectKey);
                    parentElement.Add(childElement);
                }
                else if (value != null) // 단순 값 (스칼라)
                {
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' has a non-null scalar value (type: {value.GetType()}). Adding as element value: '{value}'.");
                    parentElement.Add(new XElement(key, value.ToString()));
                }
                else
                {
                    // value가 null이고, 속성도 아니고, 리스트나 딕셔너리도 아닌 경우 (단순 null 값)
                    // XML에서는 null 값을 표현하는 표준 방식이 없으므로, 빈 요소를 추가하거나 아무것도 하지 않을 수 있습니다.
                    // 여기서는 빈 요소를 추가합니다.
                    Debug.WriteLine($"[YamlToXmlConverter.AddNodes] Key '{key}' has null value. Adding empty element.");
                    parentElement.Add(new XElement(key));
                }
            }
        }

        // Helper method to convert and add nodes from a non-specifically-typed IDictionary
        private static void ConvertAndAddNonGenericDictionary(XElement targetElement, IDictionary dictionary)
        {
            var convertedDict = new Dictionary<string, object>();
            if (dictionary == null)
            {
                Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Input dictionary is null for target element '{targetElement.Name}'.");
                // 빈 딕셔너리에 대해 AddNodes를 호출하면 아무것도 추가되지 않음
                AddNodes(targetElement, convertedDict);
                return;
            }

            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key == null)
                {
                    Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Null key found in dictionary for element '{targetElement.Name}'. Skipping entry.");
                    continue;
                }

                string stringKey = entry.Key.ToString(); // 키를 문자열로 변환
                if (!string.IsNullOrEmpty(stringKey))
                {
                    Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Converting key '{entry.Key}' (type: {entry.Key.GetType()}) to string key '{stringKey}'. Value type: {entry.Value?.GetType().ToString() ?? "null"}");
                    convertedDict[stringKey] = entry.Value;
                }
                else
                {
                    Debug.WriteLine($"[YamlToXmlConverter.ConvertAndAddNonGenericDictionary] Warning: Dictionary Key for element '{targetElement.Name}' could not be converted to a non-empty string. Original Key: '{entry.Key}', Type: {entry.Key.GetType()}");
                }
            }
            AddNodes(targetElement, convertedDict); // Recurse with the strongly-typed dictionary
        }
    }
}
