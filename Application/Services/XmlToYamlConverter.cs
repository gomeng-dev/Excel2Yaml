using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ExcelToYamlAddin.Infrastructure.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core
{
    /// <summary>
    /// XML을 YAML로 변환하는 클래스
    /// YamlToXmlConverter의 역변환 기능을 제공합니다.
    /// </summary>
    public class XmlToYamlConverter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<XmlToYamlConverter>();
        
        // XML 속성을 나타내는 접두사
        private const string AttributePrefix = "_";
        
        // 텍스트 내용을 나타내는 특수 키
        private const string TextContentKey = "__text";
        
        // 배열로 처리된 속성들을 추적 (한 번 배열이면 계속 배열 유지)
        private HashSet<string> arrayProperties = new HashSet<string>();
        
        // 객체로 처리된 속성들을 추적 (속성이나 자식 요소가 있는 경우)
        private HashSet<string> objectProperties = new HashSet<string>();

        /// <summary>
        /// XML 문자열을 YAML 문자열로 변환합니다.
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <returns>YAML 문자열</returns>
        public string ConvertToYaml(string xmlContent)
        {
            try
            {
                Logger.Information("XML to YAML 변환 시작");
                
                // 새로운 방식 사용: ConvertToDictionary를 통해 변환
                var dictionary = ConvertToDictionary(xmlContent);
                
                // Dictionary를 YAML로 시리얼라이즈
                var serializer = new YamlDotNet.Serialization.SerializerBuilder()
                    .Build();
                
                var yamlContent = serializer.Serialize(dictionary);
                
                Logger.Information("XML to YAML 변환 완료");
                Logger.Debug($"변환된 YAML:\n{yamlContent}");
                
                return yamlContent;
            }
            catch (Exception ex)
            {
                Logger.Error($"XML to YAML 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// XML을 IDictionary로 변환합니다. (YamlToXmlConverter와 호환되는 형식)
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <returns>IDictionary 형식의 데이터</returns>
        public IDictionary<string, object> ConvertToDictionary(string xmlContent)
        {
            try
            {
                Logger.Information("XML to Dictionary 변환 시작");
                
                // 속성 추적 초기화
                arrayProperties.Clear();
                objectProperties.Clear();
                
                // XML 파싱
                var doc = XDocument.Parse(xmlContent);
                var root = doc.Root;
                
                // 1단계: 전체 문서를 스캔해서 배열 속성들을 미리 식별
                ScanForArrayProperties(root, root.Name.LocalName);
                
                if (root == null)
                {
                    throw new InvalidOperationException("XML 루트 요소를 찾을 수 없습니다.");
                }
                
                // 새로운 형태 처리: 루트가 시트명이고 자식들이 배열 요소인 경우
                if (HasDirectChildrenAsArrayElements(root))
                {
                    return ConvertDirectChildrenStructure(root);
                }
                
                // 기존 방식으로 처리
                var result = new Dictionary<string, object>();
                result[root.Name.LocalName] = ConvertElementToDictionary(root, root.Name.LocalName);
                
                Logger.Information("XML to Dictionary 변환 완료");
                return result;
            }
            catch (Exception ex)
            {
                Logger.Error($"XML to Dictionary 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 루트 요소가 직접 자식들을 배열 요소로 가지는 구조인지 확인합니다.
        /// 예: <Missions3><Mission>...</Mission><Mission>...</Mission></Missions3>
        /// 또는: <Sheet1><Test>...</Test><Test>...</Test></Sheet1>
        /// </summary>
        private bool HasDirectChildrenAsArrayElements(XElement root)
        {
            if (!root.HasElements)
            {
                Logger.Information($"[XmlToYamlConverter] 루트 '{root.Name.LocalName}'에 자식 요소가 없음");
                return false;
            }

            // 모든 직접 자식들의 이름을 그룹화
            var childGroups = root.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToList();

            Logger.Information($"[XmlToYamlConverter] 루트 '{root.Name.LocalName}' 자식 그룹 수: {childGroups.Count}");

            // 하나의 그룹에 여러 요소가 있는 경우 (배열 구조)
            foreach (var group in childGroups)
            {
                Logger.Information($"[XmlToYamlConverter] 그룹 '{group.Key}': {group.Count()}개 요소");
                if (group.Count() > 1)
                {
                    Logger.Information($"[XmlToYamlConverter] ✓ 배열 구조 감지: {group.Key} 요소가 {group.Count()}개 발견");
                    return true;
                }
            }

            Logger.Information($"[XmlToYamlConverter] ✗ 배열 구조 감지되지 않음: 모든 그룹이 단일 요소");
            return false;
        }

        /// <summary>
        /// 직접 자식 구조를 배열 형태로 변환합니다.
        /// 루트 시트명 하위에 배열을 생성하고, 각 자식을 객체로 변환합니다.
        /// 각 자식 객체의 이름을 키로 사용하여 YAML 구조를 만듭니다.
        /// </summary>
        private IDictionary<string, object> ConvertDirectChildrenStructure(XElement root)
        {
            Logger.Information($"[XmlToYamlConverter] 직접 자식 구조 변환 시작: 루트={root.Name.LocalName}");
            
            var result = new Dictionary<string, object>();
            var arrayList = new List<object>();

            // 자식 요소들을 순서대로 처리 (그룹화하지 않음)
            foreach (var child in root.Elements())
            {
                // 각 자식을 객체로 변환하되, 자식의 이름을 키로 사용
                var childObject = new Dictionary<string, object>();
                var childContent = ConvertElementToDictionary(child, $"{root.Name.LocalName}.{child.Name.LocalName}");
                
                // 자식 요소의 이름을 키로 하고, 내용을 값으로 하는 객체 생성
                // 예: <Test>...</Test> -> { Test: {...} }
                childObject[child.Name.LocalName] = childContent;
                arrayList.Add(childObject);
                
                Logger.Information($"[XmlToYamlConverter] 자식 요소 변환: {child.Name.LocalName} -> 객체에 키로 추가됨");
            }

            // 첫 번째 자식 요소의 이름을 배열의 키로 사용
            // XML: <Sheet1><Test>...</Test><Test>...</Test></Sheet1>
            // YAML: Test: [{ Test: {...} }, { Test: {...} }]
            string firstChildName = root.Elements().First().Name.LocalName;
            string arrayKey = GetArrayKeyFromChildName(firstChildName);
            result[arrayKey] = arrayList;

            Logger.Information($"[XmlToYamlConverter] 직접 자식 구조 변환 완료: 배열 키={arrayKey}, 항목 수={arrayList.Count}");
            Logger.Information($"[XmlToYamlConverter] 각 배열 항목은 '{firstChildName}: {{...}}' 형태로 구성됨");
            return result;
        }

        /// <summary>
        /// 자식 요소 이름으로부터 배열의 키 이름을 결정합니다.
        /// 원본 데이터를 보존하기 위해 단순한 규칙을 적용합니다.
        /// </summary>
        private string GetArrayKeyFromChildName(string childName)
        {
            if (string.IsNullOrEmpty(childName))
                return "Items";

            // 자식 이름을 그대로 사용 (데이터 오염 방지)
            // 예: Mission -> Mission (복수형 변환하지 않음)
            return childName;
        }

        /// <summary>
        /// XML 요소를 YAML 노드로 변환합니다.
        /// </summary>
        private YamlNode ConvertElementToYaml(XElement element)
        {
            // 동일한 이름의 자식 요소들을 그룹화
            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToDictionary(g => g.Key, g => g.ToList());
            
            // 루트가 배열인지 확인 (모든 자식이 동일한 이름)
            if (childGroups.Count == 1 && childGroups.First().Value.Count > 1)
            {
                // 배열로 처리
                var sequence = new YamlSequenceNode();
                foreach (var child in childGroups.First().Value)
                {
                    sequence.Add(ConvertElementToYaml(child));
                }
                
                // 루트 요소를 포함하는 매핑 생성
                var rootMapping = new YamlMappingNode();
                rootMapping.Add(element.Name.LocalName, sequence);
                return rootMapping;
            }
            
            // 객체로 처리
            var mapping = new YamlMappingNode();
            
            // 속성 처리 (xmlns 같은 네임스페이스 속성 제외)
            foreach (var attr in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
            {
                var attributeName = AttributePrefix + attr.Name.LocalName;
                mapping.Add(attributeName, new YamlScalarNode(attr.Value));
                
                // DescFormat 요소의 Arg 속성들 특별 디버깅
                if (element.Name.LocalName == "DescFormat" && attr.Name.LocalName.StartsWith("Arg"))
                {
                    Logger.Information($"🔍 DescFormat XML 속성 변환: {attr.Name.LocalName} -> {attributeName} = '{attr.Value}'");
                }
            }
            
            // 텍스트 내용이 있고 자식 요소가 없는 경우
            if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
            {
                mapping.Add(TextContentKey, new YamlScalarNode(element.Value));
            }
            
            // 자식 요소 처리
            foreach (var group in childGroups)
            {
                if (group.Value.Count > 1)
                {
                    // 배열인 경우
                    var sequence = new YamlSequenceNode();
                    foreach (var child in group.Value)
                    {
                        var childNode = ConvertChildElementToYaml(child);
                        sequence.Add(childNode);
                    }
                    mapping.Add(group.Key, sequence);
                }
                else
                {
                    // 단일 요소
                    var child = group.Value.First();
                    var childNode = ConvertChildElementToYaml(child);
                    mapping.Add(group.Key, childNode);
                }
            }
            
            // 최상위 레벨에서만 요소 이름을 키로 사용
            if (element.Parent == null)
            {
                var rootMapping = new YamlMappingNode();
                rootMapping.Add(element.Name.LocalName, mapping);
                return rootMapping;
            }
            
            return mapping;
        }

        /// <summary>
        /// 자식 XML 요소를 YAML 노드로 변환합니다.
        /// </summary>
        private YamlNode ConvertChildElementToYaml(XElement element)
        {
            // 속성이나 자식 요소가 있는 경우 객체로 처리
            if (element.HasAttributes || element.HasElements)
            {
                var mapping = new YamlMappingNode();
                
                // 속성 처리
                foreach (var attr in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
                {
                    mapping.Add(AttributePrefix + attr.Name.LocalName, new YamlScalarNode(attr.Value));
                }
                
                // 텍스트 내용이 있는 경우
                if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
                {
                    mapping.Add(TextContentKey, new YamlScalarNode(element.Value));
                }
                
                // 자식 요소 처리
                var childGroups = element.Elements()
                    .GroupBy(e => e.Name.LocalName)
                    .ToDictionary(g => g.Key, g => g.ToList());
                
                foreach (var group in childGroups)
                {
                    if (group.Value.Count > 1)
                    {
                        // 배열인 경우
                        var sequence = new YamlSequenceNode();
                        foreach (var child in group.Value)
                        {
                            sequence.Add(ConvertChildElementToYaml(child));
                        }
                        mapping.Add(group.Key, sequence);
                    }
                    else
                    {
                        // 단일 요소
                        var child = group.Value.First();
                        mapping.Add(group.Key, ConvertChildElementToYaml(child));
                    }
                }
                
                return mapping;
            }
            
            // 단순 텍스트 요소
            return new YamlScalarNode(element.Value);
        }

        /// <summary>
        /// 전체 XML 문서를 스캔해서 배열 속성과 객체 속성들을 미리 식별합니다.
        /// </summary>
        private void ScanForArrayProperties(XElement element, string parentPath)
        {
            if (!element.HasElements) return;
            
            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToDictionary(g => g.Key, g => g.ToList());
            
            foreach (var group in childGroups)
            {
                var propertyPath = string.IsNullOrEmpty(parentPath) ? group.Key : $"{parentPath}.{group.Key}";
                
                if (group.Value.Count > 1)
                {
                    // 중복 요소 발견 시 배열 속성으로 등록
                    arrayProperties.Add(propertyPath);
                    Logger.Information($"[XmlToYamlConverter] 배열 속성 식별: {propertyPath} ({group.Value.Count}개 요소)");
                }
                
                // 객체 속성 식별: 속성이나 자식 요소가 있는 경우
                foreach (var child in group.Value)
                {
                    if (child.HasAttributes || child.HasElements)
                    {
                        objectProperties.Add(propertyPath);
                        Logger.Information($"[XmlToYamlConverter] 객체 속성 식별: {propertyPath} (속성/자식 요소 포함)");
                        break; // 하나라도 객체면 전체가 객체
                    }
                }
                
                // 재귀적으로 자식 요소들도 스캔
                foreach (var child in group.Value)
                {
                    ScanForArrayProperties(child, propertyPath);
                }
            }
        }

        /// <summary>
        /// XML 요소를 Dictionary로 변환합니다.
        /// </summary>
        private object ConvertElementToDictionary(XElement element, string parentPath = "")
        {
            // 객체 속성으로 식별된 경우 단순 텍스트도 객체로 변환
            if (objectProperties.Contains(parentPath) && !element.HasAttributes && !element.HasElements)
            {
                // 단순 텍스트를 객체로 변환
                var textDict = new Dictionary<string, object>();
                textDict[TextContentKey] = element.Value;
                Logger.Information($"[XmlToYamlConverter] 단순 텍스트를 객체로 변환: {parentPath} = '{element.Value}' -> {{__text: '{element.Value}'}}");
                return textDict;
            }
            
            // 속성이나 자식 요소가 있는 경우 딕셔너리로 처리
            if (element.HasAttributes || element.HasElements)
            {
                var dict = new Dictionary<string, object>();
                
                // 속성 처리
                foreach (var attr in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
                {
                    dict[AttributePrefix + attr.Name.LocalName] = attr.Value;
                }
                
                // 텍스트 내용이 있는 경우
                if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
                {
                    dict[TextContentKey] = element.Value;
                }
                
                // 자식 요소 처리
                var childGroups = element.Elements()
                    .GroupBy(e => e.Name.LocalName)
                    .ToDictionary(g => g.Key, g => g.ToList());
                
                foreach (var group in childGroups)
                {
                    var propertyPath = string.IsNullOrEmpty(parentPath) ? group.Key : $"{parentPath}.{group.Key}";
                    
                    if (arrayProperties.Contains(propertyPath))
                    {
                        // 배열 속성으로 식별된 경우 항상 배열로 처리 (단일 요소라도)
                        var list = new List<object>();
                        foreach (var child in group.Value)
                        {
                            list.Add(ConvertElementToDictionary(child, propertyPath));
                        }
                        dict[group.Key] = list;
                        Logger.Information($"[XmlToYamlConverter] 배열 속성 {group.Key} 처리: {group.Value.Count}개 (경로: {propertyPath})");
                    }
                    else
                    {
                        // 배열 속성이 아닌 단일 요소는 객체로 처리
                        var child = group.Value.First();
                        dict[group.Key] = ConvertElementToDictionary(child, propertyPath);
                        Logger.Information($"[XmlToYamlConverter] 단일 객체 {group.Key} 처리 (경로: {propertyPath})");
                    }
                }
                
                return dict;
            }
            
            // 일반 단순 텍스트 요소
            return element.Value;
        }

    }
}