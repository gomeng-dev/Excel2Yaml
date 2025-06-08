using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ExcelToYamlAddin.Logging;
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
                
                // XML 파싱
                var doc = XDocument.Parse(xmlContent);
                var root = doc.Root;
                
                if (root == null)
                {
                    throw new InvalidOperationException("XML 루트 요소를 찾을 수 없습니다.");
                }
                
                // YAML 문서 생성
                var rootNode = ConvertElementToYaml(root);
                var yamlDocument = new YamlDocument(rootNode);
                
                // YAML 문자열로 변환
                var yamlStream = new YamlStream(yamlDocument);
                var writer = new System.IO.StringWriter();
                yamlStream.Save(writer, false);
                
                var yamlContent = writer.ToString();
                Logger.Information("XML to YAML 변환 완료");
                
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
                
                // XML 파싱
                var doc = XDocument.Parse(xmlContent);
                var root = doc.Root;
                
                if (root == null)
                {
                    throw new InvalidOperationException("XML 루트 요소를 찾을 수 없습니다.");
                }
                
                // 루트 요소를 딕셔너리로 변환
                var result = new Dictionary<string, object>();
                result[root.Name.LocalName] = ConvertElementToDictionary(root);
                
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
                mapping.Add(AttributePrefix + attr.Name.LocalName, new YamlScalarNode(attr.Value));
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
        /// XML 요소를 Dictionary로 변환합니다.
        /// </summary>
        private object ConvertElementToDictionary(XElement element)
        {
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
                    if (group.Value.Count > 1)
                    {
                        // 배열인 경우
                        var list = new List<object>();
                        foreach (var child in group.Value)
                        {
                            list.Add(ConvertElementToDictionary(child));
                        }
                        dict[group.Key] = list;
                    }
                    else
                    {
                        // 단일 요소
                        var child = group.Value.First();
                        dict[group.Key] = ConvertElementToDictionary(child);
                    }
                }
                
                return dict;
            }
            
            // 단순 텍스트 요소
            return element.Value;
        }
    }
}