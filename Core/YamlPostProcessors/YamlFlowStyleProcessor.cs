using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.RepresentationModel;
using YamlDotNet.Core.Events;

namespace ExcelToJsonAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일의 특정 필드를 Flow 스타일로 변환하는 프로세서입니다.
    /// </summary>
    public class YamlFlowStyleProcessor
    {
        private readonly Dictionary<string, List<int>> flowStyleFields;
        private readonly Dictionary<string, List<int>> flowStyleItemsFields;

        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="flowStyleFields">Flow 스타일로 변환할 필드명 (쉼표로 구분)</param>
        /// <param name="flowStyleItemsFields">Flow 스타일로 변환할 항목 필드명 (쉼표로 구분)</param>
        public YamlFlowStyleProcessor(string flowStyleFields = "", string flowStyleItemsFields = "")
        {
            this.flowStyleFields = ParseFields(flowStyleFields);
            this.flowStyleItemsFields = ParseFields(flowStyleItemsFields);
        }

        /// <summary>
        /// 설정 문자열을 파싱하여 프로세서를 생성합니다.
        /// </summary>
        /// <param name="flowStyleConfig">설정 문자열</param>
        /// <returns>YamlFlowStyleProcessor 인스턴스</returns>
        public static YamlFlowStyleProcessor FromConfigString(string flowStyleConfig)
        {
            string flowStyleFields = "";
            string flowStyleItemsFields = "";
            
            if (!string.IsNullOrWhiteSpace(flowStyleConfig))
            {
                string[] parts = flowStyleConfig.Split('|');
                if (parts.Length >= 1)
                    flowStyleFields = parts[0]; // 빈 문자열도 그대로 사용
                if (parts.Length >= 2)
                    flowStyleItemsFields = parts[1]; // 빈 문자열도 그대로 사용
            }
            
            return new YamlFlowStyleProcessor(flowStyleFields, flowStyleItemsFields);
        }

        /// <summary>
        /// 필드명 문자열을 파싱하여 Dictionary로 변환합니다.
        /// </summary>
        /// <param name="fieldsString">필드명 문자열 (쉼표로 구분)</param>
        /// <returns>파싱된 필드 Dictionary</returns>
        private Dictionary<string, List<int>> ParseFields(string fieldsString)
        {
            var fieldDict = new Dictionary<string, List<int>>();
            
            if (string.IsNullOrWhiteSpace(fieldsString))
                return fieldDict;
            
            var fields = fieldsString.Split(',').Select(f => f.Trim()).Where(f => !string.IsNullOrWhiteSpace(f));
            var pattern = new Regex(@"^(?<name>\w+)(\[(?<index>\d+)\])?$");
            
            foreach (var field in fields)
            {
                var match = pattern.Match(field);
                if (match.Success)
                {
                    string name = match.Groups["name"].Value;
                    string indexStr = match.Groups["index"].Value;
                    
                    if (!fieldDict.ContainsKey(name))
                    {
                        fieldDict[name] = new List<int>();
                    }
                    
                    if (!string.IsNullOrEmpty(indexStr) && int.TryParse(indexStr, out int index))
                    {
                        fieldDict[name].Add(index);
                    }
                    else
                    {
                        // indexStr이 비어있으면 모든 인덱스를 의미 (null 대신 빈 리스트 사용)
                        // 빈 리스트는 "모든 인덱스"를 의미
                    }
                }
                else
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 경고: 필드명 '{field}'이(가) 올바른 형식이 아닙니다. 무시됩니다.");
                }
            }
            
            return fieldDict;
        }

        /// <summary>
        /// YAML 파일을 Flow 스타일로 처리합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlFile(string yamlPath)
        {
            try
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] YAML 파일 처리: {yamlPath}");
                Debug.WriteLine($"[YamlFlowStyleProcessor] Flow 스타일 필드: {string.Join(", ", flowStyleFields.Keys)}");
                Debug.WriteLine($"[YamlFlowStyleProcessor] Flow 스타일 항목 필드: {string.Join(", ", flowStyleItemsFields.Keys)}");

                // Flow 스타일 필드와 Flow 스타일 항목 필드가 모두 비어있으면 처리하지 않음
                if (flowStyleFields.Count == 0 && flowStyleItemsFields.Count == 0)
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] Flow 스타일 필드가 모두 비어있어 처리를 중단합니다.");
                    return true; // 처리하지 않더라도 성공으로 처리
                }

                if (!File.Exists(yamlPath))
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 오류: YAML 파일을 찾을 수 없습니다.");
                    return false;
                }

                // 1. YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlPath);
                
                if (string.IsNullOrWhiteSpace(yamlContent))
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 오류: YAML 파일이 비어 있습니다.");
                    return false;
                }

                try
                {
                    // YamlDotNet의 RepresentationModel 방식으로 시도
                    return ProcessWithRepresentationModel(yamlPath, yamlContent);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] RepresentationModel 방식 처리 중 오류: {ex.Message}");
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 텍스트 기반 방식으로 대체합니다.");
                    
                    // 실패하면 텍스트 기반 방식으로 처리
                    return ProcessYamlTextDirectly(yamlPath, yamlContent);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 처리 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[YamlFlowStyleProcessor] 스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// YamlDotNet의 RepresentationModel을 사용해 YAML을 처리합니다.
        /// </summary>
        /// <param name="yamlPath">YAML 파일 경로</param>
        /// <param name="yamlContent">YAML 내용</param>
        /// <returns>처리 성공 여부</returns>
        private bool ProcessWithRepresentationModel(string yamlPath, string yamlContent)
        {
            // YamlDotNet을 사용한 YAML 파싱
            var yaml = new YamlStream();
            using (var reader = new StringReader(yamlContent))
            {
                yaml.Load(reader);
            }

            if (yaml.Documents.Count == 0 || yaml.Documents[0].RootNode == null)
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 오류: YAML 문서가 비어 있습니다.");
                return false;
            }

            // 루트 노드 처리
            var rootNode = yaml.Documents[0].RootNode;
            
            // 시퀀스인 경우
            if (rootNode is YamlSequenceNode rootSequence)
            {
                ProcessYamlSequence(rootSequence);
            }
            // 매핑인 경우
            else if (rootNode is YamlMappingNode rootMapping)
            {
                ProcessYamlMapping(rootMapping, new Dictionary<string, int>());
            }

            // 처리된 YAML 저장
            string processedYaml;

            using (var writer = new StringWriter())
            {
                // 들여쓰기 설정 - 2칸으로 설정
                var emitterSettings = new YamlDotNet.Core.EmitterSettings();
                emitterSettings = emitterSettings.WithBestIndent(2);
                emitterSettings = emitterSettings.WithIndentedSequences(); // 시퀀스 인덴트 적용
                
                var emitter = new YamlDotNet.Core.Emitter(writer, emitterSettings);
                yaml.Save(emitter, false);
                processedYaml = writer.ToString();
            }
            
            // Flow 스타일 내부에 인덴트 적용
            processedYaml = ApplyIndentationToFlowStyle(processedYaml);
            
            // 저장
            File.WriteAllText(yamlPath, processedYaml);

            Debug.WriteLine($"[YamlFlowStyleProcessor] YAML 파일 Flow 스타일 처리 완료: {yamlPath}");
            return true;
        }

        /// <summary>
        /// YAML 시퀀스 노드를 처리합니다.
        /// </summary>
        /// <param name="sequenceNode">처리할 시퀀스 노드</param>
        private void ProcessYamlSequence(YamlSequenceNode sequenceNode)
        {
            foreach (var node in sequenceNode.Children)
            {
                if (node is YamlMappingNode mappingNode)
                {
                    ProcessYamlMapping(mappingNode, new Dictionary<string, int>());
                }
            }
        }

        /// <summary>
        /// YAML 매핑 노드를 처리합니다.
        /// </summary>
        /// <param name="mappingNode">처리할 매핑 노드</param>
        /// <param name="fieldCounters">필드 카운터</param>
        private void ProcessYamlMapping(YamlMappingNode mappingNode, Dictionary<string, int> fieldCounters)
        {
            foreach (var entry in mappingNode.Children)
            {
                if (!(entry.Key is YamlScalarNode keyNode))
                    continue;

                string key = keyNode.Value;
                
                // 필드 카운팅
                if (!fieldCounters.ContainsKey(key))
                {
                    fieldCounters[key] = 0;
                }
                int currentIndex = fieldCounters[key]++;

                // Flow 필드 처리 - 필드 자체를 Flow 스타일로 설정
                if (flowStyleFields.ContainsKey(key))
                {
                    var indices = flowStyleFields[key];
                    if (indices.Count == 0 || indices.Contains(currentIndex))
                    {
                        // 필드 값 자체를 Flow 스타일로 설정
                        if (entry.Value is YamlMappingNode valueMapping)
                        {
                            valueMapping.Style = YamlDotNet.Core.Events.MappingStyle.Flow;
                            
                            // 인덴트를 유지하기 위해 각 자식 노드의 스타일도 설정
                            foreach (var childEntry in valueMapping.Children)
                            {
                                if (childEntry.Value is YamlMappingNode childMapping)
                                {
                                    childMapping.Style = YamlDotNet.Core.Events.MappingStyle.Flow;
                                }
                                else if (childEntry.Value is YamlSequenceNode childSequence)
                                {
                                    childSequence.Style = YamlDotNet.Core.Events.SequenceStyle.Flow;
                                }
                            }
                        }
                        else if (entry.Value is YamlSequenceNode valueSequence)
                        {
                            valueSequence.Style = YamlDotNet.Core.Events.SequenceStyle.Flow;
                            
                            // 인덴트를 유지하기 위해 각 자식 노드의 스타일도 설정
                            for (int i = 0; i < valueSequence.Children.Count; i++)
                            {
                                if (valueSequence.Children[i] is YamlMappingNode childMapping)
                                {
                                    childMapping.Style = YamlDotNet.Core.Events.MappingStyle.Flow;
                                }
                                else if (valueSequence.Children[i] is YamlSequenceNode childSequence)
                                {
                                    childSequence.Style = YamlDotNet.Core.Events.SequenceStyle.Flow;
                                }
                            }
                        }
                    }
                }

                // Flow 항목 필드 처리 - 해당 필드의 자식(항목)들을 Flow 스타일로 설정
                if (flowStyleItemsFields.ContainsKey(key) && entry.Value is YamlSequenceNode itemsSequence)
                {
                    var indices = flowStyleItemsFields[key];
                    if (indices.Count == 0 || indices.Contains(currentIndex))
                    {
                        // 자식 항목들만 Flow 스타일로 설정 (시퀀스 자체는 Block 스타일로 유지)
                        for (int i = 0; i < itemsSequence.Children.Count; i++)
                        {
                            if (itemsSequence.Children[i] is YamlMappingNode itemMapping)
                            {
                                // 각 항목을 Flow 스타일로 설정
                                itemMapping.Style = YamlDotNet.Core.Events.MappingStyle.Flow;
                            }
                            else if (itemsSequence.Children[i] is YamlSequenceNode itemSequence)
                            {
                                // 각 항목(시퀀스)을 Flow 스타일로 설정
                                itemSequence.Style = YamlDotNet.Core.Events.SequenceStyle.Flow;
                            }
                        }
                    }
                }

                // 재귀적으로 처리
                if (entry.Value is YamlMappingNode nextMapping)
                {
                    var nextCounters = new Dictionary<string, int>(fieldCounters);
                    ProcessYamlMapping(nextMapping, nextCounters);
                }
                else if (entry.Value is YamlSequenceNode nextSequence)
                {
                    foreach (var item in nextSequence.Children)
                    {
                        if (item is YamlMappingNode itemMapping)
                        {
                            var nextCounters = new Dictionary<string, int>(fieldCounters);
                            ProcessYamlMapping(itemMapping, nextCounters);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// YAML 텍스트를 직접 처리하여 Flow 스타일을 적용합니다.
        /// </summary>
        /// <param name="yamlPath">YAML 파일 경로</param>
        /// <param name="yamlContent">YAML 내용</param>
        /// <returns>처리 성공 여부</returns>
        private bool ProcessYamlTextDirectly(string yamlPath, string yamlContent)
        {
            try
            {
                string processedYaml = ApplyFlowStyleToText(yamlContent);
                File.WriteAllText(yamlPath, processedYaml);
                Debug.WriteLine($"[YamlFlowStyleProcessor] YAML 파일 Flow 스타일 처리 완료(텍스트 직접 처리): {yamlPath}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 텍스트 직접 처리 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Flow 스타일 내부에 인덴트를 적용합니다.
        /// </summary>
        /// <param name="yamlText">YAML 텍스트</param>
        /// <returns>인덴트가 적용된 YAML 텍스트</returns>
        private string ApplyIndentationToFlowStyle(string yamlText)
        {
            // 기존 개행은 유지하되 새로운 개행은 추가하지 않음
            return yamlText;
        }

        /// <summary>
        /// YAML 텍스트에 Flow 스타일을 적용합니다.
        /// </summary>
        /// <param name="yamlText">YAML 텍스트</param>
        /// <returns>Flow 스타일이 적용된 YAML 텍스트</returns>
        private string ApplyFlowStyleToText(string yamlText)
        {
            string processedYaml = yamlText;
            
            // 필드에 Flow 스타일 적용 - 필드 자체를 Flow 스타일로 출력
            foreach (var field in flowStyleFields.Keys)
            {
                // 정규식 패턴으로 해당 필드의 매핑을 찾아 Flow 스타일로 변환
                // 필드명 다음에 중괄호가 오는 형태 (매핑)
                string pattern = $@"({field}:)(\s*\n\s*)(\{{\s*\n)";
                processedYaml = Regex.Replace(processedYaml, pattern, "$1 { ", RegexOptions.Multiline);
                
                // 중괄호 닫기 스타일 변경
                pattern = $@"(\n\s*\}})(\s*\n)";
                processedYaml = Regex.Replace(processedYaml, pattern, " }", RegexOptions.Multiline);
                
                // 필드명 다음에 대괄호가 오는 형태 (시퀀스)
                pattern = $@"({field}:)(\s*\n\s*)(\[\s*\n)";
                processedYaml = Regex.Replace(processedYaml, pattern, "$1 [ ", RegexOptions.Multiline);
                
                // 대괄호 닫기 스타일 변경
                pattern = $@"(\n\s*\])(\s*\n)";
                processedYaml = Regex.Replace(processedYaml, pattern, " ]", RegexOptions.Multiline);
                
                // 쉼표 뒤에 자동 개행 추가하는 부분 제거
            }
            
            // Flow 항목 필드에 Flow 스타일 적용 - 자식 항목들만 Flow 스타일로 출력
            foreach (var field in flowStyleItemsFields.Keys)
            {
                // 해당 필드 내의 시퀀스 항목이 매핑인 경우 Flow 스타일로 변환
                // 시퀀스의 각 항목이 매핑인 경우를 찾음 (- 다음에 중괄호가 오는 형태)
                string fieldPattern = $@"({field}:(?:\s*\n)(?:\s*-).*?)";
                var matches = Regex.Matches(processedYaml, fieldPattern, RegexOptions.Singleline);
                
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        string itemPattern = @"(\s*-\s*\n\s*)(\{\s*\n)";
                        string itemReplacement = "$1{ ";
                        string itemText = match.Value;
                        string processedItemText = Regex.Replace(itemText, itemPattern, itemReplacement, RegexOptions.Multiline);
                        
                        // 중괄호 닫기 스타일 변경
                        processedItemText = Regex.Replace(processedItemText, @"(\n\s*\})(\s*\n)", " }", RegexOptions.Multiline);
                        
                        // 쉼표 뒤에 개행과 들여쓰기 추가하는 부분 제거
                        
                        // 원본 텍스트 교체
                        processedYaml = processedYaml.Replace(itemText, processedItemText);
                    }
                }
                
                // 시퀀스 항목이 다시 시퀀스인 경우를 처리
                fieldPattern = $@"({field}:(?:\s*\n)(?:\s*-).*?)";
                matches = Regex.Matches(processedYaml, fieldPattern, RegexOptions.Singleline);
                
                if (matches.Count > 0)
                {
                    foreach (Match match in matches)
                    {
                        string itemPattern = @"(\s*-\s*\n\s*)(\[\s*\n)";
                        string itemReplacement = "$1[ ";
                        string itemText = match.Value;
                        string processedItemText = Regex.Replace(itemText, itemPattern, itemReplacement, RegexOptions.Multiline);
                        
                        // 대괄호 닫기 스타일 변경
                        processedItemText = Regex.Replace(processedItemText, @"(\n\s*\])(\s*\n)", " ]", RegexOptions.Multiline);
                        
                        // 쉼표 뒤에 개행과 들여쓰기 추가하는 부분 제거
                        
                        // 원본 텍스트 교체
                        processedYaml = processedYaml.Replace(itemText, processedItemText);
                    }
                }
            }
            
            return processedYaml;
        }

        /// <summary>
        /// 설정 문자열로부터 YAML 파일을 Flow 스타일로 처리합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="flowStyleConfig">Flow 스타일 설정 문자열</param>
        /// <returns>처리 성공 여부</returns>
        public static bool ProcessYamlFileFromConfig(string yamlPath, string flowStyleConfig)
        {
            // 설정 문자열이 비어있으면 후처리 실행하지 않음
            if (string.IsNullOrWhiteSpace(flowStyleConfig))
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 설정 문자열이 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            YamlFlowStyleProcessor processor = FromConfigString(flowStyleConfig);
            return processor.ProcessYamlFile(yamlPath);
        }
    }
} 