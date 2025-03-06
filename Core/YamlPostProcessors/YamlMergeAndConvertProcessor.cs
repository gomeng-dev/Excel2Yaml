using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using ExcelToYamlAddin.Config;
using ExcelToYamlAddin.Logging;
using ExcelToYamlAddin.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.RepresentationModel;
using System.Text.RegularExpressions;

namespace ExcelToYamlAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 병합 및 변환 프로세서
    /// YamlDotNet을 직접 사용하여 YAML 구조를 유지하면서 병합하고 직렬화합니다.
    /// </summary>
    public class YamlMergeAndConvertProcessor
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlMergeAndConvertProcessor>();

        // 기본 설정들
        private readonly string idPath;
        private readonly string mergePaths;
        private readonly string keyPaths;
        private readonly string arrayFieldPaths;
        private readonly YamlStyle outputStyle;
        private readonly int indentSize;
        private readonly bool preserveQuotes;
        private readonly bool includeEmptyFields;
        private readonly bool debugMode;

        // 설정 상수
        private const char KEY_PATH_SEPARATOR = '.';
        private const char STRATEGY_SEPARATOR = ':';
        private const char ARRAY_PATH_SEPARATOR = ';';

        public YamlMergeAndConvertProcessor(
            string idPath = "",
            string mergePaths = "",
            string keyPaths = "",
            string arrayFieldPaths = "",
            YamlStyle outputStyle = YamlStyle.Block,
            int indentSize = 2,
            bool preserveQuotes = false,
            bool includeEmptyFields = false,
            bool debugMode = false)
        {
            this.idPath = idPath;
            this.mergePaths = mergePaths;
            this.keyPaths = keyPaths;
            this.arrayFieldPaths = arrayFieldPaths;
            this.outputStyle = outputStyle;
            this.indentSize = indentSize;
            this.preserveQuotes = preserveQuotes;
            this.includeEmptyFields = includeEmptyFields;
            this.debugMode = debugMode;
        }

        /// <summary>
        /// 설정 문자열에서 프로세서를 생성합니다.
        /// </summary>
        /// <param name="configString">설정 문자열 (형식: "idPath|mergePaths|keyPaths|arrayFieldPaths")</param>
        /// <param name="outputStyle">출력 YAML 스타일</param>
        /// <param name="indentSize">들여쓰기 크기</param>
        /// <param name="preserveQuotes">따옴표 유지 여부</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부</param>
        /// <param name="debugMode">디버그 모드 여부</param>
        /// <returns>YamlMergeAndConvertProcessor 인스턴스</returns>
        public static YamlMergeAndConvertProcessor FromConfigString(
            string configString,
            YamlStyle outputStyle = YamlStyle.Block,
            int indentSize = 2,
            bool preserveQuotes = false,
            bool includeEmptyFields = false,
            bool debugMode = false)
        {
            string idPath = "";
            string mergePaths = "";
            string keyPaths = "";
            string arrayFieldPaths = "";

            if (!string.IsNullOrWhiteSpace(configString))
            {
                string[] parts = configString.Split('|');
                if (parts.Length >= 1)
                    idPath = parts[0];
                if (parts.Length >= 2)
                    mergePaths = parts[1];
                if (parts.Length >= 3)
                    keyPaths = parts[2];
                if (parts.Length >= 4)
                    arrayFieldPaths = parts[3];
            }

            return new YamlMergeAndConvertProcessor(
                idPath, mergePaths, keyPaths, arrayFieldPaths,
                outputStyle, indentSize, preserveQuotes, includeEmptyFields, debugMode);
        }

        /// <summary>
        /// YAML 파일을 처리합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlFile(string yamlPath)
        {
            try
            {
                if (debugMode)
                {
                    Logger.Debug($"YAML 파일 처리 시작: {yamlPath}");
                    Logger.Debug($"ID 경로: {idPath}, 병합 경로: {mergePaths}, 키 경로: {keyPaths}, 배열 필드 경로: {arrayFieldPaths}");
                }

                // 파일 존재 여부 확인
                if (!File.Exists(yamlPath))
                {
                    Logger.Error($"YAML 파일을 찾을 수 없습니다: {yamlPath}");
                    return false;
                }

                // YAML 파일 내용 읽기
                string yamlContent = File.ReadAllText(yamlPath);

                // YAML 파싱
                var yaml = new YamlStream();
                using (var reader = new StringReader(yamlContent))
                {
                    yaml.Load(reader);
                }

                // 루트 노드가 존재하는지 확인
                if (yaml.Documents.Count == 0 || yaml.Documents[0].RootNode == null)
                {
                    Logger.Error("YAML 문서에 루트 노드가 없습니다.");
                    return false;
                }

                // 루트 노드 가져오기
                var rootNode = yaml.Documents[0].RootNode;

                // 병합 경로가 있을 경우 처리
                if (!string.IsNullOrEmpty(mergePaths))
                {
                    // 병합 경로 목록
                    string[] paths = mergePaths.Split(',');
                    foreach (string path in paths)
                    {
                        if (!string.IsNullOrEmpty(path))
                        {
                            // 해당 경로에서 병합 처리 수행
                            ProcessMergePath(rootNode, path.Trim());
                        }
                    }
                }

                // 처리된 YAML 문서 직렬화
                using (var writer = new StringWriter())
                {
                    var yamlStr = new YamlStream(yaml.Documents.ToArray());
                    yamlStr.Save(writer, false);

                    // 최종 YAML 문자열 가져오기
                    string result = writer.ToString();

                    // 추가 포맷팅 처리
                    result = PostProcessYaml(result);

                    // 결과를 원본 파일에 덮어쓰기
                    File.WriteAllText(yamlPath, result);
                }

                if (debugMode)
                {
                    Logger.Debug($"YAML 처리 완료 및 파일 저장: {yamlPath}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "YAML 처리 중 오류 발생");
                return false;
            }
        }

        /// <summary>
        /// 특정 경로에서 병합 처리를 수행합니다.
        /// </summary>
        /// <param name="rootNode">YAML 루트 노드</param>
        /// <param name="mergePath">병합 경로</param>
        private void ProcessMergePath(YamlNode rootNode, string mergePath)
        {
            if (rootNode == null || string.IsNullOrEmpty(mergePath))
                return;

            // 경로 분석
            string[] pathParts = mergePath.Split(KEY_PATH_SEPARATOR);
            
            // 현재 노드 설정
            YamlNode currentNode = rootNode;
            
            // 마지막 경로 부분까지 탐색
            for (int i = 0; i < pathParts.Length; i++)
            {
                string part = pathParts[i];
                
                // 현재 노드가 매핑(객체)인지 확인
                if (currentNode is YamlMappingNode mappingNode)
                {
                    // 키가 있는지 확인
                    if (mappingNode.Children.ContainsKey(new YamlScalarNode(part)))
                    {
                        // 다음 노드로 이동
                        currentNode = mappingNode.Children[new YamlScalarNode(part)];
                        
                        // 마지막 경로 부분인 경우, 여기서 병합 수행
                        if (i == pathParts.Length - 1)
                        {
                            // 배열인 경우에만 병합 수행
                            if (currentNode is YamlSequenceNode sequenceNode)
                            {
                                MergeSequenceNode(sequenceNode);
                            }
                            else
                            {
                                Logger.Warning($"병합 경로의 마지막 노드가 배열이 아닙니다: {mergePath}");
                            }
                        }
                    }
                    else
                    {
                        // 경로가 존재하지 않음
                        Logger.Warning($"YAML 경로를 찾을 수 없습니다: {part} in {mergePath}");
                        return;
                    }
                }
                else if (currentNode is YamlSequenceNode)
                {
                    // 시퀀스(배열) 노드에서는 더 이상 경로 탐색이 불가능
                    Logger.Warning($"배열 노드 내에서 추가 경로 탐색이 불가능합니다: {part} in {mergePath}");
                    return;
                }
                else
                {
                    // 스칼라(기본 값) 노드에서는 더 이상 경로 탐색이 불가능
                    Logger.Warning($"스칼라 노드 내에서 추가 경로 탐색이 불가능합니다: {part} in {mergePath}");
                    return;
                }
            }
        }

        /// <summary>
        /// 시퀀스 노드(배열)의 아이템을 ID 기준으로 병합합니다.
        /// </summary>
        /// <param name="sequenceNode">시퀀스 노드</param>
        private void MergeSequenceNode(YamlSequenceNode sequenceNode)
        {
            if (sequenceNode == null || string.IsNullOrEmpty(idPath))
                return;

            if (debugMode)
            {
                Logger.Debug($"시퀀스 노드 병합 시작, ID 경로: {idPath}");
            }

            // ID 경로 분석
            string[] idPathParts = idPath.Split(KEY_PATH_SEPARATOR);

            // ID 값을 기준으로 노드를 그룹화
            var nodeGroups = new Dictionary<string, List<YamlNode>>();
            
            // 첫 번째 단계: 모든 노드를 ID 값으로 그룹화
            foreach (var node in sequenceNode.Children)
            {
                if (node is YamlMappingNode mappingNode)
                {
                    // ID 값 추출
                    string idValue = ExtractIdValue(mappingNode, idPathParts);
                    
                    if (!string.IsNullOrEmpty(idValue))
                    {
                        if (!nodeGroups.ContainsKey(idValue))
                        {
                            nodeGroups[idValue] = new List<YamlNode>();
                        }
                        
                        nodeGroups[idValue].Add(node);
                    }
                }
            }
            
            // 두 번째 단계: 동일한 ID를 가진 노드들을 병합
            foreach (var group in nodeGroups)
            {
                if (group.Value.Count > 1)
                {
                    // 첫 번째 노드를 기준으로 병합
                    var baseNode = group.Value[0] as YamlMappingNode;
                    
                    // 나머지 노드의 속성들을 기준 노드에 병합
                    for (int i = 1; i < group.Value.Count; i++)
                    {
                        var nodeToMerge = group.Value[i] as YamlMappingNode;
                        MergeMappingNodes(baseNode, nodeToMerge);
                    }
                }
            }
            
            // 세 번째 단계: 시퀀스 노드의 중복 아이템 제거
            var uniqueNodes = new List<YamlNode>();
            var processedIds = new HashSet<string>();
            
            foreach (var node in sequenceNode.Children)
            {
                if (node is YamlMappingNode mappingNode)
                {
                    string idValue = ExtractIdValue(mappingNode, idPathParts);
                    
                    if (!string.IsNullOrEmpty(idValue) && !processedIds.Contains(idValue))
                    {
                        uniqueNodes.Add(node);
                        processedIds.Add(idValue);
                    }
                }
                else
                {
                    // ID가 없는 노드는 그대로 유지
                    uniqueNodes.Add(node);
                }
            }
            
            // 시퀀스 노드 업데이트
            sequenceNode.Children.Clear();
            foreach (var node in uniqueNodes)
            {
                sequenceNode.Add(node);
            }
            
            if (debugMode)
            {
                Logger.Debug($"시퀀스 노드 병합 완료, 결과 아이템 수: {sequenceNode.Children.Count}");
            }
        }

        /// <summary>
        /// 매핑 노드에서 ID 값을 추출합니다.
        /// </summary>
        /// <param name="node">매핑 노드</param>
        /// <param name="idPathParts">ID 경로 부분들</param>
        /// <returns>ID 값</returns>
        private string ExtractIdValue(YamlMappingNode node, string[] idPathParts)
        {
            if (node == null || idPathParts == null || idPathParts.Length == 0)
                return null;
            
            YamlNode currentNode = node;
            
            // ID 경로를 따라 탐색
            for (int i = 0; i < idPathParts.Length; i++)
            {
                string part = idPathParts[i];
                
                if (currentNode is YamlMappingNode mappingNode)
                {
                    // 키가 있는지 확인
                    if (mappingNode.Children.ContainsKey(new YamlScalarNode(part)))
                    {
                        // 다음 노드로 이동
                        currentNode = mappingNode.Children[new YamlScalarNode(part)];
                        
                        // 마지막 경로 부분인 경우, ID 값 반환
                        if (i == idPathParts.Length - 1 && currentNode is YamlScalarNode scalarNode)
                        {
                            return scalarNode.Value;
                        }
                    }
                    else
                    {
                        // 경로가 존재하지 않음
                        return null;
                    }
                }
                else
                {
                    // 더 이상 탐색이 불가능
                    return null;
                }
            }
            
            return null;
        }

        /// <summary>
        /// 두 매핑 노드를 병합합니다.
        /// </summary>
        /// <param name="baseNode">기준 노드</param>
        /// <param name="nodeToMerge">병합할 노드</param>
        private void MergeMappingNodes(YamlMappingNode baseNode, YamlMappingNode nodeToMerge)
        {
            if (baseNode == null || nodeToMerge == null)
                return;
            
            // 키 경로 설정
            Dictionary<string, string> keyPathsDict = ParseKeyPaths();
            
            // 배열 필드 경로 설정
            Dictionary<string, string> arrayFieldPathsDict = ParseArrayFieldPaths();
            
            // 노드 속성 병합
            foreach (var entry in nodeToMerge.Children)
            {
                string key = (entry.Key as YamlScalarNode)?.Value;
                
                if (string.IsNullOrEmpty(key))
                    continue;
                
                // 키 경로에 따른 병합 전략 적용
                string mergeStrategy = GetMergeStrategy(key, keyPathsDict);
                
                if (baseNode.Children.ContainsKey(entry.Key))
                {
                    // 이미 존재하는 키인 경우 병합 전략에 따라 처리
                    var baseValue = baseNode.Children[entry.Key];
                    var mergeValue = entry.Value;
                    
                    if (mergeStrategy == "replace")
                    {
                        // 값 대체
                        baseNode.Children[entry.Key] = mergeValue;
                    }
                    else if (mergeStrategy == "match")
                    {
                        // 하위 값들을 매칭하여 병합 (재귀적으로 처리)
                        if (baseValue is YamlMappingNode baseMappingNode && 
                            mergeValue is YamlMappingNode mergeMappingNode)
                        {
                            MergeMappingNodes(baseMappingNode, mergeMappingNode);
                        }
                        else if (baseValue is YamlSequenceNode baseSequenceNode && 
                                mergeValue is YamlSequenceNode mergeSequenceNode)
                        {
                            // 배열 필드 경로에 따른 처리
                            string arrayStrategy = GetArrayStrategy(key, arrayFieldPathsDict);
                            MergeSequenceNodes(baseSequenceNode, mergeSequenceNode, arrayStrategy);
                        }
                        else
                        {
                            // 타입이 일치하지 않으면 대체
                            baseNode.Children[entry.Key] = mergeValue;
                        }
                    }
                    else
                    {
                        // 기본 전략: match (매칭하여 병합)
                        if (baseValue is YamlMappingNode baseMappingNode && 
                            mergeValue is YamlMappingNode mergeMappingNode)
                        {
                            MergeMappingNodes(baseMappingNode, mergeMappingNode);
                        }
                        else if (baseValue is YamlSequenceNode baseSequenceNode && 
                                mergeValue is YamlSequenceNode mergeSequenceNode)
                        {
                            // 배열 필드 경로에 따른 처리
                            string arrayStrategy = GetArrayStrategy(key, arrayFieldPathsDict);
                            MergeSequenceNodes(baseSequenceNode, mergeSequenceNode, arrayStrategy);
                        }
                        else
                        {
                            // 타입이 일치하지 않으면 대체
                            baseNode.Children[entry.Key] = mergeValue;
                        }
                    }
                }
                else
                {
                    // 존재하지 않는 키인 경우 추가
                    baseNode.Add(entry.Key, entry.Value);
                }
            }
        }

        /// <summary>
        /// 두 시퀀스 노드(배열)를 병합합니다.
        /// </summary>
        /// <param name="baseNode">기준 노드</param>
        /// <param name="nodeToMerge">병합할 노드</param>
        /// <param name="strategy">병합 전략</param>
        private void MergeSequenceNodes(YamlSequenceNode baseNode, YamlSequenceNode nodeToMerge, string strategy)
        {
            if (baseNode == null || nodeToMerge == null)
                return;
            
            if (strategy == "append")
            {
                // 배열 요소 추가
                foreach (var item in nodeToMerge.Children)
                {
                    baseNode.Add(item);
                }
            }
            else if (strategy == "prepend")
            {
                // 배열 앞부분에 요소 추가
                var newChildren = new List<YamlNode>(nodeToMerge.Children);
                newChildren.AddRange(baseNode.Children);
                
                baseNode.Children.Clear();
                foreach (var item in newChildren)
                {
                    baseNode.Add(item);
                }
            }
            else if (strategy == "replace")
            {
                // 배열 내용 대체
                baseNode.Children.Clear();
                foreach (var item in nodeToMerge.Children)
                {
                    baseNode.Add(item);
                }
            }
            else
            {
                // 기본 전략: append
                foreach (var item in nodeToMerge.Children)
                {
                    baseNode.Add(item);
                }
            }
        }

        /// <summary>
        /// 키 경로 문자열을 파싱하여 Dictionary 형태로 반환합니다.
        /// </summary>
        /// <returns>키 경로 Dictionary</returns>
        private Dictionary<string, string> ParseKeyPaths()
        {
            var result = new Dictionary<string, string>();
            
            if (string.IsNullOrEmpty(keyPaths))
                return result;
            
            string[] entries = keyPaths.Split(',');
            foreach (string entry in entries)
            {
                if (string.IsNullOrEmpty(entry))
                    continue;
                
                string[] parts = entry.Split(STRATEGY_SEPARATOR);
                string path = parts[0].Trim();
                string strategy = parts.Length > 1 ? parts[1].Trim() : "match";
                
                result[path] = strategy;
            }
            
            return result;
        }

        /// <summary>
        /// 배열 필드 경로 문자열을 파싱하여 Dictionary 형태로 반환합니다.
        /// </summary>
        /// <returns>배열 필드 경로 Dictionary</returns>
        private Dictionary<string, string> ParseArrayFieldPaths()
        {
            var result = new Dictionary<string, string>();
            
            if (string.IsNullOrEmpty(arrayFieldPaths))
                return result;
            
            string[] entries = arrayFieldPaths.Split(',');
            foreach (string entry in entries)
            {
                if (string.IsNullOrEmpty(entry))
                    continue;
                
                string[] parts = entry.Split(ARRAY_PATH_SEPARATOR);
                string path = parts[0].Trim();
                string strategy = parts.Length > 1 ? parts[1].Trim() : "append";
                
                result[path] = strategy;
            }
            
            return result;
        }

        /// <summary>
        /// 키에 대한 병합 전략을 가져옵니다.
        /// </summary>
        /// <param name="key">키</param>
        /// <param name="keyPathsDict">키 경로 Dictionary</param>
        /// <returns>병합 전략</returns>
        private string GetMergeStrategy(string key, Dictionary<string, string> keyPathsDict)
        {
            if (keyPathsDict.ContainsKey(key))
                return keyPathsDict[key];
            
            // 점 표기법 경로 처리 (key1.key2.key3)
            foreach (var entry in keyPathsDict)
            {
                string path = entry.Key;
                string[] pathParts = path.Split(KEY_PATH_SEPARATOR);
                
                if (pathParts.Length > 0 && pathParts[pathParts.Length - 1] == key)
                {
                    return entry.Value;
                }
            }
            
            return "match"; // 기본 전략
        }

        /// <summary>
        /// 키에 대한 배열 병합 전략을 가져옵니다.
        /// </summary>
        /// <param name="key">키</param>
        /// <param name="arrayFieldPathsDict">배열 필드 경로 Dictionary</param>
        /// <returns>배열 병합 전략</returns>
        private string GetArrayStrategy(string key, Dictionary<string, string> arrayFieldPathsDict)
        {
            if (arrayFieldPathsDict.ContainsKey(key))
                return arrayFieldPathsDict[key];
            
            // 점 표기법 경로 처리 (key1.key2.key3)
            foreach (var entry in arrayFieldPathsDict)
            {
                string path = entry.Key;
                string[] pathParts = path.Split(KEY_PATH_SEPARATOR);
                
                if (pathParts.Length > 0 && pathParts[pathParts.Length - 1] == key)
                {
                    return entry.Value;
                }
            }
            
            return "append"; // 기본 전략
        }

        /// <summary>
        /// YAML 문자열을 후처리합니다.
        /// </summary>
        /// <param name="yaml">원본 YAML 문자열</param>
        /// <returns>후처리된 YAML 문자열</returns>
        private string PostProcessYaml(string yaml)
        {
            // 들여쓰기 문제 수정
            if (indentSize != 2)
            {
                // 현재 YamlDotNet에서는 들여쓰기 크기를 직접 조절할 수 없으므로
                // 출력된 YAML 문자열에서 추가 작업 수행
                // (이 부분은 필요에 따라 구현)
            }
            
            // 따옴표 처리
            if (!preserveQuotes)
            {
                // 불필요한 따옴표 제거 (필요에 따라 구현)
            }
            
            // Flow Style 적용 (outputStyle == YamlStyle.Flow)
            if (outputStyle == YamlStyle.Flow)
            {
                // YamlDotNet은 기본적으로 Block 스타일을 사용하므로
                // Flow 스타일로 변환 작업 수행 (필요에 따라 구현)
            }
            
            // 빈 속성 처리
            if (!includeEmptyFields)
            {
                // 빈 속성 제거 (필요에 따라 구현)
            }
            
            return yaml;
        }

        /// <summary>
        /// 설정 문자열을 사용하여 YAML 파일을 처리합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="configString">설정 문자열 (형식: "idPath|mergePaths|keyPaths|arrayFieldPaths")</param>
        /// <param name="outputStyle">출력 YAML 스타일</param>
        /// <param name="indentSize">들여쓰기 크기</param>
        /// <param name="preserveQuotes">따옴표 유지 여부</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부</param>
        /// <returns>처리 성공 여부</returns>
        public static bool ProcessYamlFileFromConfig(
            string yamlPath,
            string configString,
            YamlStyle outputStyle = YamlStyle.Block,
            int indentSize = 2,
            bool preserveQuotes = false,
            bool includeEmptyFields = false)
        {
            try
            {
                var processor = FromConfigString(
                    configString, outputStyle, indentSize, preserveQuotes, includeEmptyFields, true);
                
                return processor.ProcessYamlFile(yamlPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlMergeAndConvertProcessor] 설정 문자열 처리 중 오류: {ex.Message}");
                Logger.Error(ex, "YAML 파일 처리 중 오류 발생");
                
                // 기존 방식으로 처리 시도
                return YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(yamlPath, configString);
            }
        }
    }
} 