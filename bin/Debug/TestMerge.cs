using System;
using System.IO;
using System.Collections.Generic;
using ExcelToJsonAddin.Core.YamlPostProcessors;

namespace ExcelToJsonAddin
{
    class TestMerge
    {
        static void Main(string[] args)
        {
            Console.WriteLine("YAML 병합 테스트 시작");
            string yamlFilePath = @"AwesomeBoxSetInfo.yaml";
            
            if (!File.Exists(yamlFilePath))
            {
                Console.WriteLine($"파일을 찾을 수 없습니다: {yamlFilePath}");
                return;
            }
            
            // 병합 키로 category를 사용
            string keyPaths = "items.*.category:merge";
            
            // 설정 문자열 구성 (idPath|mergePaths|keyPaths)
            string mergeConfig = $"id|items|{keyPaths}";
            
            Console.WriteLine($"병합 설정: {mergeConfig}");
            
            // YAML 파일 처리
            bool result = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(yamlFilePath, mergeConfig);
            
            if (result)
            {
                Console.WriteLine("병합 성공");
            }
            else
            {
                Console.WriteLine("병합 실패");
            }
        }
    }
} 