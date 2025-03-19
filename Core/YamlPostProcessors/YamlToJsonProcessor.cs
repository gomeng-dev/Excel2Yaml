using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace ExcelToYamlAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일을 JSON으로 변환하는 클래스
    /// </summary>
    public static class YamlToJsonProcessor
    {
        /// <summary>
        /// YAML 파일을 JSON으로 변환하는 메서드
        /// </summary>
        /// <param name="yamlFilePath">변환할 YAML 파일 경로</param>
        /// <param name="jsonFilePath">저장할 JSON 파일 경로</param>
        /// <returns>변환 성공 여부</returns>
        public static bool ConvertYamlToJson(string yamlFilePath, string jsonFilePath)
        {
            if (!File.Exists(yamlFilePath))
            {
                Debug.WriteLine($"[YamlToJsonProcessor] YAML 파일이 존재하지 않습니다: {yamlFilePath}");
                return false;
            }

            try
            {
                // 디렉토리 생성 확인
                string jsonDir = Path.GetDirectoryName(jsonFilePath);
                if (!Directory.Exists(jsonDir))
                {
                    Directory.CreateDirectory(jsonDir);
                }

                // YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlFilePath);

                // YAML을 객체로 변환
                var deserializer = new DeserializerBuilder().Build();
                var yamlObject = deserializer.Deserialize<object>(yamlContent);

                // 객체를 JSON으로 변환
                var jsonContent = JsonConvert.SerializeObject(yamlObject, Formatting.Indented);

                // JSON 파일로 저장
                File.WriteAllText(jsonFilePath, jsonContent);
                
                Debug.WriteLine($"[YamlToJsonProcessor] 변환 완료: {yamlFilePath} -> {jsonFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlToJsonProcessor] 변환 중 오류 발생: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 여러 YAML 파일을 JSON으로 일괄 변환하는 메서드
        /// </summary>
        /// <param name="yamlJsonPairs">YAML-JSON 파일 경로 쌍 목록</param>
        /// <returns>변환 성공한 파일 경로 목록</returns>
        public static List<string> BatchConvertYamlToJson(List<Tuple<string, string>> yamlJsonPairs)
        {
            List<string> convertedFiles = new List<string>();

            if (yamlJsonPairs == null || yamlJsonPairs.Count == 0)
            {
                Debug.WriteLine("[YamlToJsonProcessor] 변환할 파일이 없습니다.");
                return convertedFiles;
            }

            foreach (var pair in yamlJsonPairs)
            {
                string yamlFilePath = pair.Item1;
                string jsonFilePath = pair.Item2;

                if (ConvertYamlToJson(yamlFilePath, jsonFilePath))
                {
                    convertedFiles.Add(jsonFilePath);
                }
            }

            return convertedFiles;
        }

        /// <summary>
        /// 지정된 디렉토리의 모든 YAML 파일을 대상 디렉토리에 JSON으로 변환하는 메서드
        /// </summary>
        /// <param name="sourceDir">YAML 파일이 있는 디렉토리 경로</param>
        /// <param name="targetDir">JSON 파일을 저장할 디렉토리 경로</param>
        /// <param name="searchPattern">검색 패턴 (기본값: *.yaml)</param>
        /// <param name="preserveDirectoryStructure">디렉토리 구조 유지 여부</param>
        /// <returns>변환 성공한 파일 경로 목록</returns>
        public static List<string> ConvertAllYamlFilesInDirectory(
            string sourceDir, 
            string targetDir, 
            string searchPattern = "*.yaml",
            bool preserveDirectoryStructure = false)
        {
            List<string> convertedFiles = new List<string>();

            if (!Directory.Exists(sourceDir))
            {
                Debug.WriteLine($"[YamlToJsonProcessor] 소스 디렉토리가 없습니다: {sourceDir}");
                return convertedFiles;
            }

            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // YAML 파일 목록 가져오기
            var yamlFiles = Directory.GetFiles(sourceDir, searchPattern, SearchOption.AllDirectories);
            
            foreach (var yamlFile in yamlFiles)
            {
                try
                {
                    string relativePath = yamlFile.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar);
                    string jsonFileName = Path.GetFileNameWithoutExtension(relativePath) + ".json";
                    
                    string jsonFilePath;
                    if (preserveDirectoryStructure)
                    {
                        // 디렉토리 구조 유지
                        string relativeDir = Path.GetDirectoryName(relativePath);
                        string targetSubDir = Path.Combine(targetDir, relativeDir);
                        
                        if (!Directory.Exists(targetSubDir))
                        {
                            Directory.CreateDirectory(targetSubDir);
                        }
                        
                        jsonFilePath = Path.Combine(targetSubDir, jsonFileName);
                    }
                    else
                    {
                        // 모든 파일을 대상 디렉토리에 직접 저장
                        jsonFilePath = Path.Combine(targetDir, jsonFileName);
                    }

                    if (ConvertYamlToJson(yamlFile, jsonFilePath))
                    {
                        convertedFiles.Add(jsonFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[YamlToJsonProcessor] '{yamlFile}' 파일 처리 중 오류 발생: {ex.Message}");
                }
            }

            return convertedFiles;
        }
    }
} 