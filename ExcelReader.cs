using ClosedXML.Excel;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Collections.Generic;
using ExcelToJsonAddin.Config;
using ExcelToJsonAddin.Core;
using Newtonsoft.Json;

namespace ExcelToJsonAddin
{
    public class ExcelReader
    {
        private readonly ExcelToJsonConfig config;
        private const string AUTO_GEN_MARKER = "!";

        public ExcelReader(ExcelToJsonConfig config)
        {
            this.config = config ?? new ExcelToJsonConfig();
        }

        public void ProcessExcelFile(string inputPath, string outputPath)
        {
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
            {
                throw new FileNotFoundException("엑셀 파일을 찾을 수 없습니다.", inputPath);
            }

            try
            {
                using (var workbook = new XLWorkbook(inputPath))
                {
                    var targetSheets = ExtractAutoGenTargetSheets(workbook);
                    var completedSheetNames = new HashSet<string>();

                    foreach (var sheet in targetSheets)
                    {
                        string sheetName = RemoveAutoGenMarkerFromSheetName(sheet);
                        
                        // 중복 시트 검사
                        if (completedSheetNames.Contains(sheetName))
                        {
                            throw new InvalidOperationException($"'{sheetName}' 시트가 중복되었습니다!");
                        }
                        
                        completedSheetNames.Add(sheetName);
                        
                        // 출력 파일 경로 설정
                        string outputDir = Path.GetDirectoryName(outputPath);
                        string baseFileName = Path.GetFileNameWithoutExtension(outputPath);
                        string ext = config.OutputFormat == OutputFormat.Json ? ".json" : ".yaml";
                        
                        // 시트별 출력 파일 경로
                        string sheetOutputPath;
                        if (targetSheets.Count > 1)
                        {
                            // 여러 시트가 있는 경우 시트 이름으로 파일 생성
                            sheetOutputPath = Path.Combine(outputDir, $"{sheetName}{ext}");
                        }
                        else
                        {
                            // 단일 시트인 경우 지정된 경로 사용
                            sheetOutputPath = outputPath;
                        }
                        
                        // 데이터 파싱을 위한 스키마 파서와 생성기
                        var scheme = new Scheme(sheet);
                        
                        if (config.OutputFormat == OutputFormat.Json)
                        {
                            // JSON 생성 및 저장
                            string jsonStr = Generator.GenerateJson(scheme, config.IncludeEmptyFields);
                            File.WriteAllText(sheetOutputPath, jsonStr);
                        }
                        else
                        {
                            // YAML 생성 및 저장
                            string yamlStr = YamlGenerator.Generate(
                                scheme,
                                config.YamlStyle,
                                config.YamlIndentSize,
                                config.YamlPreserveQuotes,
                                config.IncludeEmptyFields);
                            File.WriteAllText(sheetOutputPath, yamlStr);
                        }

                        // MD5 해시 생성
                        if (config.EnableHashGen)
                        {
                            using (var md5 = MD5.Create())
                            using (var stream = File.OpenRead(sheetOutputPath))
                            {
                                var hash = md5.ComputeHash(stream);
                                var hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                                File.WriteAllText($"{sheetOutputPath}.md5", hashString);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Excel 변환 중 오류: {ex.Message}", ex);
            }
        }

        private string RemoveAutoGenMarkerFromSheetName(IXLWorksheet sheet)
        {
            return sheet.Name.Replace(AUTO_GEN_MARKER, "");
        }

        private List<IXLWorksheet> ExtractAutoGenTargetSheets(XLWorkbook workbook)
        {
            var targetSheets = new List<IXLWorksheet>();
            foreach (var sheet in workbook.Worksheets)
            {
                if (sheet.Name.StartsWith(AUTO_GEN_MARKER))
                {
                    targetSheets.Add(sheet);
                }
            }
            return targetSheets;
        }
    }
} 