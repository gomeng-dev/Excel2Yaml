using ClosedXML.Excel;
using ExcelToYamlAddin.Infrastructure.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.Services
{
    public class ExcelReader
    {
        private readonly ExcelToYamlConfig config;
        private const string AUTO_GEN_MARKER = "!";

        public ExcelReader(ExcelToYamlConfig config)
        {
            this.config = config ?? new ExcelToYamlConfig();
        }

        public void ProcessExcelFile(string inputPath, string outputPath)
        {
            ProcessExcelFile(inputPath, outputPath, null);
        }

        public void ProcessExcelFile(string inputPath, string outputPath, string targetSheetName)
        {
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
            {
                throw new FileNotFoundException("엑셀 파일을 찾을 수 없습니다.", inputPath);
            }

            try
            {
                using (var workbook = new XLWorkbook(inputPath))
                {
                    var targetSheets = new List<IXLWorksheet>();

                    if (!string.IsNullOrEmpty(targetSheetName))
                    {
                        // 특정 시트만 처리
                        var sheet = workbook.Worksheet(targetSheetName);
                        if (sheet != null)
                        {
                            targetSheets.Add(sheet);
                        }
                        else
                        {
                            throw new InvalidOperationException($"'{targetSheetName}' 시트를 찾을 수 없습니다.");
                        }
                    }
                    else
                    {
                        // 모든 자동 생성 대상 시트 처리
                        targetSheets = ExtractAutoGenTargetSheets(workbook);
                    }

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
                            Debug.WriteLine($"[ExcelReader] YAML 생성 전 config.IncludeEmptyFields 값: {config.IncludeEmptyFields}");

                            // 현재 Sheet 정보 로깅
                            string sheetNameWithMarker = sheet.Name;
                            string cleanSheetName = RemoveAutoGenMarkerFromSheetName(sheet);
                            Debug.WriteLine($"[ExcelReader] 처리 중인 시트: '{sheetNameWithMarker}' (마커 제외: '{cleanSheetName}')");

                            // 현재 처리 중인 시트에 대한 설정 확인하고 적용 (Ribbon에서 전달된 설정은 활성 시트 기준)
                            bool sheetSetting = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetNameWithMarker);
                            bool excelSheetSetting = ExcelConfigManager.Instance.GetConfigBool(sheetNameWithMarker, "YamlEmptyFields", false);
                            bool globalSetting = Properties.Settings.Default.AddEmptyYamlFields;

                            // 시트별 설정을 우선적으로 확인하고, 없으면 전역 설정 사용
                            bool effectiveSetting = excelSheetSetting || sheetSetting || globalSetting;

                            // 중요: 현재 시트에 맞게 config 값을 업데이트
                            bool originalConfigValue = config.IncludeEmptyFields;
                            config.IncludeEmptyFields = effectiveSetting;

                            Debug.WriteLine($"[ExcelReader] 시트 '{sheetNameWithMarker}'의 설정 - Excel 시트별: {excelSheetSetting}, SheetPath 설정: {sheetSetting}, 전역: {globalSetting}");
                            Debug.WriteLine($"[ExcelReader] 설정 변경: {originalConfigValue} -> {config.IncludeEmptyFields} (시트별 설정 적용)");
                            Debug.WriteLine($"[ExcelReader] config 객체 참조 확인: HashCode={config.GetHashCode()}, 실제 IncludeEmptyFields={config.IncludeEmptyFields}");

                            // 추가 확인을 위해 설정값을 복사하여 로컬 변수로 전달
                            bool localIncludeEmptyFields = config.IncludeEmptyFields; // 값 복사
                            Debug.WriteLine($"[ExcelReader] 로컬 변수로 복사: localIncludeEmptyFields={localIncludeEmptyFields}");

                            string yamlStr = YamlGenerator.Generate(
                                scheme,
                                config.YamlStyle,
                                config.YamlIndentSize,
                                config.YamlPreserveQuotes,
                                config.IncludeEmptyFields);
                            Debug.WriteLine($"[ExcelReader] YAML 생성 완료 (includeEmptyFields: {config.IncludeEmptyFields})");

                            // YAML 결과 확인
                            bool hasEmptyArraySyntax = yamlStr.Contains("[]");
                            Debug.WriteLine($"[ExcelReader] YAML 결과에 빈 배열 표기([]) 포함 여부: {hasEmptyArraySyntax}");

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