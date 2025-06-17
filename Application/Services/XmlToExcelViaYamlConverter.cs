using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Core.YamlToExcel;
using ExcelToYamlAddin.Infrastructure.Logging;

namespace ExcelToYamlAddin.Core
{
    /// <summary>
    /// XML을 YAML로 변환한 후 Excel로 변환하는 통합 컨버터
    /// XmlToExcelConverter를 대체하는 새로운 구현
    /// </summary>
    public class XmlToExcelViaYamlConverter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<XmlToExcelViaYamlConverter>();
        
        private readonly XmlToYamlConverter _xmlToYamlConverter;
        private readonly YamlToExcelConverter _yamlToExcelConverter;

        public XmlToExcelViaYamlConverter()
        {
            _xmlToYamlConverter = new XmlToYamlConverter();
            _yamlToExcelConverter = new YamlToExcelConverter();
        }

        /// <summary>
        /// XML 파일을 Excel 워크북으로 변환합니다.
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <param name="sheetName">생성할 시트 이름 (기본값: XML 루트 요소 이름)</param>
        /// <returns>변환된 Excel 워크북</returns>
        public IXLWorkbook ConvertToExcel(string xmlContent, string sheetName = null)
        {
            try
            {
                Logger.Information("XML to Excel (via YAML) 변환 시작");
                
                // 1단계: XML을 YAML로 변환
                Logger.Information("1단계: XML → YAML 변환");
                var yamlContent = _xmlToYamlConverter.ConvertToYaml(xmlContent);
                
                if (string.IsNullOrWhiteSpace(yamlContent))
                {
                    throw new InvalidOperationException("XML을 YAML로 변환한 결과가 비어있습니다.");
                }
                
                // 콘솔에 YAML 내용 출력 (디버깅용)
                Logger.Information("========== 생성된 YAML 내용 시작 ==========");
                Console.WriteLine(yamlContent);
                Logger.Information("========== 생성된 YAML 내용 끝 ==========");
                
                Logger.Debug($"변환된 YAML:\n{yamlContent}");
                
                // 2단계: 시트 이름 결정
                if (string.IsNullOrEmpty(sheetName))
                {
                    // XML 루트 요소 이름을 시트 이름으로 사용
                    var doc = System.Xml.Linq.XDocument.Parse(xmlContent);
                    sheetName = doc.Root?.Name.LocalName ?? "Sheet1";
                }
                
                // 시트 이름 앞에 ! 추가 (Excel2Yaml 규칙)
                if (!sheetName.StartsWith("!"))
                {
                    sheetName = "!" + sheetName;
                }
                
                // 3단계: YAML을 Excel로 변환
                Logger.Information($"2단계: YAML → Excel 변환 (시트명: {sheetName})");
                
                // YAML 내용에서 루트 요소를 추출
                // XML을 YAML로 변환하면 최상위에 루트 요소 이름이 키로 포함되므로 이를 제거
                try
                {
                    var yaml = new YamlDotNet.RepresentationModel.YamlStream();
                    yaml.Load(new System.IO.StringReader(yamlContent));
                    
                    if (yaml.Documents.Count > 0 && yaml.Documents[0].RootNode is YamlDotNet.RepresentationModel.YamlMappingNode rootMapping)
                    {
                        // 루트가 단일 키를 가진 매핑이면, 그 값을 직접 사용
                        if (rootMapping.Children.Count == 1)
                        {
                            var rootKey = rootMapping.Children.Keys.First();
                            var rootValue = rootMapping.Children[rootKey];
                            
                            // 루트 값을 새로운 YAML 문서로 변환
                            var newYamlDoc = new YamlDotNet.RepresentationModel.YamlDocument(rootValue);
                            var newYamlStream = new YamlDotNet.RepresentationModel.YamlStream(newYamlDoc);
                            var writer = new System.IO.StringWriter();
                            newYamlStream.Save(writer, false);
                            yamlContent = writer.ToString();
                            
                            Logger.Debug($"루트 요소 '{rootKey}' 제거 후 YAML:\n{yamlContent}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"YAML 루트 요소 추출 중 오류 (무시하고 계속): {ex.Message}");
                }
                
                var workbook = _yamlToExcelConverter.ConvertToWorkbook(yamlContent, sheetName);
                
                Logger.Information("XML to Excel (via YAML) 변환 완료");
                return workbook;
            }
            catch (Exception ex)
            {
                Logger.Error($"XML to Excel 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// XML 파일을 Excel 파일로 변환하여 저장합니다.
        /// </summary>
        /// <param name="xmlPath">입력 XML 파일 경로</param>
        /// <param name="excelPath">출력 Excel 파일 경로</param>
        /// <param name="sheetName">생성할 시트 이름 (기본값: XML 루트 요소 이름)</param>
        public void ConvertFile(string xmlPath, string excelPath, string sheetName = null)
        {
            try
            {
                Logger.Information($"XML 파일 변환 시작: {xmlPath} → {excelPath}");
                
                // XML 파일 읽기
                var xmlContent = File.ReadAllText(xmlPath);
                
                // Excel로 변환
                var workbook = ConvertToExcel(xmlContent, sheetName);
                
                // Excel 파일 저장
                var directory = Path.GetDirectoryName(excelPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                workbook.SaveAs(excelPath);
                workbook.Dispose();
                
                Logger.Information($"변환 완료: {excelPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"파일 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// XML 내용을 중간 YAML로 변환합니다. (디버깅용)
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <returns>변환된 YAML 내용</returns>
        public string ConvertXmlToYaml(string xmlContent)
        {
            return _xmlToYamlConverter.ConvertToYaml(xmlContent);
        }

        /// <summary>
        /// XML을 Excel로 변환한 후 임시 경로에 저장하고 경로를 반환합니다.
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <param name="fileName">파일 이름 (확장자 제외)</param>
        /// <returns>저장된 Excel 파일 경로</returns>
        public string ConvertAndSaveToTemp(string xmlContent, string fileName = "converted")
        {
            try
            {
                var tempPath = Path.Combine(Path.GetTempPath(), $"{fileName}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                
                var workbook = ConvertToExcel(xmlContent);
                workbook.SaveAs(tempPath);
                workbook.Dispose();
                
                return tempPath;
            }
            catch (Exception ex)
            {
                Logger.Error($"임시 파일 저장 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }
    }
}