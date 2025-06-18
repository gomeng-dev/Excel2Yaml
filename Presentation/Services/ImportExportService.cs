using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using YamlDotNet.RepresentationModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToYamlAddin.Presentation.Services
{
    /// <summary>
    /// 파일 Import/Export 관련 기능을 처리하는 서비스
    /// </summary>
    public class ImportExportService
    {
        /// <summary>
        /// XML 파일을 Excel로 가져옵니다.
        /// </summary>
        public bool ImportXmlToExcel(string xmlFilePath, Excel.Workbook currentWorkbook)
        {
            try
            {
                string xmlContent = File.ReadAllText(xmlFilePath);
                string fileName = Path.GetFileNameWithoutExtension(xmlFilePath);

                // XML을 Excel로 변환 (XML → YAML → Excel)
                var converter = new Core.XmlToExcelViaYamlConverter();
                var workbook = converter.ConvertToExcel(xmlContent, fileName);

                // ClosedXML 워크북을 임시 파일로 저장
                string tempFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xlsx");
                workbook.SaveAs(tempFile);

                // 시트 이름 생성 (XML은 '!' 접두사 없이)
                string newSheetName = GenerateUniqueSheetName(currentWorkbook, fileName, addExclamation: false);

                // 임시 파일을 새 시트로 복사
                bool success = CopyTempFileToNewSheet(tempFile, newSheetName, currentWorkbook);
                
                // 임시 파일 정리
                CleanupTempFile(tempFile);

                if (success)
                {
                    // 새로 추가된 시트 활성화
                    var newSheet = currentWorkbook.Worksheets[newSheetName];
                    newSheet.Activate();
                }

                return success;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ImportXmlToExcel] XML 변환 중 오류: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// YAML 파일을 Excel로 가져옵니다.
        /// </summary>
        public bool ImportYamlToExcel(string yamlFilePath, Excel.Workbook currentWorkbook)
        {
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);

                // YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlFilePath);
                Debug.WriteLine($"[ImportYamlToExcel] YAML 파일 읽기 완료: {yamlFilePath}");

                // 루트 요소 처리 (XML에서 변환된 경우)
                yamlContent = ProcessYamlRootElement(yamlContent);
                
                // YAML을 Excel로 변환
                string tempFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xlsx");
                
                Debug.WriteLine($"[ImportYamlToExcel] YAML to Excel 변환 시작: {tempFile}");
                var converter = new Core.YamlToExcel.YamlToExcelConverter();
                converter.ConvertFromContent(yamlContent, tempFile);
                Debug.WriteLine($"[ImportYamlToExcel] YAML to Excel 변환 완료: {tempFile}");

                // 새 시트 이름 생성 및 복사
                string newSheetName = GenerateUniqueSheetName(currentWorkbook, fileName, addExclamation: true);
                bool success = CopyTempFileToNewSheet(tempFile, newSheetName, currentWorkbook);
                
                // 임시 파일 정리
                CleanupTempFile(tempFile);
                
                return success;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ImportYamlToExcel] YAML 변환 중 오류: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// JSON 파일을 Excel로 가져옵니다.
        /// </summary>
        public bool ImportJsonToExcel(string jsonPath, Excel.Workbook currentWorkbook)
        {
            try
            {
                // JSON을 YAML로 변환 후 Excel로 변환
                string jsonContent = File.ReadAllText(jsonPath);
                
                // JSON을 YAML로 변환 (Newtonsoft.Json 사용)
                var jsonObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);
                var serializer = new YamlDotNet.Serialization.SerializerBuilder().Build();
                string yamlContent = serializer.Serialize(jsonObject);

                // YAML을 Excel로 변환
                var yamlToExcel = new Core.YamlToExcel.YamlToExcelConverter();
                string fileName = Path.GetFileNameWithoutExtension(jsonPath);
                string excelPath = Path.Combine(Path.GetDirectoryName(jsonPath), $"{fileName}.xlsx");

                yamlToExcel.ConvertFromContent(yamlContent, excelPath);

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ImportJsonToExcel] JSON 변환 중 오류: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Import 작업을 위한 파일 선택 대화상자를 표시합니다.
        /// </summary>
        public string ShowImportFileDialog(string fileType, string fileExtensions)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = $"{fileType} 파일 선택";
                openFileDialog.Filter = $"{fileType} 파일 ({fileExtensions})|{fileExtensions}|모든 파일 (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                return openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : null;
            }
        }

        /// <summary>
        /// 중복되지 않는 새 시트 이름을 생성합니다.
        /// </summary>
        private string GenerateUniqueSheetName(dynamic workbook, string baseName, bool addExclamation = true)
        {
            string prefix = addExclamation ? "!" : "";
            string newSheetName = $"{prefix}{baseName}";
            int suffix = 1;
            
            while (WorksheetExists(workbook, newSheetName))
            {
                newSheetName = $"{prefix}{baseName}_{suffix++}";
            }
            
            return newSheetName;
        }

        /// <summary>
        /// 워크시트 존재 여부를 확인합니다.
        /// </summary>
        private bool WorksheetExists(dynamic workbook, string sheetName)
        {
            foreach (dynamic sheet in workbook.Worksheets)
            {
                if (sheet.Name == sheetName)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 임시 파일을 현재 워크북의 새 시트로 복사합니다.
        /// </summary>
        private bool CopyTempFileToNewSheet(string tempFilePath, string sheetName, Excel.Workbook currentWorkbook)
        {
            Excel.Application app = currentWorkbook.Application;
            
            try
            {
                Debug.WriteLine($"[CopyTempFileToNewSheet] 임시 파일 열기 시작: {tempFilePath}");
                var tempWorkbook = app.Workbooks.Open(tempFilePath);
                var sourceSheet = tempWorkbook.Worksheets[1];
                Debug.WriteLine($"[CopyTempFileToNewSheet] 임시 파일 열기 완료");

                Debug.WriteLine($"[CopyTempFileToNewSheet] 시트 복사 시작: {sheetName}");
                sourceSheet.Copy(After: currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);
                var newSheet = currentWorkbook.ActiveSheet;
                newSheet.Name = sheetName;
                Debug.WriteLine($"[CopyTempFileToNewSheet] 시트 복사 완료: {sheetName}");

                // 임시 워크북 닫기
                try
                {
                    Debug.WriteLine($"[CopyTempFileToNewSheet] 임시 워크북 닫기");
                    tempWorkbook.Close(false);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[CopyTempFileToNewSheet] 임시 워크북 닫기 중 오류 (무시): {ex.Message}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CopyTempFileToNewSheet] 시트 복사 실패: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 임시 파일을 안전하게 삭제합니다.
        /// </summary>
        private void CleanupTempFile(string tempFilePath)
        {
            try
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                    Debug.WriteLine($"[CleanupTempFile] 임시 파일 삭제 완료: {tempFilePath}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CleanupTempFile] 임시 파일 삭제 실패 (무시): {ex.Message}");
            }
        }

        /// <summary>
        /// YAML 루트 요소를 처리합니다 (XML에서 변환된 경우 단일 루트 키 제거).
        /// </summary>
        private string ProcessYamlRootElement(string yamlContent)
        {
            try
            {
                Debug.WriteLine($"[ProcessYamlRootElement] YAML 구조 분석 시작");
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));
                
                if (yaml.Documents.Count > 0 && yaml.Documents[0].RootNode is YamlMappingNode rootMapping)
                {
                    // 루트가 단일 키를 가진 매핑이면, 그 값을 직접 사용
                    if (rootMapping.Children.Count == 1)
                    {
                        var rootKey = rootMapping.Children.Keys.First();
                        var rootValue = rootMapping.Children[rootKey];
                        
                        Debug.WriteLine($"[ProcessYamlRootElement] 루트 요소 '{rootKey}' 감지, 제거 중...");
                        
                        // 루트 값을 새로운 YAML 문서로 변환
                        var newYamlDoc = new YamlDocument(rootValue);
                        var newYamlStream = new YamlStream(newYamlDoc);
                        var writer = new StringWriter();
                        newYamlStream.Save(writer, false);
                        yamlContent = writer.ToString();
                        
                        Debug.WriteLine($"[ProcessYamlRootElement] 루트 요소 제거 완료");
                    }
                }
                Debug.WriteLine($"[ProcessYamlRootElement] YAML 구조 분석 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessYamlRootElement] 루트 요소 처리 중 오류 (무시하고 계속): {ex.Message}");
            }
            
            return yamlContent;
        }
    }
}