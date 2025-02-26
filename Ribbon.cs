using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using ExcelToJsonAddin.Config;
using System.Diagnostics;
using Microsoft.Office.Core;
using System.Reflection;

namespace ExcelToJsonAddin
{
    public partial class Ribbon : RibbonBase
    {
        // 옵션 설정
        private bool includeEmptyFields = false;
        private bool enableHashGen = false;
        
        private readonly ExcelToJsonConfig config = new ExcelToJsonConfig();

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        // 리본 로드 시 호출
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                Debug.WriteLine("리본 UI가 로드되었습니다.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"리본 로드 중 오류: {ex.Message}");
            }
        }
        
        // JSON으로 변환 버튼 클릭
        public void OnConvertToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 적용
                config.IncludeEmptyFields = includeEmptyFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Json;
                
                // 변환 처리
                ConvertExcelFile(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"JSON 변환 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"JSON 변환 오류: {ex}");
            }
        }
        
        // YAML으로 변환 버튼 클릭
        public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 적용
                config.IncludeEmptyFields = includeEmptyFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Yaml;
                
                // 변환 처리
                ConvertExcelFile(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML 변환 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"YAML 변환 오류: {ex}");
            }
        }

        // 빈 필드 포함 옵션 체크박스 상태 가져오기
        public bool GetEmptyFieldsState(IRibbonControl control)
        {
            return includeEmptyFields;
        }

        // 빈 필드 포함 옵션 체크박스 클릭
        public void OnEmptyFieldsClicked(IRibbonControl control, bool pressed)
        {
            includeEmptyFields = pressed;
        }

        // MD5 해시 생성 옵션 체크박스 상태 가져오기
        public bool GetHashGenState(IRibbonControl control)
        {
            return enableHashGen;
        }

        // MD5 해시 생성 옵션 체크박스 클릭
        public void OnHashGenClicked(IRibbonControl control, bool pressed)
        {
            enableHashGen = pressed;
        }

        // 고급 설정 버튼 클릭
        public void OnSettingsClick(IRibbonControl control)
        {
            MessageBox.Show("고급 설정 기능은 개발 중입니다.", "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Excel 파일 변환 처리
        private void ConvertExcelFile(ExcelToJsonConfig config)
        {
            string tempFile = null;
            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                
                // 임시 파일로 저장
                tempFile = addIn.SaveToTempFile();
                if (string.IsNullOrEmpty(tempFile))
                {
                    MessageBox.Show("임시 파일을 생성할 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 저장 경로 선택
                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.Filter = config.OutputFormat == OutputFormat.Json 
                        ? "JSON 파일 (*.json)|*.json"
                        : "YAML 파일 (*.yaml)|*.yaml|YML 파일 (*.yml)|*.yml";
                    dlg.DefaultExt = config.OutputFormat == OutputFormat.Json ? ".json" : ".yaml";
                    
                    if (dlg.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    // 변환 처리
                    var excelReader = new ExcelReader(config);
                    excelReader.ProcessExcelFile(tempFile, dlg.FileName);

                    // 결과 파일 경로
                    string resultFile = dlg.FileName;
                    
                    // 변환 완료 메시지
                    MessageBox.Show($"변환이 완료되었습니다.\n저장 위치: {resultFile}", 
                        "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    // 결과 파일이 있는 폴더 열기
                    if (MessageBox.Show("변환된 파일이 있는 폴더를 열까요?", 
                        "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Process.Start("explorer.exe", $"/select,\"{resultFile}\"");
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"파일 처리 중 오류가 발생했습니다: {ex.Message}", 
                    "파일 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"파일 접근 권한이 없습니다: {ex.Message}", 
                    "권한 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"변환 중 오류가 발생했습니다: {ex.Message}", 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 임시 파일 정리
                if (!string.IsNullOrEmpty(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"임시 파일 삭제 중 오류 발생: {ex.Message}");
                    }
                }
            }
        }

        // 리소스 텍스트 가져오기
        private static string GetResourceText(string resourceName)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            
            foreach (string name in assembly.GetManifestResourceNames())
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var stream = assembly.GetManifestResourceStream(name))
                    {
                        if (stream != null)
                        {
                            using (var reader = new StreamReader(stream))
                            {
                                return reader.ReadToEnd();
                            }
                        }
                    }
                }
            }
            
            return null;
        }
    }
} 