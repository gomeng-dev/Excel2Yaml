using ExcelToYamlAddin.Infrastructure.Configuration;
using ExcelToYamlAddin.Core;
using ExcelToYamlAddin.Application.PostProcessing;
using ExcelToYamlAddin.Forms;
using ExcelToYamlAddin.Infrastructure.Logging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Threading; // CancellationToken을 위해 추가
using Excel = Microsoft.Office.Interop.Excel;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Presentation.Helpers;

namespace ExcelToYamlAddin
{
    public partial class Ribbon : RibbonBase
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<Ribbon>();
        
        // 서비스 필드
        private readonly Presentation.Services.ConversionService _conversionService;
        private readonly Presentation.Services.ImportExportService _importExportService;
        private readonly Presentation.Services.PostProcessingService _postProcessingService;
        
        // 옵션 설정
        private bool enableHashGen = false;
        private bool addEmptyYamlFields = false;

        private readonly ExcelToYamlConfig config = new ExcelToYamlConfig();

        // 설정 폼 인스턴스 저장을 위한 필드
        private Forms.SheetPathSettingsForm settingsForm = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            LoadDefaultSettings();
            
            // 서비스 초기화
            _conversionService = new Presentation.Services.ConversionService();
            _importExportService = new Presentation.Services.ImportExportService();
            _postProcessingService = new Presentation.Services.PostProcessingService();
        }

        /// <summary>
        /// 기본 설정값을 로드합니다.
        /// </summary>
        private void LoadDefaultSettings()
        {
            try
            {
                // 초기값은 false로 설정(Ribbon_Load에서 시트별 설정을 로드)
                addEmptyYamlFields = false;

                // Properties.Settings.Default에서 기본 설정 로드
                addEmptyYamlFields = Properties.Settings.Default.AddEmptyYamlFields;
                Debug.WriteLine($"[Ribbon] 기본 설정값 로드: addEmptyYamlFields={addEmptyYamlFields}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Ribbon] 설정 로드 중 오류 발생: {ex.Message}");
            }
        }

        /// <summary>
        /// 현재 활성 시트의 설정과 전역 설정을 조합하여 최종 빈 필드 포함 설정을 반환합니다.
        /// </summary>
        /// <param name="configKey">Excel 설정에서 조회할 키 이름</param>
        /// <returns>시트별 설정 또는 전역 설정 중 하나라도 true이면 true, 그렇지 않으면 false</returns>
        private bool GetEffectiveEmptyFieldsSetting(string configKey)
        {
            return Presentation.Helpers.RibbonHelpers.GetEffectiveEmptyFieldsSetting(configKey, addEmptyYamlFields);
        }

        /// <summary>
        /// YAML 변환을 위한 공통 설정을 초기화합니다.
        /// </summary>
        /// <param name="methodName">호출한 메소드 이름 (로깅용)</param>
        /// <returns>초기화된 ExcelToYamlConfig 객체</returns>
        private ExcelToYamlConfig InitializeYamlConfig(string methodName)
        {
            var yamlConfig = new ExcelToYamlConfig
            {
                OutputFormat = OutputFormat.Yaml,
                IncludeEmptyFields = GetEffectiveEmptyFieldsSetting("YamlEmptyFields"),
                EnableHashGen = enableHashGen
            };

            SheetPathManager.Instance.Initialize();
            Debug.WriteLine($"[{methodName}] 변환 전 SheetPathManager 재초기화 완료");
            Debug.WriteLine($"[{methodName}] YAML 변환 설정 - 빈 필드 포함: {yamlConfig.IncludeEmptyFields}, 해시 생성: {yamlConfig.EnableHashGen}");

            return yamlConfig;
        }

        /// <summary>
        /// YAML 변환 결과를 담는 구조체
        /// </summary>
        private class YamlConversionResult
        {
            public List<string> YamlFiles { get; set; } = new List<string>();
            public string TempDirectory { get; set; } = string.Empty;
        }

        /// <summary>
        /// Excel을 YAML로 변환하는 공통 로직을 수행합니다.
        /// </summary>
        /// <param name="convertibleSheets">변환할 시트 목록</param>
        /// <param name="useTemporaryFiles">임시 파일을 사용할지 여부</param>
        /// <param name="methodName">호출한 메소드 이름 (로깅용)</param>
        /// <returns>변환된 YAML 파일 경로 목록과 임시 디렉토리 경로</returns>
        private YamlConversionResult ConvertToYaml(List<Excel.Worksheet> convertibleSheets, bool useTemporaryFiles, string methodName)
        {
            var config = InitializeYamlConfig(methodName);
            var result = new YamlConversionResult();
            
            if (useTemporaryFiles)
            {
                // 임시 디렉토리 생성
                string tempDir = Path.Combine(Path.GetTempPath(), $"Excel2YamlTemp_{Guid.NewGuid().ToString().Substring(0, 8)}");
                Directory.CreateDirectory(tempDir);
                config.WorkingDirectory = tempDir;
                
                Debug.WriteLine($"[{methodName}] 임시 디렉토리 생성: {tempDir}");
                result.TempDirectory = tempDir;
                
                // 임시 파일로 저장
                string tempFile = Globals.ThisAddIn.SaveToTempFile();
                result.YamlFiles = _conversionService.ConvertExcelFileToTemp(config, tempDir, convertibleSheets, tempFile);
            }
            else
            {
                // 일반 변환
                this.config.OutputFormat = config.OutputFormat;
                this.config.IncludeEmptyFields = config.IncludeEmptyFields;
                this.config.EnableHashGen = config.EnableHashGen;
                
                // 프로그레스 표시 없이 변환 (기존 로직에서는 ConvertExcelFile이 자체적으로 프로그레스를 표시함)
                result.YamlFiles = ConvertExcelFile(this.config);
            }
            
            return result;
        }

        // 리본 로드 시 호출
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                Debug.WriteLine("리본 로드 시작");

                // SheetPathManager 인스턴스 초기화
                var pathManager = ExcelToYamlAddin.Infrastructure.Configuration.SheetPathManager.Instance;
                if (pathManager == null)
                {
                    Debug.WriteLine("[Ribbon_Load] SheetPathManager 인스턴스를 가져올 수 없습니다.");
                }

                // ExcelConfigManager 인스턴스 초기화
                var excelConfigManager = ExcelConfigManager.Instance;
                if (excelConfigManager == null)
                {
                    Debug.WriteLine("[Ribbon_Load] ExcelConfigManager 인스턴스를 가져올 수 없습니다.");
                }

                Debug.WriteLine("리본 UI가 로드되었습니다.");

                // 디버그 모드에서만 HTML 내보내기 버튼 표시
#if DEBUG
                btnExportToHtml.Visible = true;
                Debug.WriteLine("디버그 모드: HTML 내보내기 버튼 표시");
#else
                btnExportToHtml.Visible = false;
                Debug.WriteLine("릴리즈 모드: HTML 내보내기 버튼 숨김");
#endif

                // 활성 시트의 설정 로드
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app != null && app.ActiveSheet != null && excelConfigManager != null)
                    {
                        Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
                        if (activeSheet != null)
                        {
                            string sheetName = activeSheet.Name;
                            addEmptyYamlFields = excelConfigManager.GetConfigBool(sheetName, "YamlEmptyFields", addEmptyYamlFields);
                            Debug.WriteLine($"[Ribbon_Load] 시트 '{sheetName}'에서 YAML 빈 배열 필드 설정 로드: {addEmptyYamlFields}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[Ribbon_Load] 활성 시트 설정 로드 중 오류: {ex.Message}");
                }

                // 설정 로드 확인
                if (pathManager != null)
                {
                    pathManager.Initialize(); // 설정 다시 로드

                    // 현재 워크북 설정
                    if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
                    {
                        string workbookPath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
                        pathManager.SetCurrentWorkbook(workbookPath);
                        Debug.WriteLine($"[Ribbon_Load] 현재 워크북 설정: {workbookPath}");

                        // 워크북의 모든 시트에 대한 설정 확인
                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                        {
                            bool yamlOption = pathManager.GetYamlEmptyFieldsOption(sheet.Name);
                            Debug.WriteLine($"[Ribbon_Load] 시트 '{sheet.Name}' YAML 설정: {yamlOption}");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("[Ribbon_Load] 활성화된 워크북이 없습니다.");
                    }
                }
                else
                {
                    Debug.WriteLine("[Ribbon_Load] SheetPathManager 인스턴스를 가져올 수 없습니다.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"리본 로드 중 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 변환을 위한 시트 준비 및 유효성 검사를 수행합니다.
        /// 공통 초기화, 시트 분석, 사용자 확인 등의 로직을 포함합니다.
        /// </summary>
        /// <param name="outConvertibleSheets">변환 가능한 시트 목록입니다.</param>
        /// <returns>변환을 계속 진행할 수 있으면 true, 그렇지 않으면 false를 반환합니다.</returns>
        private bool PrepareAndValidateSheets(out List<Excel.Worksheet> outConvertibleSheets)
        {
            // RibbonHelpers 사용하되, 설정 폼 관련 부분만 별도 처리
            bool result = Presentation.Helpers.RibbonHelpers.PrepareAndValidateSheets(out outConvertibleSheets);
            
            // 설정 폼이 열린 경우 추가 처리
            if (!result && outConvertibleSheets != null)
            {
                // RibbonHelpers에서 처리하지 못한 settingsForm 관련 로직
                if (settingsForm != null && !settingsForm.IsDisposed) 
                { 
                    settingsForm.Activate(); 
                    return false; 
                }
                
                // 워크시트 이름들을 미리 추출 (COM 객체 무효화 방지)
                var sheetNames = new List<string>();
                foreach (var sheet in outConvertibleSheets)
                {
                    try
                    {
                        sheetNames.Add(sheet.Name);
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Debug.WriteLine($"워크시트 이름 추출 실패 (COM Exception: {ex.ErrorCode:X})");
                    }
                }
                
                settingsForm = new Forms.SheetPathSettingsForm(outConvertibleSheets);
                settingsForm.FormClosed += (s, args) => { HandleSettingsFormClosedSafe(sheetNames); settingsForm = null; };
                settingsForm.StartPosition = FormStartPosition.CenterScreen; 
                settingsForm.Show();
            }
            
            return result;
        }

        /// <summary>
        /// 지정된 YAML 파일 목록에 대해 공통 후처리 작업을 적용합니다.
        /// </summary>
        /// <param name="yamlFilePaths">후처리를 적용할 YAML 파일 경로 목록입니다.</param>
        /// <param name="convertibleSheets">변환 가능한 원본 Excel 시트 목록입니다. 파일 이름과 시트 설정을 매칭하는 데 사용됩니다.</param>
        /// <param name="progress">진행 상태를 보고할 IProgress 객체입니다.</param>
        /// <param name="cancellationToken">작업 취소를 위한 CancellationToken입니다.</param>
        /// <param name="initialProgressPercentage">이 후처리 단계가 시작될 때의 전체 진행률입니다.</param>
        /// <param name="progressRange">이 후처리 단계가 전체 진행률에서 차지하는 범위입니다.</param>
        /// <param name="isForJsonConversion">JSON으로 변환하기 위한 중간 단계의 YAML 파일에 대한 후처리인 경우 true로 설정합니다. 이 경우 일부 후처리 단계가 생략될 수 있습니다.</param>
        /// <returns>키 경로 병합 및 Flow 스타일 처리 성공 횟수를 포함하는 튜플을 반환합니다.</returns>
        private (int mergeKeyPathsSuccessCount, int flowStyleSuccessCount) ApplyYamlPostProcessing(
            List<string> yamlFilePaths,
            List<Excel.Worksheet> convertibleSheets,
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            CancellationToken cancellationToken,
            int initialProgressPercentage,
            int progressRange,
            bool isForJsonConversion = false)
        {
            // PostProcessingService를 사용하여 후처리 적용
            return _postProcessingService.ApplyYamlPostProcessing(
                yamlFilePaths,
                convertibleSheets,
                progress,
                cancellationToken,
                initialProgressPercentage,
                progressRange,
                isForJsonConversion,
                addEmptyYamlFields);
        }

        // YAML으로 변환 버튼 클릭
        public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 초기화 및 다시 로드
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // YAML 변환 처리
                var yamlResult = ConvertToYaml(convertibleSheets, useTemporaryFiles: false, "OnConvertToYamlClick");
                List<string> convertedFiles = yamlResult.YamlFiles;

                // 변환 결과 추적
                int successCount = 0;
                int currentPostProcessMergeSuccessCount = 0; // 스코프 문제 해결을 위해 이름 변경 및 외부 선언
                int currentPostProcessFlowSuccessCount = 0;  // 스코프 문제 해결을 위해 이름 변경 및 외부 선언

                // YAML 후처리 기능 적용
                if (convertedFiles != null && convertedFiles.Count > 0)
                {
                    try
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 확인: {convertedFiles.Count}개 파일");
                        successCount = convertedFiles.Count;


                        // 후처리를 위한 프로그레스 바 표시
                        using (var progressForm = new Forms.ProgressForm())
                        {
                            progressForm.RunOperation((progress, cancellationToken) =>
                            {
                                try
                                {
                                    (currentPostProcessMergeSuccessCount, currentPostProcessFlowSuccessCount) = ApplyYamlPostProcessing(
                                        convertedFiles, convertibleSheets, progress, cancellationToken,
                                        initialProgressPercentage: 0, progressRange: 100, isForJsonConversion: false);

                                    // 모든 작업 완료
                                    progress.Report(new Forms.ProgressForm.ProgressInfo
                                    {
                                        Percentage = 100,
                                        StatusMessage = "모든 파일 처리 완료",
                                        IsCompleted = true
                                    });
                                }
                                catch (OperationCanceledException)
                                {
                                    // 취소 처리
                                    progress.Report(new Forms.ProgressForm.ProgressInfo
                                    {
                                        Percentage = 100,
                                        StatusMessage = "후처리가 취소되었습니다.",
                                        IsCompleted = true
                                    });
                                }
                            }, "YAML 후처리 중...");

                            progressForm.ShowDialog();

                            // 취소된 경우
                            if (progressForm.DialogResult == DialogResult.Cancel)
                            {
                                MessageBox.Show("후처리 작업이 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 중 오류 발생: {ex.Message}");
                    }

                    // 모든 후처리 작업이 완료된 후 메시지 표시
                    string message = $"{successCount}개의 시트가 성공적으로 변환되었습니다.";

                    if (currentPostProcessMergeSuccessCount > 0)
                        message += $"\n키 경로 병합 처리: {currentPostProcessMergeSuccessCount}개 파일";

                    if (currentPostProcessFlowSuccessCount > 0)
                        message += $"\nFlow 스타일 처리: {currentPostProcessFlowSuccessCount}개 파일";

                    if (convertedFiles.Count > 0)
                    {
                        message += "\n\n변환된 파일:";
                        foreach (var file in convertedFiles.Take(5))  // 첫 5개만 표시
                        {
                            message += $"\n{file}";
                        }

                        if (convertedFiles.Count > 5)
                        {
                            message += $"\n... 외 {convertedFiles.Count - 5}개 파일";
                        }
                    }

                    MessageBox.Show(message, "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("변환된 시트가 없습니다. 시트별 저장 경로를 설정했는지 확인하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML 변환 중 오류가 발생했습니다: {ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[Ribbon] YAML 변환 오류: {ex.Message}");
                Debug.WriteLine($"[Ribbon] 스택 트레이스: {ex.StackTrace}");
            }
        }

        // XML로 변환 버튼 클릭 (새로 추가된 메소드)
        public void OnConvertToXmlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // XML 변환 설정
                // 1단계: YAML로 먼저 변환하므로 OutputFormat.Yaml 설정
                config.OutputFormat = OutputFormat.Yaml;

                // 현재 활성 시트 또는 전역 설정에 따라 빈 필드 포함 여부 결정
                // 현재 시트의 빈 필드 포함 설정 결정 (YAML 단계용)
                config.IncludeEmptyFields = GetEffectiveEmptyFieldsSetting("YamlEmptyFields");
                Debug.WriteLine($"[OnConvertToXmlClick] YAML 변환 시 빈 필드 포함 설정: {config.IncludeEmptyFields}");

                config.EnableHashGen = enableHashGen; // 해시 생성 옵션 (YAML 단계용)

                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToXmlClick] 변환 전 SheetPathManager 재초기화 완료");

                // 프로그레스 바 표시 및 변환 작업 수행
                using (var progressForm = new Forms.ProgressForm())
                {
                    List<string> finalXmlFiles = new List<string>();
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 0, "Excel → YAML 변환 준비 중...");

                        // 1단계: Excel을 YAML로 변환 (임시 파일에 저장)
                        var yamlResult = ConvertToYaml(convertibleSheets, useTemporaryFiles: true, "OnConvertToXmlClick");
                        List<string> tempYamlFiles = yamlResult.YamlFiles;
                        string tempDir = yamlResult.TempDirectory;

                        try
                        {
                            Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 10, "Excel → YAML 파일 생성 중...");

                            if (tempYamlFiles.Count == 0)
                            {
                                Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 100, "변환할 YAML 파일이 없습니다.", isCompleted: true);
                                return;
                            }

                            // 2단계: YAML 후처리 (MergeKeyPaths, FlowStyle 등)
                            Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 30, "YAML 후처리 진행 중...");
                            ApplyYamlPostProcessing(tempYamlFiles, convertibleSheets, progress, cancellationToken, 30, 30, isForJsonConversion: false); // isForJsonConversion = false로 모든 후처리 적용

                            // 3단계: 후처리된 YAML을 XML로 변환
                            Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 60, "YAML → XML 변환 중...");

                            var yamlParser = new YamlDotNet.Serialization.DeserializerBuilder().Build();
                            int processedXmlCount = 0;

                            foreach (var yamlFilePath in tempYamlFiles)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                string sheetFileName = Path.GetFileNameWithoutExtension(yamlFilePath); // 예: "Sheet1"
                                string originalSheetName = convertibleSheets.FirstOrDefault(s => 
                                    (s.Name.StartsWith("!") ? s.Name.Substring(1) : s.Name).Equals(sheetFileName, StringComparison.OrdinalIgnoreCase))?.Name ?? sheetFileName;

                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 60 + (int)((double)processedXmlCount / tempYamlFiles.Count * 35),
                                    StatusMessage = $"'{sheetFileName}' YAML → XML 변환 중..."
                                });

                                string yamlContent = File.ReadAllText(yamlFilePath);
                                
                                // YAML이 배열인지 객체인지 동적으로 판단하여 역직렬화
                                object deserializedYaml = yamlParser.Deserialize<object>(yamlContent);
                                
                                IDictionary<string, object> dataForXml;
                                string xmlRootElementName = sheetFileName; // 기본값

                                if (deserializedYaml is IDictionary<string, object> yamlDict)
                                {
                                    // YAML이 객체인 경우
                                    if (yamlDict.Count == 1 && yamlDict.Values.First() is IDictionary<string, object> innerData)
                                    {
                                        xmlRootElementName = yamlDict.Keys.First();
                                        dataForXml = innerData;
                                    }
                                    else
                                    {
                                        dataForXml = yamlDict;
                                    }
                                }
                                else if (deserializedYaml is IList<object> yamlList)
                                {
                                    // YAML이 배열인 경우 (루트 요소 제거 후 배열이 된 경우)
                                    dataForXml = new Dictionary<string, object>
                                    {
                                        { "Items", yamlList }  // 배열을 Items로 감싸서 XML 구조 생성
                                    };
                                    xmlRootElementName = sheetFileName;
                                }
                                else
                                {
                                    // 기타 경우 (스칼라 값 등)
                                    dataForXml = new Dictionary<string, object>
                                    {
                                        { "Value", deserializedYaml }
                                    };
                                    xmlRootElementName = sheetFileName;
                                }
                                
                                string xmlString = Application.PostProcessing.YamlToXmlConverter.ConvertToXmlString(dataForXml, xmlRootElementName);

                                string savePath = SheetPathManager.Instance.GetSheetPath(originalSheetName);
                                if (string.IsNullOrEmpty(savePath)) savePath = Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName); // 기본 경로

                                string xmlFilePath = Path.Combine(savePath, $"{sheetFileName}.xml");
                                Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath)); // 경로 생성
                                File.WriteAllText(xmlFilePath, xmlString);
                                finalXmlFiles.Add(xmlFilePath);
                                processedXmlCount++;
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = $"{finalXmlFiles.Count}개 파일 XML 변환 완료", IsCompleted = true });
                        }
                        catch (OperationCanceledException)
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = "XML 변환 작업이 취소되었습니다.", IsCompleted = true });
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OnConvertToXmlClick] XML 변환 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = $"오류 발생: {ex.Message}", IsCompleted = true, HasError = true, ErrorMessage = ex.Message });
                        }
                        finally
                        {
                            Presentation.Helpers.RibbonHelpers.CleanupTempDirectory(tempDir, "OnConvertToXmlClick");
                        }
                    }, "Excel → YAML → XML 변환 중...");

                    progressForm.ShowDialog();

                    if (progressForm.DialogResult != DialogResult.Cancel && finalXmlFiles.Count > 0)
                    {
                        MessageBox.Show($"{finalXmlFiles.Count}개의 시트가 성공적으로 XML로 변환되었습니다.", "XML 변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (progressForm.DialogResult != DialogResult.Cancel)
                    {
                        MessageBox.Show("변환된 XML 파일이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 변환 중 오류가 발생했습니다: {ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[Ribbon] XML 변환 오류: {ex.Message}");
                Debug.WriteLine($"[Ribbon] 스택 트레이스: {ex.StackTrace}");
            }
        }


        // YAML을 JSON으로 변환 버튼 클릭
        public void OnConvertYamlToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // 프로그레스 바 표시 및 변환 작업 수행
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        progress.Report(new Forms.ProgressForm.ProgressInfo
                        {
                            Percentage = 0,
                            StatusMessage = "YAML 파일 변환 준비 중..."
                        });

                        // 1단계: YAML로 변환 (임시 파일에 저장)
                        var yamlResult = ConvertToYaml(convertibleSheets, useTemporaryFiles: true, "OnConvertYamlToJsonClick");
                        List<string> yamlFiles = yamlResult.YamlFiles;
                        string tempDir = yamlResult.TempDirectory;
                        List<Tuple<string, string>> convertPairs = new List<Tuple<string, string>>();

                        try
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 10,
                                StatusMessage = "YAML 파일 생성 완료..."
                            });

                            if (yamlFiles.Count == 0)
                            {
                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 100,
                                    StatusMessage = "변환할 YAML 파일이 없습니다.",
                                    IsCompleted = true
                                });
                                return;
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 30,
                                StatusMessage = "YAML 후처리 진행 중..."
                            });
                            ApplyYamlPostProcessing(yamlFiles, convertibleSheets, progress, cancellationToken, 30, 20, isForJsonConversion: true);

                            // 3단계: JSON 변환 준비
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 50,
                                StatusMessage = $"{yamlFiles.Count}개 YAML 파일 JSON 변환 준비 중..."
                            });

                            // YAML 파일을 JSON으로 변환할 쌍 생성
                            foreach (var yamlFile in yamlFiles)
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                string fileName = Path.GetFileNameWithoutExtension(yamlFile);
                                string sheetName = fileName.StartsWith("!") ? fileName : "!" + fileName;

                                // 실제 저장 경로 가져오기
                                string savePath = SheetPathManager.Instance.GetSheetPath(sheetName);
                                if (string.IsNullOrEmpty(savePath))
                                {
                                    // !가 없는 이름으로 시도
                                    savePath = SheetPathManager.Instance.GetSheetPath(fileName);
                                }

                                if (!string.IsNullOrEmpty(savePath))
                                {
                                    string jsonFilePath = Path.Combine(savePath, fileName + ".json");
                                    convertPairs.Add(new Tuple<string, string>(yamlFile, jsonFilePath));
                                }
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 70,
                                StatusMessage = "YAML에서 JSON으로 변환 중..."
                            });

                            // YAML을 JSON으로 변환
                            List<string> convertedJsonFiles = Application.PostProcessing.YamlToJsonProcessor.BatchConvertYamlToJson(convertPairs);

                            // 결과 보고
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 100,
                                StatusMessage = $"{convertedJsonFiles.Count}개 파일 변환 완료",
                                IsCompleted = true
                            });
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OnConvertYamlToJsonClick] 변환 중 오류 발생: {ex.Message}");
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 100,
                                StatusMessage = $"오류 발생: {ex.Message}",
                                IsCompleted = true
                            });
                        }
                        finally
                        {
                            // 임시 디렉토리 정리
                            Presentation.Helpers.RibbonHelpers.CleanupTempDirectory(tempDir, "OnConvertYamlToJsonClick");
                        }
                    }, "YAML → JSON 변환 중...");

                    progressForm.ShowDialog();

                    // 변환 완료 메시지
                    if (progressForm.DialogResult != DialogResult.Cancel)
                    {
                        MessageBox.Show("YAML → JSON 변환이 완료되었습니다.", "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML → JSON 변환 중 오류가 발생했습니다: {ex.Message}",
                    "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[OnConvertYamlToJsonClick] 오류: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // 도움말 버튼 클릭 이벤트 핸들러
        private void OnHelpButtonClick(object sender, RibbonControlEventArgs e)
        {
            OnHelpClick(null);
        }

        public void OnHelpClick(IRibbonControl control)
        {
            try
            {
                // 임베디드 리소스에서 HTML 내용 로드
                string htmlContent = null;
                using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("ExcelToYamlAddin.Docs.Readme.html"))
                {
                    if (stream == null)
                    {
                        // 임베디드 리소스를 찾을 수 없는 경우, 물리적 파일을 시도해봅니다.
                        string addinPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        string readmePath = Path.Combine(addinPath, "Docs", "Readme.html");

                        if (File.Exists(readmePath))
                        {
                            // 물리적 파일이 존재하면 내용을 읽어옵니다.
                            htmlContent = File.ReadAllText(readmePath, Encoding.UTF8);
                        }
                        else
                        {
                            MessageBox.Show("도움말 리소스를 찾을 수 없습니다.",
                                           "리소스 없음",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Warning);
                            Debug.WriteLine("[OnHelpClick] 도움말 리소스를 찾을 수 없습니다.");
                            return;
                        }
                    }
                    else
                    {
                        // 임베디드 리소스가 있으면 내용을 읽어옵니다.
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            htmlContent = reader.ReadToEnd();
                        }
                    }
                }

                // 도움말 폼 생성 (비모달)
                Form helpForm = new Form()
                {
                    Text = "Excel2YAML 사용 설명서",
                    Size = new Size(1000, 700),
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.SizableToolWindow,
                    ShowInTaskbar = false,
                    TopMost = true
                };

                var browser = new WebBrowser();
                browser.Dock = DockStyle.Fill;
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = htmlContent;

                helpForm.Controls.Add(browser);
                helpForm.Show(); // 비모달로 표시

                Debug.WriteLine("[OnHelpClick] 도움말 내용을 플로팅 윈도우에 표시했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"도움말 파일을 열 수 없습니다: {ex.Message}",
                               "오류",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
                Debug.WriteLine($"[OnHelpClick] 오류: {ex.Message}");
            }
        }

        // 시트별 경로 설정 버튼 클릭
        public void OnSheetPathSettingsClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;

                if (app.ActiveWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 워크북 경로 설정
                string workbookPath = app.ActiveWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(app.ActiveWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 시트별 경로 설정 폼 열기
                using (var form = new SheetPathSettingsForm(convertibleSheets))
                {
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"시트별 경로 설정 중 오류가 발생했습니다: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Debug.WriteLine($"시트별 경로 설정 오류: {ex}");
            }
        }

        /// <summary>
        /// SheetPathSettingsForm이 닫힐 때 호출되는 안전한 핸들러입니다.
        /// </summary>
        /// <param name="sheetNames">시트 이름 목록입니다.</param>
        private void HandleSettingsFormClosedSafe(List<string> sheetNames)
        {
            // 설정 후 다시 활성화된 시트 수 확인
            int enabledSheetsCount = 0;
            foreach (var sheetName in sheetNames)
            {
                if (SheetPathManager.Instance.IsSheetEnabled(sheetName))
                {
                    enabledSheetsCount++;
                }
            }

            // 활성화된 시트가 없으면 변환 취소 메시지 (선택적)
            if (enabledSheetsCount == 0) { /* MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.", "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information); */ }
        }


        // Excel 파일 변환 처리 (수정: 변환된 파일 목록 반환)
        private List<string> ConvertExcelFile(ExcelToYamlConfig config)
        {
            string tempFile = null;
            List<string> convertedFiles = new List<string>();

            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return convertedFiles;
                }

                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // 디버깅을 위한 로그 추가
                Debug.WriteLine($"현재 워크북 경로: {workbookPath}");

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                Debug.WriteLine($"변환 가능한 시트 수: {convertibleSheets.Count}");
                foreach (var sheet in convertibleSheets)
                {
                    Debug.WriteLine($"시트 이름: {sheet.Name}");
                }

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return convertedFiles;
                }

                // 임시 파일로 저장
                tempFile = addIn.SaveToTempFile();
                if (string.IsNullOrEmpty(tempFile))
                {
                    MessageBox.Show("임시 파일을 생성할 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return convertedFiles;
                }

                // 프로그레스 바 표시 및 변환 작업 수행
                string outputFormat = config.OutputFormat == OutputFormat.Json ? "JSON" : "YAML";
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        convertedFiles = _conversionService.ConvertExcelFile(
                            config,
                            activeWorkbook,
                            tempFile,
                            convertibleSheets,
                            progress,
                            cancellationToken);
                    }, $"Excel → {outputFormat} 변환 중");

                    progressForm.ShowDialog();

                    // 취소된 경우
                    if (progressForm.DialogResult == DialogResult.Cancel)
                    {
                        MessageBox.Show("변환 작업이 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                return convertedFiles;
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
                        Debug.WriteLine($"[ConvertExcelFile] 임시 파일 삭제 중 오류 발생: {ex.Message}");
                    }
                }
            }

            // 예외 발생 시 이 지점에 도달하므로 빈 목록 반환
            return convertedFiles;
        }

        // HTML 내보내기 버튼 클릭
        public void OnExportToHtmlClick(object sender, RibbonControlEventArgs e)
        {
            ExportCurrentSheetToHtml();
        }

        // 디버깅: 현재 시트를 HTML로 내보내기
        public void ExportCurrentSheetToHtml()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeSheet = app.ActiveSheet as Excel.Worksheet;
                
                if (activeSheet == null)
                {
                    MessageBox.Show("활성 시트가 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 저장 경로 선택
                var saveDialog = new SaveFileDialog
                {
                    Title = "HTML 파일 저장",
                    Filter = "HTML 파일 (*.html)|*.html",
                    FileName = $"{activeSheet.Name}_debug.html",
                    RestoreDirectory = true
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelToHtmlExporter.ExportToHtml(activeSheet, saveDialog.FileName);
                    
                    // 브라우저에서 열기
                    var result = MessageBox.Show($"HTML 파일이 생성되었습니다.\n\n{saveDialog.FileName}\n\n브라우저에서 열어보시겠습니까?",
                        "내보내기 완료", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    
                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"HTML 내보내기 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // XML 가져오기 버튼 클릭
        public void OnImportXmlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 파일 선택
                string xmlFilePath = _importExportService.ShowImportFileDialog("XML", "*.xml");
                if (xmlFilePath == null) return;

                // 워크북 유효성 검사
                var currentWorkbook = Presentation.Helpers.RibbonHelpers.ValidateWorkbookForImport();
                if (currentWorkbook == null) return;

                string fileName = Path.GetFileNameWithoutExtension(xmlFilePath);
                bool success = false;

                // 프로그레스 바와 함께 Import 실행
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 10, "XML 파일 분석 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 30, "XML을 YAML로 변환 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 60, "YAML을 Excel로 변환 중...");
                        success = _importExportService.ImportXmlToExcel(xmlFilePath, currentWorkbook);

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 90, "시트 정리 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 100, "XML 가져오기 완료", isCompleted: true);
                    }, $"XML 파일 '{fileName}' 가져오는 중...");

                    progressForm.ShowDialog();

                    // 취소된 경우
                    if (progressForm.DialogResult == DialogResult.Cancel)
                    {
                        MessageBox.Show("XML 가져오기가 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                if (success)
                {
                    string newSheetName = fileName; // ImportXmlToExcel에서 이미 고유한 이름을 생성함
                    
                    MessageBox.Show($"XML 파일이 성공적으로 Excel로 변환되었습니다.\n\n파일: {fileName}.xml\n시트: {newSheetName}",
                        "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 변환 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // YAML 가져오기 버튼 클릭
        public void OnImportYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 파일 선택
                string yamlFilePath = _importExportService.ShowImportFileDialog("YAML", "*.yaml;*.yml");
                if (yamlFilePath == null) return;

                // 워크북 유효성 검사
                var currentWorkbook = Presentation.Helpers.RibbonHelpers.ValidateWorkbookForImport();
                if (currentWorkbook == null) return;

                string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);
                bool success = false;

                // 프로그레스 바와 함께 Import 실행
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 10, "YAML 파일 읽기 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 30, "YAML 구조 분석 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 60, "YAML을 Excel로 변환 중...");
                        success = _importExportService.ImportYamlToExcel(yamlFilePath, currentWorkbook);

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 90, "시트 추가 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 100, "YAML 가져오기 완료", isCompleted: true);
                    }, $"YAML 파일 '{fileName}' 가져오는 중...");

                    progressForm.ShowDialog();

                    // 취소된 경우
                    if (progressForm.DialogResult == DialogResult.Cancel)
                    {
                        MessageBox.Show("YAML 가져오기가 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                
                if (success)
                {
                    string newSheetName = "!" + fileName; // ImportYamlToExcel에서 이미 고유한 이름을 생성함
                    
                    MessageBox.Show($"YAML 파일이 성공적으로 가져와졌습니다.\n\n시트 이름: {newSheetName}",
                        "가져오기 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML 변환 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Import 관련 헬퍼 메서드들은 ImportExportService로 이동됨



        // 엑셀 파일을 임시 디렉토리에 YAML로 변환하는 메서드 (삭제 - ConversionService로 이동)

        // JSON 가져오기 버튼 클릭
        private void OnImportJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Logger.Information("JSON 가져오기 버튼 클릭");

                // 파일 선택
                string jsonPath = _importExportService.ShowImportFileDialog("JSON", "*.json");
                if (jsonPath == null) return;

                Logger.Information($"선택된 JSON 파일: {jsonPath}");

                // 워크북 유효성 검사
                var currentWorkbook = Presentation.Helpers.RibbonHelpers.ValidateWorkbookForImport();
                if (currentWorkbook == null) return;

                string fileName = Path.GetFileNameWithoutExtension(jsonPath);
                bool success = false;

                // 프로그레스 바와 함께 Import 실행
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 10, "JSON 파일 읽기 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 30, "JSON을 YAML로 변환 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 60, "YAML을 Excel로 변환 중...");
                        success = _importExportService.ImportJsonToExcel(jsonPath, currentWorkbook);

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 90, "Excel 파일 저장 중...");
                        cancellationToken.ThrowIfCancellationRequested();

                        Presentation.Helpers.RibbonHelpers.ReportProgress(progress, 100, "JSON 가져오기 완료", isCompleted: true);
                    }, $"JSON 파일 '{fileName}' 가져오는 중...");

                    progressForm.ShowDialog();

                    // 취소된 경우
                    if (progressForm.DialogResult == DialogResult.Cancel)
                    {
                        MessageBox.Show("JSON 가져오기가 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                if (success)
                {
                    string excelPath = Path.Combine(Path.GetDirectoryName(jsonPath), $"{fileName}.xlsx");

                    Logger.Information($"JSON → Excel 변환 완료: {excelPath}");

                    // 결과 알림
                    var result = MessageBox.Show(
                        $"JSON 파일을 Excel로 변환했습니다.\n\n생성된 파일: {excelPath}\n\n파일을 열어보시겠습니까?",
                        "변환 완료",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(excelPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"JSON 가져오기 중 오류 발생: {ex.Message}", ex);
                MessageBox.Show($"JSON 가져오기 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}