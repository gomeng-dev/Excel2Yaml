using ExcelToYamlAddin.Logging;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelToYamlAddin
{
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                // 리본 UI XML 로드
                return GetResourceText("ExcelToYamlAddin.RibbonUI.xml");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"리본 UI 로드 오류: {ex.Message}");
                MessageBox.Show($"리본 UI 로드 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"></customUI>";
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            // 실제 리소스 이름 찾기 (대소문자나 경로 차이가 있을 수 있음)
            string actualResourceName = resourceNames.FirstOrDefault(rn => rn.EndsWith("RibbonUI.xml", StringComparison.OrdinalIgnoreCase));

            if (actualResourceName != null)
            {
                using (Stream stream = asm.GetManifestResourceStream(actualResourceName))
                {
                    if (stream != null)
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }

            Debug.WriteLine($"리소스를 찾을 수 없음: {resourceName}");
            Debug.WriteLine($"사용 가능한 리소스: {string.Join(", ", resourceNames)}");
            return null;
        }
    }

    public partial class ThisAddIn
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ThisAddIn>();

        // RibbonController 인스턴스를 COM에 등록
        protected override object RequestComAddInAutomationService()
        {
            return new RibbonController();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 애드인 시작 시 초기화
            try
            {
                // Add-in 초기화 로깅
                Logger.Debug("Excel To JSON Add-in 시작");

                // 레지스트리에 LoadBehavior 값을 3으로 강제 설정
                SetLoadBehaviorToAuto();

                // COM 추가 기능 활성화 상태 확인 및 설정 (스레드 없이 안전하게 실행)
                SafeEnsureComAddinEnabled();

                // 설치 후 첫 실행 여부 확인 (스레드 중단 문제 없이 안전하게 처리)
                try
                {
                    bool isFirstRun = IsFirstRun();

                    if (isFirstRun)
                    {
                        Debug.WriteLine("애드인 첫 실행 감지: 자동 활성화 설정을 적용합니다.");

                        // 첫 실행 플래그 설정
                        SetFirstRunFlag(false);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"첫 실행 설정 중 오류 (무시됨): {ex.Message}");
                    // 첫 실행 확인은 중요하지만 필수는 아니므로 오류가 발생해도 계속 진행
                }

                // SheetPathManager 초기화 및 설정 미리 로드
                ExcelToYamlAddin.Config.SheetPathManager.Instance.Initialize();

                // 현재 워크북 설정
                if (this.Application.ActiveWorkbook != null)
                {
                    string workbookPath = this.Application.ActiveWorkbook.FullName;
                    ExcelToYamlAddin.Config.SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
                    Logger.Information("현재 워크북 설정: {0}", workbookPath);
                }

                // Ribbon 인스턴스 생성 및 등록
                var ribbon = new Ribbon();
                Debug.WriteLine("Ribbon 인스턴스가 생성되었습니다.");
                Logger.Information("Excel 애드인 시작됨");
            }
            catch (System.Threading.ThreadAbortException)
            {
                // ThreadAbortException은 특별 처리 - 무시하고 계속 진행
                Debug.WriteLine("스레드 중단 예외 발생 (무시됨)");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"애드인 초기화 중 오류: {ex.Message}");
                try
                {
                    MessageBox.Show($"애드인 초기화 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                    // MessageBox 표시 중 오류가 발생해도 무시
                }
            }
        }

        // 첫 실행 여부 확인
        private bool IsFirstRun()
        {
            try
            {
                string keyPath = @"Software\Microsoft\Office\Excel\Addins\ExcelToYamlAddin";
                RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true);

                if (key != null)
                {
                    object value = key.GetValue("FirstRun");
                    if (value == null)
                    {
                        // 값이 없으면 첫 실행으로 간주
                        return true;
                    }

                    return Convert.ToBoolean(value);
                }

                // 키가 없으면 첫 실행으로 간주
                return true;
            }
            catch
            {
                // 오류 발생 시 기본값으로 첫 실행 아님
                return false;
            }
        }

        // 첫 실행 플래그 설정
        private void SetFirstRunFlag(bool isFirstRun)
        {
            try
            {
                string keyPath = @"Software\Microsoft\Office\Excel\Addins\ExcelToYamlAddin";
                RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true);

                if (key != null)
                {
                    key.SetValue("FirstRun", isFirstRun ? 1 : 0, RegistryValueKind.DWord);
                    key.Close();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"첫 실행 플래그 설정 중 오류: {ex.Message}");
            }
        }

        // 스레드 중단 없이 안전하게 COM 추가 기능 활성화 상태 확인 및 설정
        private void SafeEnsureComAddinEnabled()
        {
            try
            {
                // 현재 애드인의 ProgID 가져오기
                string progID = "ExcelToYamlAddin";

                // COM 추가 기능 컬렉션 가져오기
                Office.COMAddIns comAddIns = this.Application.COMAddIns;

                // 현재 애드인 찾기
                foreach (Office.COMAddIn addIn in comAddIns)
                {
                    if (addIn.ProgId.Contains(progID))
                    {
                        // 비활성화 상태면 활성화
                        if (!addIn.Connect)
                        {
                            Debug.WriteLine($"COM 추가 기능 '{addIn.ProgId}'가 비활성화 상태입니다. 활성화합니다.");
                            addIn.Connect = true;
                        }
                        else
                        {
                            Debug.WriteLine($"COM 추가 기능 '{addIn.ProgId}'가 이미 활성화되어 있습니다.");
                        }

                        // Sleep 호출 없이 바로 종료
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"COM 추가 기능 활성화 확인 중 오류 (무시됨): {ex.Message}");
                // 오류가 발생해도 계속 진행
            }
        }

        // 레지스트리에 LoadBehavior 값을 3으로 설정
        private void SetLoadBehaviorToAuto()
        {
            try
            {
                string keyPath = @"Software\Microsoft\Office\Excel\Addins\ExcelToYamlAddin";
                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath, true);

                if (key != null)
                {
                    key.SetValue("LoadBehavior", 3, Microsoft.Win32.RegistryValueKind.DWord);
                    key.Close();
                    Debug.WriteLine("로드 동작을 '자동'으로 설정했습니다.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"레지스트리 설정 오류: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 애드인 종료 시 정리 코드
            Logger.Information("Excel 애드인 종료됨");
        }

        // 임시 파일로 저장
        public string SaveToTempFile()
        {
            try
            {
                // 임시 파일 경로 생성
                string tempDir = Path.GetTempPath();
                string tempFileName = $"ExcelToYaml_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string tempFile = Path.Combine(tempDir, tempFileName);

                // 현재 활성 워크북 저장
                this.Application.ActiveWorkbook.SaveCopyAs(tempFile);

                Logger.Information("임시 파일 저장: {0}", tempFile);
                return tempFile;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "임시 파일 저장 실패");
                return null;
            }
        }

        // 현재 워크시트 이름 가져오기
        public string GetActiveSheetName()
        {
            try
            {
                if (this.Application.ActiveSheet is Excel.Worksheet sheet)
                {
                    return sheet.Name;
                }
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "워크시트 이름 가져오기 실패");
                return null;
            }
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
