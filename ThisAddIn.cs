using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelToJsonAddin.Logging;
using ExcelToJsonAddin.Config;

namespace ExcelToJsonAddin
{
    public partial class ThisAddIn
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ThisAddIn>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 애드인 시작 시 초기화
            try
            {
                // Add-in 초기화 로깅
                Logger.Debug("Excel To JSON Add-in 시작");
                
                // SheetPathManager 초기화 및 설정 미리 로드
                SheetPathManager.Instance.Initialize();
                
                // 현재 워크북 설정
                if (this.Application.ActiveWorkbook != null)
                {
                    string workbookPath = this.Application.ActiveWorkbook.FullName;
                    SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
                    Logger.Information("현재 워크북 설정: {0}", workbookPath);
                }
                
                // Ribbon 인스턴스 생성 및 등록
                var ribbon = new Ribbon();
                Debug.WriteLine("Ribbon 인스턴스가 생성되었습니다.");
                Logger.Information("Excel 애드인 시작됨");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"애드인 초기화 중 오류: {ex.Message}");
                MessageBox.Show($"애드인 초기화 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                string tempFileName = $"ExcelToJson_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
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
