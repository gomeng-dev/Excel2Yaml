using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelToJsonAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 애드인 시작 시 Globals 초기화
            Globals.ThisAddIn = this;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 애드인 종료 시 정리 코드
        }

        // 활성 워크시트 가져오기
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)Application.ActiveSheet;
        }

        // 활성 워크북 가져오기
        public Excel.Workbook GetActiveWorkbook()
        {
            return Application.ActiveWorkbook;
        }

        // 임시 파일로 저장
        public string SaveToTempFile()
        {
            try
            {
                string tempFile = Path.Combine(Path.GetTempPath(), $"excel2json_temp_{Guid.NewGuid()}.xlsx");
                Excel.Workbook workbook = GetActiveWorkbook();
                
                // 현재 워크북을 복사하여 임시 파일로 저장
                workbook.SaveCopyAs(tempFile);
                
                return tempFile;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"임시 파일 생성 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
