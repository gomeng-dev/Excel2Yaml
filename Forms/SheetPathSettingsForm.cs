using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using ExcelToJsonAddin.Config;
using System.IO;
using System.Diagnostics;

namespace ExcelToJsonAddin.Forms
{
    public partial class SheetPathSettingsForm : Form
    {
        private Dictionary<string, string> sheetPaths;
        private List<Worksheet> convertibleSheets;

        public SheetPathSettingsForm(List<Worksheet> sheets)
        {
            this.convertibleSheets = sheets;
            InitializeComponent();
            dataGridView.CellEndEdit += DataGridView_CellEndEdit;
            LoadSheetPaths();
            PopulateSheetsList();
        }

        private void LoadSheetPaths()
        {
            sheetPaths = new Dictionary<string, string>();

            // 워크북 전체 경로와 파일명 확인
            string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
            string workbookName = Path.GetFileName(fullWorkbookPath);
            Debug.WriteLine($"[SheetPathSettingsForm] 워크북 전체 경로: {fullWorkbookPath}");
            Debug.WriteLine($"[SheetPathSettingsForm] 워크북 이름: {workbookName}");

            // 설정 파일 경로 확인
            string configFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelToJsonAddin",
                "SheetPaths.xml");

            Debug.WriteLine($"[SheetPathSettingsForm] XML 설정 파일 경로: {configFilePath}");
            Debug.WriteLine($"[SheetPathSettingsForm] 설정 파일 존재 여부: {File.Exists(configFilePath)}");

            // SheetPathManager 인스턴스 초기화하고 LazyLoadSheetPaths 강제 호출
            var pathManager = SheetPathManager.Instance;
            pathManager.Initialize();

            // 이 시점에서 파일 저장된 내용 덤프
            var allWorkbooks = pathManager.GetAllWorkbookPaths();
            if (allWorkbooks != null && allWorkbooks.Count > 0)
            {
                Debug.WriteLine($"[SheetPathSettingsForm] 저장된 워크북 수: {allWorkbooks.Count}");
                foreach (var wb in allWorkbooks)
                {
                    Debug.WriteLine($"[SheetPathSettingsForm] 저장된 워크북: {wb}");
                }
            }
            else
            {
                Debug.WriteLine($"[SheetPathSettingsForm] 저장된 워크북이 없습니다.");
            }

            bool foundSettings = false;

            // 1. 전체 경로로 시도
            pathManager.SetCurrentWorkbook(fullWorkbookPath);

            // 저장된 설정 가져오기 (전체 경로)
            var savedPaths = pathManager.GetAllSheetPaths();

            if (savedPaths != null && savedPaths.Count > 0)
            {
                foundSettings = true;
                Debug.WriteLine($"[SheetPathSettingsForm] 저장된 시트 경로 설정 수: {savedPaths.Count}");
                foreach (var path in savedPaths)
                {
                    sheetPaths[path.Key] = path.Value;
                    Debug.WriteLine($"[SheetPathSettingsForm] 로드된 시트 경로: '{path.Key}' -> '{path.Value}'");
                }
            }
            else
            {
                Debug.WriteLine($"[SheetPathSettingsForm] '{fullWorkbookPath}'에 대한 저장된 시트 경로 설정이 없습니다.");
            }

            // 2. 파일명만으로도 시도
            if (!foundSettings)
            {
                pathManager.SetCurrentWorkbook(workbookName);
                savedPaths = pathManager.GetAllSheetPaths();

                if (savedPaths != null && savedPaths.Count > 0)
                {
                    foundSettings = true;
                    Debug.WriteLine($"[SheetPathSettingsForm] 파일명으로 시도 - 저장된 시트 경로 설정 수: {savedPaths.Count}");
                    foreach (var path in savedPaths)
                    {
                        sheetPaths[path.Key] = path.Value;
                        Debug.WriteLine($"[SheetPathSettingsForm] 로드된 시트 경로: '{path.Key}' -> '{path.Value}'");
                    }
                }
                else
                {
                    Debug.WriteLine($"[SheetPathSettingsForm] 파일명으로도 저장된 시트 경로 설정이 없습니다.");
                }
            }

            // 3. 전체 경로로 다시 설정 (저장 버튼이 눌렸을 때 올바른 경로를 사용하도록)
            pathManager.SetCurrentWorkbook(fullWorkbookPath);
        }

        private void PopulateSheetsList()
        {
            dataGridView.Rows.Clear();

            // 워크북 전체 경로와 파일명 확인
            string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
            string workbookName = Path.GetFileName(fullWorkbookPath);

            // SheetPathManager 인스턴스 가져오기
            var pathManager = SheetPathManager.Instance;
            pathManager.SetCurrentWorkbook(fullWorkbookPath);

            foreach (var sheet in convertibleSheets)
            {
                string sheetName = sheet.Name;
                
                // 경로가 있는지 확인
                string path = "";
                bool pathExists = sheetPaths.ContainsKey(sheetName);
                if (pathExists)
                {
                    path = sheetPaths[sheetName];
                }

                // 활성화 상태는 XML에서 가져오기 (경로 존재 여부와 독립적)
                bool enabled = pathManager.IsSheetEnabled(sheetName);
                
                // 경로가 없는데 활성화된 상태는 올바르지 않음 (경로가 없으면 비활성화 상태로 표시)
                if (string.IsNullOrEmpty(path))
                {
                    enabled = false;
                }

                int rowIndex = dataGridView.Rows.Add(sheetName, enabled, path);
                var row = dataGridView.Rows[rowIndex];
            }
        }

        private void SelectPath(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dataGridView.Rows.Count)
                return;

            var row = dataGridView.Rows[rowIndex];
            string sheetName = row.Cells[0].Value.ToString();
            string currentPath = row.Cells[2].Value?.ToString() ?? "";

            // 윈도우 탐색기 스타일 폴더 선택 다이얼로그 사용
            string selectedPath = ShowFolderBrowserDialog(sheetName, currentPath);

            if (!string.IsNullOrEmpty(selectedPath))
            {
                row.Cells[2].Value = selectedPath;
                
                // 이전에 경로가 비어있었을 때만 체크박스를 자동으로 체크
                if (string.IsNullOrEmpty(currentPath))
                {
                    row.Cells[1].Value = true;
                }
                // 이미 경로가 있었다면 체크박스 상태 유지 (사용자 설정 존중)
            }
            else if (string.IsNullOrEmpty(currentPath))
            {
                // 사용자가 폴더를 선택하지 않았고 기존 경로도 없으면 체크 해제
                row.Cells[1].Value = false;
            }
        }

        private string ShowFolderBrowserDialog(string title, string initialFolder)
        {
            // Windows 탐색기 스타일 폴더 선택 다이얼로그
            using (OpenFileDialog folderBrowser = new OpenFileDialog())
            {
                // 폴더 선택을 위한 설정
                folderBrowser.ValidateNames = false;
                folderBrowser.CheckFileExists = false;
                folderBrowser.CheckPathExists = true;
                folderBrowser.FileName = "폴더 선택";

                // 파일이 아닌 폴더만 선택하도록 함
                folderBrowser.Filter = "폴더|*.";
                folderBrowser.Title = title;

                // 초기 폴더 설정
                if (!string.IsNullOrEmpty(initialFolder) && Directory.Exists(initialFolder))
                {
                    folderBrowser.InitialDirectory = initialFolder;
                }
                else
                {
                    string defaultDir = Properties.Settings.Default.LastExportPath;
                    if (!string.IsNullOrEmpty(defaultDir) && Directory.Exists(defaultDir))
                    {
                        folderBrowser.InitialDirectory = defaultDir;
                    }
                }

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    // 선택된 파일이 아닌 선택된 폴더 경로 반환
                    return Path.GetDirectoryName(folderBrowser.FileName);
                }

                return string.Empty;
            }
        }

        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Debug.WriteLine($"[DataGridView_CellValueChanged] 행: {e.RowIndex}, 열: {e.ColumnIndex}");

            // 체크박스 변경 처리 (인덱스 1 - '활성화' 열)
            if (e.ColumnIndex == 1 && e.RowIndex >= 0)
            {
                var row = dataGridView.Rows[e.RowIndex];
                bool isChecked = (bool)row.Cells[1].Value;

                // 시트 이름 추출
                string sheetName = row.Cells[0].Value.ToString();
                Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 활성화 상태 변경: {isChecked}");

                // 항상 출력 경로 텍스트 칸은 수정 가능하게 합니다.
                row.Cells[2].ReadOnly = false;

                // 체크박스가 선택되었으나 경로가 비어있으면 폴더 선택 다이얼로그 표시
                string currentPath = row.Cells[2].Value?.ToString() ?? "";
                if (isChecked && string.IsNullOrEmpty(currentPath))
                {
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'에 대한 경로가 없어 폴더 선택 다이얼로그 표시");
                    OpenFolderSelectionDialog(e.RowIndex);
                }
            }
            // 변경된 행을 즉시 XML와 동기화
            if(e.RowIndex >= 0) UpdateSheetPathForRow(e.RowIndex);
        }

        // 새로 추가: 셀 편집 종료 시에도 XML와 동기화
        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0) UpdateSheetPathForRow(e.RowIndex);
        }

        // 공통 메서드: 특정 행의 데이터를 XML에 업데이트
        private void UpdateSheetPathForRow(int rowIndex)
        {
            var row = dataGridView.Rows[rowIndex];
            string sheetName = row.Cells[0].Value.ToString();
            bool enabled = (bool)row.Cells[1].Value;
            string path = row.Cells[2].Value?.ToString() ?? "";

            string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
            string workbookName = Path.GetFileName(fullWorkbookPath);

            var pathManager = SheetPathManager.Instance;
            pathManager.SetCurrentWorkbook(fullWorkbookPath);

            // 변경: 경로가 있는 경우, 활성화 상태와 관계없이 항상 경로 정보 저장
            if(!string.IsNullOrEmpty(path))
            {
                Debug.WriteLine($"[UpdateSheetPathForRow] 저장: 시트 '{sheetName}', 경로 '{path}', 활성화 상태: {enabled}");
                pathManager.SetSheetPath(workbookName, sheetName, path, enabled);
                if (workbookName != fullWorkbookPath)
                {
                    pathManager.SetSheetPath(fullWorkbookPath, sheetName, path, enabled);
                }
            }
            else
            {
                // 경로가 비어있는 경우에만 경로 정보 삭제
                Debug.WriteLine($"[UpdateSheetPathForRow] 제거: 시트 '{sheetName}' (경로가 비어있음)");
                pathManager.RemoveSheetPath(workbookName, sheetName);
                if (workbookName != fullWorkbookPath)
                {
                    pathManager.RemoveSheetPath(fullWorkbookPath, sheetName);
                }
            }

            pathManager.SaveSettings();
        }

        // 디자이너에서 참조하는 이벤트 핸들러 재추가
        private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 폴더 선택 버튼 클릭 시
            if (e.ColumnIndex == 3 && e.RowIndex >= 0)
            {
                Debug.WriteLine($"[DataGridView_CellContentClick] 폴더 선택 버튼 클릭: 행={e.RowIndex}");
                OpenFolderSelectionDialog(e.RowIndex);
            }
        }

        private void OpenFolderSelectionDialog(int rowIndex)
        {
            SelectPath(rowIndex);
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            // 워크북 전체 경로와 파일명
            string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
            string workbookName = Path.GetFileName(fullWorkbookPath);
            Debug.WriteLine($"[SheetPathSettingsForm] 저장 시 워크북 전체 경로: {fullWorkbookPath}");
            Debug.WriteLine($"[SheetPathSettingsForm] 저장 시 워크북 이름: {workbookName}");

            // 저장 경로 매니저 가져오기
            var pathManager = SheetPathManager.Instance;

            // 저장 전 설정 백업
            var allWorkbooks = pathManager.GetAllWorkbookPaths();
            Debug.WriteLine($"[SheetPathSettingsForm] 저장 전 워크북 수: {(allWorkbooks != null ? allWorkbooks.Count : 0)}");

            // 현재 워크북 설정 - 전체 경로와 파일명
            pathManager.SetCurrentWorkbook(fullWorkbookPath);

            // 시트별 경로 정보 저장 (변경: 경로가 있으면 활성화 상태와 관계없이 저장)
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                var row = dataGridView.Rows[i];
                string sheetName = row.Cells[0].Value.ToString();
                bool enabled = (bool)row.Cells[1].Value;
                string path = row.Cells[2].Value?.ToString() ?? "";

                if (!string.IsNullOrEmpty(path))
                {
                    Debug.WriteLine($"[SheetPathSettingsForm] 시트 경로 설정: '{sheetName}' -> '{path}', 활성화 상태: {enabled}");
                    pathManager.SetSheetPath(workbookName, sheetName, path, enabled);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetSheetPath(fullWorkbookPath, sheetName, path, enabled);
                    }
                }
                else
                {
                    // 경로가 비어있는 경우에만 정보 삭제
                    pathManager.RemoveSheetPath(workbookName, sheetName);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.RemoveSheetPath(fullWorkbookPath, sheetName);
                    }
                }
            }

            // 설정 저장
            pathManager.SaveSettings();

            // 저장 후 설정 확인
            allWorkbooks = pathManager.GetAllWorkbookPaths();
            Debug.WriteLine($"[SheetPathSettingsForm] 저장 후 워크북 수: {(allWorkbooks != null ? allWorkbooks.Count : 0)}");

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SheetPathSettingsForm_Load(object sender, EventArgs e)
        {
            // 설정 파일 경로 표시
            string configFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelToJsonAddin",
                "SheetPaths.xml");

            lblConfigPath.Text = "설정 파일 경로: " + configFilePath;

            // 경로가 긴 경우 표시를 위해 폼 크기 조정
            this.MinimumSize = new System.Drawing.Size(700, 450);
        }
    }
}
