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

            // Del 키 이벤트 추가
            this.dataGridView.KeyDown += new KeyEventHandler(DataGridView_KeyDown);
            
            // 폼 리사이즈 이벤트 추가
            this.Resize += new EventHandler(SheetPathSettingsForm_Resize);
            
            // 시트 경로 설정 로드
            LoadSheetPaths();
            
            // 시트 목록 채우기
            PopulateSheetsList();
        }

        /// <summary>
        /// 폼 크기가 변경될 때 DataGridView 크기를 조정합니다.
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인수</param>
        private void SheetPathSettingsForm_Resize(object sender, EventArgs e)
        {
            AdjustDataGridViewSize();
        }
        
        /// <summary>
        /// DataGridView 크기를 폼에 맞게 조정합니다.
        /// </summary>
        private void AdjustDataGridViewSize()
        {
            if (dataGridView != null)
            {
                int margin = 40; // 좌우 여백
                dataGridView.Width = this.ClientSize.Width - margin;
                
                // 마지막 열의 너비를 자동으로 조정 (필요시)
                dataGridView.Columns[dataGridView.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                
                Debug.WriteLine($"[SheetPathSettingsForm] DataGridView 크기 조정: 너비={dataGridView.Width}, 폼너비={this.ClientSize.Width}");
            }
        }

        /// <summary>
        /// DataGridView에서 키 입력을 처리합니다.
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인수</param>
        private void DataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            // Delete 키가 눌렸을 때
            if (e.KeyCode == Keys.Delete)
            {
                Debug.WriteLine("[SheetPathSettingsForm] Delete 키 입력 감지");
                
                // 현재 선택된 셀이 있는지 확인
                if (dataGridView.CurrentCell != null)
                {
                    // 선택된 셀이 편집 가능한 텍스트 타입 셀인지 확인
                    if (dataGridView.CurrentCell.OwningColumn is DataGridViewTextBoxColumn && 
                        !dataGridView.CurrentCell.ReadOnly)
                    {
                        // 셀 값을 빈 문자열로 설정
                        dataGridView.CurrentCell.Value = string.Empty;
                        Debug.WriteLine($"[SheetPathSettingsForm] 셀 값 삭제 - 행:{dataGridView.CurrentCell.RowIndex}, 열:{dataGridView.CurrentCell.ColumnIndex}");
                        
                        // 변경 이벤트 발생
                        DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(
                            dataGridView.CurrentCell.ColumnIndex,
                            dataGridView.CurrentCell.RowIndex);
                        DataGridView_CellValueChanged(dataGridView, args);
                        
                        // 키 처리 완료 표시
                        e.Handled = true;
                    }
                }
            }
        }

        private void LoadSheetPaths()
        {
            sheetPaths = new Dictionary<string, string>();
            
            // convertibleSheets 유효성 검사
            if (convertibleSheets == null || convertibleSheets.Count == 0)
            {
                Debug.WriteLine("[LoadSheetPaths] 오류: convertibleSheets가 null이거나 비어 있습니다.");
                return;
            }
            
            if (convertibleSheets[0] == null || convertibleSheets[0].Parent == null)
            {
                Debug.WriteLine("[LoadSheetPaths] 오류: convertibleSheets[0] 또는 Parent가 null입니다.");
                return;
            }

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
            try
            {
                // 이벤트 핸들러 임시 제거 (값 채우는 동안 이벤트 발생 방지)
                dataGridView.CellValueChanged -= DataGridView_CellValueChanged;
                dataGridView.CellEndEdit -= DataGridView_CellEndEdit;
                
                dataGridView.Rows.Clear();

                // 워크북 전체 경로와 파일명 확인
                if (convertibleSheets == null || convertibleSheets.Count == 0 || 
                    convertibleSheets[0] == null || convertibleSheets[0].Parent == null)
                {
                    Debug.WriteLine("[PopulateSheetsList] 오류: convertibleSheets 또는 Parent가 null입니다.");
                    
                    // 사용자에게 알림
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.", 
                        "시트 목록 없음", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    // 이벤트 핸들러 다시 등록
                    dataGridView.CellEndEdit += DataGridView_CellEndEdit;
                    dataGridView.CellValueChanged += DataGridView_CellValueChanged;
                    return;
                }
                
                string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
                string workbookName = Path.GetFileName(fullWorkbookPath);

                // SheetPathManager 인스턴스 가져오기
                var pathManager = SheetPathManager.Instance;
                pathManager.SetCurrentWorkbook(fullWorkbookPath);

                // YAML 선택적 필드 처리 컬럼이 존재하는지 확인
                bool hasYamlColumn = false;
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    if (dataGridView.Columns[i].Name == "YamlEmptyFields")
                    {
                        hasYamlColumn = true;
                        break;
                    }
                }

                // Flow Style 설정 컬럼 추가 확인 및 추가
                bool hasFlowStyleFieldsColumn = false;
                bool hasFlowStyleItemsFieldsColumn = false;
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    if (dataGridView.Columns[i].Name == "FlowStyleFieldsColumn")
                    {
                        hasFlowStyleFieldsColumn = true;
                    }
                    if (dataGridView.Columns[i].Name == "FlowStyleItemsFieldsColumn")
                    {
                        hasFlowStyleItemsFieldsColumn = true;
                    }
                }

                if (!hasFlowStyleFieldsColumn)
                {
                    DataGridViewTextBoxColumn flowStyleFieldsColumn = new DataGridViewTextBoxColumn();
                    flowStyleFieldsColumn.HeaderText = "Flow 필드";
                    flowStyleFieldsColumn.Name = "FlowStyleFieldsColumn";
                    flowStyleFieldsColumn.Width = 100;
                    flowStyleFieldsColumn.ToolTipText = "Flow 스타일로 변환할 필드 (예: \"details,info\")";
                    dataGridView.Columns.Add(flowStyleFieldsColumn);
                }
                
                if (!hasFlowStyleItemsFieldsColumn)
                {
                    DataGridViewTextBoxColumn flowStyleItemsFieldsColumn = new DataGridViewTextBoxColumn();
                    flowStyleItemsFieldsColumn.HeaderText = "Flow 항목 필드";
                    flowStyleItemsFieldsColumn.Name = "FlowStyleItemsFieldsColumn";
                    flowStyleItemsFieldsColumn.Width = 110;
                    flowStyleItemsFieldsColumn.ToolTipText = "Flow 스타일로 변환할 항목 필드 (예: \"triggers,events\")";
                    dataGridView.Columns.Add(flowStyleItemsFieldsColumn);
                }

                foreach (var sheet in convertibleSheets)
                {
                    try
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
                        
                        // YAML 선택적 필드 처리 상태 가져오기
                        bool yamlEmptyFields = pathManager.GetYamlEmptyFieldsOption(sheetName);
                        
                        // 후처리 키 경로 가져오기
                        string mergeKeyPaths = pathManager.GetMergeKeyPaths(sheetName);
                        Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}', 후처리 키 경로: {mergeKeyPaths}");
                        
                        // Flow Style 설정 가져오기
                        string flowStyleConfig = pathManager.GetFlowStyleConfig(sheetName);
                        Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}', Flow Style 설정: {flowStyleConfig}");
                        
                        // Flow Style 설정을 필드와 항목으로 분리
                        string flowStyleFields = "";
                        string flowStyleItemsFields = "";
                        if (!string.IsNullOrWhiteSpace(flowStyleConfig))
                        {
                            string[] parts = flowStyleConfig.Split('|');
                            if (parts.Length >= 1)
                                flowStyleFields = parts[0];
                            if (parts.Length >= 2)
                                flowStyleItemsFields = parts[1];
                        }
                        
                        // 키 경로 데이터 파싱하여 각 컬럼에 설정
                        string idPath = "";      // 기본값 제거
                        string mergePaths = ""; // 기본값 제거
                        string keyPaths = "";
                        
                        // 설정된 값이 있으면 파싱
                        if (!string.IsNullOrWhiteSpace(mergeKeyPaths))
                        {
                            string[] parts = mergeKeyPaths.Split('|');
                            if (parts.Length >= 1)
                                idPath = parts[0]; // 빈 문자열도 그대로 사용
                            if (parts.Length >= 2)
                                mergePaths = parts[1]; // 빈 문자열도 그대로 사용
                            if (parts.Length >= 3)
                                keyPaths = parts[2]; // 빈 문자열도 그대로 사용
                        }
                        
                        Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}', ID 경로: {idPath}, 병합 경로: {mergePaths}, 키 경로: {keyPaths}");
                        
                        // 경로가 없는데 활성화된 상태는 올바르지 않음 (경로가 없으면 비활성화 상태로 표시)
                        if (string.IsNullOrEmpty(path))
                        {
                            enabled = false;
                        }

                        // 상세 디버그 정보 추가
                        Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}', YAML 선택적 필드: {yamlEmptyFields}");

                        // 데이터그리드뷰뷰에 행 추가
                        int rowIndex = dataGridView.Rows.Add();
                        var row = dataGridView.Rows[rowIndex];
                        
                        // 각 셀에 직접 값 설정
                        if (row.Cells.Count > 0)
                            row.Cells[0].Value = sheetName;
                            
                        if (row.Cells.Count > 1)
                            row.Cells[1].Value = enabled;
                            
                        if (row.Cells.Count > 2)
                            row.Cells[2].Value = path;
                            
                        // YAML 선택적 필드 처리 컬럼이 존재하면 값 설정
                        if (hasYamlColumn && row.Cells.Count > 4)
                            row.Cells[4].Value = yamlEmptyFields;
                            
                        // 후처리 키 경로 컬럼에 값 설정 (인덱스 5, 6, 7)
                        if (row.Cells.Count > 5)
                            row.Cells[5].Value = idPath;
                        if (row.Cells.Count > 6)
                            row.Cells[6].Value = mergePaths;
                        if (row.Cells.Count > 7)
                            row.Cells[7].Value = keyPaths;
                            
                        // Flow Style 필드 설정 컬럼이 존재하면 값 설정
                        if (hasFlowStyleFieldsColumn && row.Cells.Count > 8)
                            row.Cells[8].Value = flowStyleFields;
                        
                        // Flow Style 항목 필드 설정 컬럼이 존재하면 값 설정
                        if (hasFlowStyleItemsFieldsColumn && row.Cells.Count > 9)
                            row.Cells[9].Value = flowStyleItemsFields;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[PopulateSheetsList] 시트 처리 중 예외 발생: {ex.Message}");
                    }
                }
                
                // 모든 행을 추가한 후 이벤트 핸들러 다시 등록
                dataGridView.CellEndEdit += DataGridView_CellEndEdit;
                dataGridView.CellValueChanged += DataGridView_CellValueChanged;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PopulateSheetsList] 예외 발생: {ex.Message}\n{ex.StackTrace}");
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
            try
            {
                Debug.WriteLine($"[DataGridView_CellValueChanged] 행: {e.RowIndex}, 열: {e.ColumnIndex}");

                // 행 인덱스가 유효하지 않으면 처리하지 않음
                if (e.RowIndex < 0 || e.RowIndex >= dataGridView.Rows.Count)
                {
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 유효하지 않은 행 인덱스: {e.RowIndex}");
                    return;
                }

                var row = dataGridView.Rows[e.RowIndex];
                
                // 체크박스 변경 처리 (인덱스 1 - '활성화' 열)
                if (e.ColumnIndex == 1 && row.Cells.Count > 1 && row.Cells[1].Value != null)
                {
                    bool isChecked = (bool)row.Cells[1].Value;

                    // 시트 이름이 유효한지 확인
                    if (row.Cells[0].Value == null)
                    {
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 이름이 null입니다.");
                        return;
                    }

                    // 시트 이름 추출
                    string sheetName = row.Cells[0].Value.ToString();
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 활성화 상태 변경: {isChecked}");

                    // 항상 출력 경로 텍스트 칸은 수정 가능하게 합니다.
                    if (row.Cells.Count > 2)
                    {
                        row.Cells[2].ReadOnly = false;
                    }

                    // 체크박스가 선택되었으나 경로가 비어있으면 폴더 선택 다이얼로그 표시
                    string currentPath = row.Cells.Count > 2 && row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "";
                    if (isChecked && string.IsNullOrEmpty(currentPath))
                    {
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'에 대한 경로가 없어 폴더 선택 다이얼로그 표시");
                        OpenFolderSelectionDialog(e.RowIndex);
                    }
                }
                // YAML 선택적 필드 처리 체크박스 변경 처리 (인덱스 4 - 'YAML 선택적 필드 처리' 열)
                else if (e.ColumnIndex == 4 && row.Cells.Count > 4 && row.Cells[0].Value != null)
                {
                    bool yamlEmptyFields = row.Cells[4].Value != null ? (bool)row.Cells[4].Value : false;
                    string sheetName = row.Cells[0].Value.ToString();
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 YAML 선택적 필드 처리 상태 변경: {yamlEmptyFields}");
                }
                // ID 경로 필드 변경 처리 (인덱스 5)
                else if (e.ColumnIndex == 5 && row.Cells.Count > 5 && row.Cells[0].Value != null)
                {
                    string idPath = row.Cells[5].Value?.ToString() ?? "";
                    string sheetName = row.Cells[0].Value.ToString();
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 ID 경로 변경: '{idPath}'");
                }
                // 병합 경로 필드 변경 처리 (인덱스 6)
                else if (e.ColumnIndex == 6 && row.Cells.Count > 6 && row.Cells[0].Value != null)
                {
                    string mergePaths = row.Cells[6].Value?.ToString() ?? "";
                    string sheetName = row.Cells[0].Value.ToString();
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 병합 경로 변경: '{mergePaths}'");
                }
                // 키 경로 필드 변경 처리 (인덱스 7)
                else if (e.ColumnIndex == 7 && row.Cells.Count > 7 && row.Cells[0].Value != null)
                {
                    string keyPaths = row.Cells[7].Value?.ToString() ?? "";
                    string sheetName = row.Cells[0].Value.ToString();
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 키 경로 변경: '{keyPaths}'");
                }
                // Flow Style 필드 설정 필드 변경 처리
                if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "FlowStyleFieldsColumn")
                {
                    UpdateSheetPathForRow(e.RowIndex);
                    string sheetName = dataGridView.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
                    string flowStyleFields = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 Flow Style 필드 설정 변경: '{flowStyleFields}'");
                }
                
                // Flow Style 항목 필드 설정 필드 변경 처리
                if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "FlowStyleItemsFieldsColumn")
                {
                    UpdateSheetPathForRow(e.RowIndex);
                    string sheetName = dataGridView.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
                    string flowStyleItemsFields = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 Flow Style 항목 필드 설정 변경: '{flowStyleItemsFields}'");
                }
                
                // 변경된 행을 즉시 XML와 동기화
                if(e.RowIndex >= 0) UpdateSheetPathForRow(e.RowIndex);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DataGridView_CellValueChanged] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // 새로 추가: 셀 편집 종료 시에도 XML와 동기화
        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex >= 0) UpdateSheetPathForRow(e.RowIndex);
        }

        // 공통 메서드: 특정 행의 데이터를 XML에 업데이트
        private void UpdateSheetPathForRow(int rowIndex)
        {
            try 
            {
                var row = dataGridView.Rows[rowIndex];
                if (row.Cells.Count <= 0 || row.Cells[0].Value == null)
                {
                    Debug.WriteLine($"[UpdateSheetPathForRow] 오류: 행 {rowIndex}의 셀 0에 값이 없습니다.");
                    return;
                }

                string sheetName = row.Cells[0].Value.ToString();
                
                // 활성화 상태 확인 (인덱스 1)
                bool enabled = row.Cells.Count > 1 && row.Cells[1].Value != null ? (bool)row.Cells[1].Value : false;
                
                // 경로 확인 (인덱스 2)
                string path = row.Cells.Count > 2 && row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : "";
                
                // YAML 선택적 필드 처리 상태 확인 (인덱스 4)
                bool yamlEmptyFields = false;
                if (row.Cells.Count > 4 && row.Cells[4].Value != null)
                {
                    yamlEmptyFields = (bool)row.Cells[4].Value;
                }
                
                // 후처리 키 경로 확인 (인덱스 5, 6, 7)
                string idPath = "";
                string mergePaths = "";
                string keyPaths = "";
                
                if (row.Cells.Count > 5)
                    idPath = row.Cells[5].Value?.ToString() ?? "";
                if (row.Cells.Count > 6)
                    mergePaths = row.Cells[6].Value?.ToString() ?? "";
                if (row.Cells.Count > 7)
                    keyPaths = row.Cells[7].Value?.ToString() ?? "";
                
                // Flow Style 필드 설정 확인
                string flowStyleFields = "";
                string flowStyleItemsFields = "";
                
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.OwningColumn.Name == "FlowStyleFieldsColumn")
                    {
                        flowStyleFields = cell.Value?.ToString() ?? "";
                    }
                    else if (cell.OwningColumn.Name == "FlowStyleItemsFieldsColumn")
                    {
                        flowStyleItemsFields = cell.Value?.ToString() ?? "";
                    }
                }
                
                // Flow Style 설정 합치기
                string flowStyleConfig = $"{flowStyleFields}|{flowStyleItemsFields}";
                
                // 합친 문자열 생성
                string mergeKeyPaths = $"{idPath}|{mergePaths}|{keyPaths}";

                // 워크북 경로가 없으면 함수 종료
                if (convertibleSheets == null || convertibleSheets.Count == 0 || 
                    convertibleSheets[0] == null || convertibleSheets[0].Parent == null)
                {
                    Debug.WriteLine($"[UpdateSheetPathForRow] 오류: convertibleSheets 또는 Parent가 null입니다.");
                    return;
                }

                string fullWorkbookPath = convertibleSheets[0].Parent.FullName;
                string workbookName = Path.GetFileName(fullWorkbookPath);

                var pathManager = SheetPathManager.Instance;
                pathManager.SetCurrentWorkbook(fullWorkbookPath);

                // 변경: 경로가 있는 경우, 활성화 상태와 관계없이 항상 경로 정보 저장
                if(!string.IsNullOrEmpty(path))
                {
                    Debug.WriteLine($"[UpdateSheetPathForRow] 저장: 시트 '{sheetName}', 경로 '{path}', 활성화 상태: {enabled}, YAML 선택적 필드: {yamlEmptyFields}");
                    pathManager.SetSheetPath(workbookName, sheetName, path, enabled, yamlEmptyFields);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetSheetPath(fullWorkbookPath, sheetName, path, enabled, yamlEmptyFields);
                    }
                    
                    // 후처리 키 경로 설정 저장
                    Debug.WriteLine($"[UpdateSheetPathForRow] 후처리 키 경로 저장: 시트 '{sheetName}', 값: '{mergeKeyPaths}'");
                    pathManager.SetMergeKeyPaths(workbookName, sheetName, mergeKeyPaths);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetMergeKeyPaths(fullWorkbookPath, sheetName, mergeKeyPaths);
                    }
                    
                    // Flow Style 설정 저장
                    Debug.WriteLine($"[UpdateSheetPathForRow] Flow Style 설정 저장: 시트 '{sheetName}', 값: '{flowStyleConfig}'");
                    pathManager.SetFlowStyleConfig(workbookName, sheetName, flowStyleConfig);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetFlowStyleConfig(fullWorkbookPath, sheetName, flowStyleConfig);
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
            catch (Exception ex)
            {
                Debug.WriteLine($"[UpdateSheetPathForRow] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
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
            try
            {
                // 워크북 전체 경로와 파일명 확인
                string workbookFullPath = convertibleSheets[0].Parent.FullName;
                string workbookName = Path.GetFileName(workbookFullPath);
                Debug.WriteLine($"[SaveButton_Click] 워크북 전체 경로: {workbookFullPath}");
                Debug.WriteLine($"[SaveButton_Click] 워크북 이름: {workbookName}");

                // SheetPathManager 인스턴스 가져오기
                var pathManager = SheetPathManager.Instance;
                // 기본적으로 전체 경로 사용
                pathManager.SetCurrentWorkbook(workbookFullPath);

                // 각 행에 대해 설정 저장
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    var row = dataGridView.Rows[i];
                    if (row.Cells.Count < 3 || row.Cells[0].Value == null)
                        continue;

                    string sheetName = row.Cells[0].Value.ToString();
                    bool isEnabled = row.Cells.Count > 1 && row.Cells[1].Value != null ? 
                                     Convert.ToBoolean(row.Cells[1].Value) : false;
                    string path = row.Cells.Count > 2 && row.Cells[2].Value != null ? 
                                  row.Cells[2].Value.ToString() : "";

                    // 활성화된 시트에 대해서만 경로 저장 검증
                    if (isEnabled && string.IsNullOrEmpty(path))
                    {
                        // 경로가 없는데 활성화하려고 하면 경고 표시
                        MessageBox.Show($"시트 '{sheetName}'에 대한 경로가 설정되지 않았습니다. 활성화하려면 경로를 지정하세요.",
                                       "경로 필요", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    Debug.WriteLine($"[SaveButton_Click] 시트 경로 저장: {sheetName} -> '{path}', 활성화: {isEnabled}");

                    // 설정 저장 (전체 경로와 파일명 모두에 설정)
                    pathManager.SetSheetPath(workbookFullPath, sheetName, path, isEnabled);
                    pathManager.SetSheetPath(workbookName, sheetName, path, isEnabled);

                    // YAML 선택적 필드 처리 설정 저장
                    if (row.Cells.Count > 4 && row.Cells[4].Value != null)
                    {
                        bool yamlEmptyFields = Convert.ToBoolean(row.Cells[4].Value);
                        pathManager.SetSheetPath(workbookFullPath, sheetName, path, isEnabled, yamlEmptyFields);
                        pathManager.SetSheetPath(workbookName, sheetName, path, isEnabled, yamlEmptyFields);
                    }
                    
                    // 후처리 키 경로 설정 저장
                    if (row.Cells.Count > 7)
                    {
                        string idPath = row.Cells[5].Value?.ToString() ?? "";
                        string mergePaths = row.Cells[6].Value?.ToString() ?? "";
                        string keyPaths = row.Cells[7].Value?.ToString() ?? "";
                        
                        // 합친 문자열 생성
                        string mergeKeyPaths = $"{idPath}|{mergePaths}|{keyPaths}";
                        
                        pathManager.SetMergeKeyPaths(workbookFullPath, sheetName, mergeKeyPaths);
                        pathManager.SetMergeKeyPaths(workbookName, sheetName, mergeKeyPaths);
                        Debug.WriteLine($"[SaveButton_Click] 후처리 키 경로 저장: {sheetName} -> ID 경로: '{idPath}', 병합 경로: '{mergePaths}', 키 경로: '{keyPaths}'");
                    }
                    
                    // Flow Style 설정
                    string flowStyleFields = "";
                    string flowStyleItemsFields = "";
                    
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Name == "FlowStyleFieldsColumn")
                        {
                            flowStyleFields = cell.Value?.ToString() ?? "";
                        }
                        else if (cell.OwningColumn.Name == "FlowStyleItemsFieldsColumn")
                        {
                            flowStyleItemsFields = cell.Value?.ToString() ?? "";
                        }
                    }
                    
                    // Flow Style 설정 저장
                    string flowStyleConfig = $"{flowStyleFields}|{flowStyleItemsFields}";
                    Debug.WriteLine($"[SaveButton_Click] Flow Style 설정 저장: {sheetName} -> '{flowStyleConfig}'");
                    pathManager.SetFlowStyleConfig(workbookFullPath, sheetName, flowStyleConfig);
                    pathManager.SetFlowStyleConfig(workbookName, sheetName, flowStyleConfig);
                }

                // 전체 설정 저장
                pathManager.SaveSettings();

                // 사용자에게 저장 완료 메시지 표시
                MessageBox.Show("설정이 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 폼 닫기
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"설정 저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[SaveButton_Click] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SheetPathSettingsForm_Load(object sender, EventArgs e)
        {
            lblConfigPath.Text = $"설정 파일 경로: {SheetPathManager.GetConfigFilePath()}";
            
            // 초기 DataGridView 크기 조정
            AdjustDataGridViewSize();
        }
    }
}
