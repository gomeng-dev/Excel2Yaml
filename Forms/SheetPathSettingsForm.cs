using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using ExcelToJsonAddin.Core;
using ExcelToJsonAddin.Config;

namespace ExcelToJsonAddin.Forms
{
    public partial class SheetPathSettingsForm : Form
    {
        private Dictionary<string, string> sheetPaths;
        private List<Worksheet> convertibleSheets;
        
        // Excel 설정 관리자 추가
        private ExcelConfigManager excelConfigManager;

        public SheetPathSettingsForm(List<Worksheet> sheets)
        {
            this.convertibleSheets = sheets;
            InitializeComponent();

            // Del 키 이벤트 추가
            this.dataGridView.KeyDown += new KeyEventHandler(DataGridView_KeyDown);
            
            // 폼 리사이즈 이벤트 추가
            this.Resize += new EventHandler(SheetPathSettingsForm_Resize);
            
            // Excel 설정 관리자 초기화
            excelConfigManager = ExcelConfigManager.Instance;
            
            // 현재 워크북 설정
            if (sheets.Count > 0)
            {
                string workbookPath = sheets[0].Parent.FullName;
                excelConfigManager.SetCurrentWorkbook(workbookPath);
                
                // 설정 시트 존재 여부 확인 및 생성
                excelConfigManager.EnsureConfigSheetExists();
                
                // 설정 로드
                excelConfigManager.LoadAllSettings();
            }
            
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
            // 시트별 경로 및 활성화 상태 로드
            sheetPaths = new Dictionary<string, string>();
            
            try
            {
                // 워크북 경로 설정
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                
                if (app.ActiveWorkbook != null)
                {
                    string workbookPath = app.ActiveWorkbook.FullName;
                    Debug.WriteLine($"[LoadSheetPaths] 현재 워크북 경로: {workbookPath}");
                    SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
                    
                    // Excel 설정 관리자 초기화
                    if (excelConfigManager != null)
                    {
                        excelConfigManager.SetCurrentWorkbook(workbookPath);
                    }
                }
                else
                {
                    Debug.WriteLine($"[LoadSheetPaths] 활성 워크북이 없습니다.");
                }
                
                sheetPaths = SheetPathManager.Instance.GetAllSheetPaths();
                Debug.WriteLine($"[LoadSheetPaths] 로드된 시트 경로 수: {sheetPaths.Count}");
                
                foreach (var path in sheetPaths)
                {
                    Debug.WriteLine($"[LoadSheetPaths] 로드된 시트 경로: 시트='{path.Key}', 경로='{path.Value}'");
                }
                
                // DataGridView 초기화
                dataGridView.Rows.Clear();
                
                foreach (var sheet in convertibleSheets)
                {
                    string sheetName = sheet.Name;
                    string displayName = sheetName;
                    
                    if (displayName.StartsWith("!"))
                    {
                        displayName = displayName.Substring(1);
                    }
                    
                    string sheetPath = "";
                    bool isEnabled = true;
                    
                    if (sheetPaths.ContainsKey(sheetName))
                    {
                        sheetPath = sheetPaths[sheetName];
                        isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                    }
                    
                    // YAML 선택적 필드 옵션 가져오기 (우선순위 변경: Excel > XML)
                    bool yamlOption = false;
                    
                    if (excelConfigManager != null)
                    {
                        // Excel 설정 먼저 확인
                        yamlOption = excelConfigManager.GetConfigBool(sheetName, "YamlEmptyFields", false);
                        
                        // Excel 설정이 없으면 XML 설정 확인
                        if (!yamlOption)
                        {
                            yamlOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                }
            }
            else
            {
                        // XML 설정만 확인
                        yamlOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                    }
                    
                    // 병합 키 경로 설정 (우선순위 변경: Excel > XML)
                    string mergeKeyPaths = "";
                    
                    if (excelConfigManager != null)
                    {
                        // Excel 설정 먼저 확인
                        mergeKeyPaths = excelConfigManager.GetConfigValue(sheetName, "MergeKeyPaths", "");
                        
                        // Excel 설정이 없으면 XML 설정 확인
                        if (string.IsNullOrEmpty(mergeKeyPaths))
                        {
                            mergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
                }
            }
            else
            {
                        // XML 설정만 확인
                        mergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
                    }
                    
                    // Flow 스타일 설정 (우선순위 변경: Excel > XML)
                    string flowStyle = "";
                    
                    if (excelConfigManager != null)
                    {
                        // Excel 설정 먼저 확인
                        flowStyle = excelConfigManager.GetConfigValue(sheetName, "FlowStyle", "");
                        
                        // Excel 설정이 없으면 XML 설정 확인
                        if (string.IsNullOrEmpty(flowStyle))
                        {
                            flowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName);
                    }
                }
                else
                {
                        // XML 설정만 확인
                        flowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName);
                    }
                    
                    int rowIndex = dataGridView.Rows.Add();
                    var row = dataGridView.Rows[rowIndex];
                    
                    row.Cells["SheetNameColumn"].Value = displayName;
                    row.Cells["SheetNameColumn"].Tag = sheetName;
                    row.Cells["EnabledColumn"].Value = isEnabled;
                    row.Cells["PathColumn"].Value = sheetPath;
                    row.Cells["YamlEmptyFields"].Value = yamlOption;
                    row.Cells["MergePathsColumn"].Value = mergeKeyPaths;
                    row.Cells["FlowStyleFieldsColumn"].Value = flowStyle;
                }
                
                // 존재하는 모든 경로 설정 로드 (표시되지 않은 시트도)
                foreach (var path in sheetPaths)
                {
                    string sheetName = path.Key;
                    bool found = false;
                    
                    foreach (var sheet in convertibleSheets)
                    {
                        if (sheet.Name == sheetName)
                        {
                            found = true;
                            break;
                        }
                    }
                    
                    if (!found)
                    {
                        string displayName = sheetName;
                        if (displayName.StartsWith("!"))
                        {
                            displayName = displayName.Substring(1);
                        }
                        
                        bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                        
                        // YAML 선택적 필드 옵션 가져오기 (우선순위 변경: Excel > XML)
                        bool yamlOption = false;
                        
                        if (excelConfigManager != null)
                        {
                            // Excel 설정 먼저 확인
                            yamlOption = excelConfigManager.GetConfigBool(sheetName, "YamlEmptyFields", false);
                            
                            // Excel 설정이 없으면 XML 설정 확인
                            if (!yamlOption)
                            {
                                yamlOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                            }
                        }
                        else
                        {
                            // XML 설정만 확인
                            yamlOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                        }
                        
                        // 병합 키 경로 설정 (우선순위 변경: Excel > XML)
                        string mergeKeyPaths = "";
                        
                        if (excelConfigManager != null)
                        {
                            // Excel 설정 먼저 확인
                            mergeKeyPaths = excelConfigManager.GetConfigValue(sheetName, "MergeKeyPaths", "");
                            
                            // Excel 설정이 없으면 XML 설정 확인
                            if (string.IsNullOrEmpty(mergeKeyPaths))
                            {
                                mergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
                            }
                        }
                        else
                        {
                            // XML 설정만 확인
                            mergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
                        }
                        
                        // Flow 스타일 설정 (우선순위 변경: Excel > XML)
                        string flowStyle = "";
                        
                        if (excelConfigManager != null)
                        {
                            // Excel 설정 먼저 확인
                            flowStyle = excelConfigManager.GetConfigValue(sheetName, "FlowStyle", "");
                            
                            // Excel 설정이 없으면 XML 설정 확인
                            if (string.IsNullOrEmpty(flowStyle))
                            {
                                flowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName);
                            }
                        }
                        else
                        {
                            // XML 설정만 확인
                            flowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName);
                        }
                        
                        int rowIndex = dataGridView.Rows.Add();
                        var row = dataGridView.Rows[rowIndex];
                        
                        row.Cells["SheetNameColumn"].Value = displayName;
                        row.Cells["SheetNameColumn"].Tag = sheetName;
                        row.Cells["EnabledColumn"].Value = isEnabled;
                        row.Cells["PathColumn"].Value = path.Value;
                        row.Cells["YamlEmptyFields"].Value = yamlOption;
                        row.Cells["MergePathsColumn"].Value = mergeKeyPaths;
                        row.Cells["FlowStyleFieldsColumn"].Value = flowStyle;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"시트별 경로 설정을 로드하는 중 오류가 발생했습니다: {ex.Message}", 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PopulateSheetsList()
                {
                    try
                    {
                // DataGridView 초기화
                dataGridView.Rows.Clear();
                
                var pathManager = SheetPathManager.Instance;
                bool showExtraColumns = true;
                
                foreach (Worksheet sheet in convertibleSheets)
                {
                    if (sheet == null)
                        continue;
                    
                    // 시트 이름
                        string sheetName = sheet.Name;
                    if (string.IsNullOrEmpty(sheetName))
                        continue;
                    
                    // '!'로 시작하는 경우 제거
                    if (sheetName.StartsWith("!"))
                    {
                        sheetName = sheetName.Substring(1);
                    }

                    // XML에서 설정 로드
                    string savePath = pathManager.GetSheetPath(sheetName) ?? "";
                    bool isEnabled = pathManager.IsSheetEnabled(sheetName);
                    
                    // YAML 설정은 Excel에서만 로드
                    bool yamlOption = false;
                    string mergeKeyPaths = "";
                    string flowStyleConfig = "";
                    
                    // Excel 설정 로드 (YAML, MergeKeyPaths, FlowStyle은 Excel에서만 로드)
                    if (excelConfigManager != null)
                    {
                        // 활성화 여부 - XML에 없으면 Excel에서 로드
                        string excelEnabled = excelConfigManager.GetConfigValue(sheetName, "Enabled");
                        if (string.IsNullOrEmpty(savePath) && !string.IsNullOrEmpty(excelEnabled))
                        {
                            isEnabled = bool.Parse(excelEnabled);
                        }
                        
                        // YAML 선택적 필드 - Excel에서만 로드
                        string excelYamlOption = excelConfigManager.GetConfigValue(sheetName, "YamlEmptyFields");
                        if (!string.IsNullOrEmpty(excelYamlOption))
                        {
                            yamlOption = bool.Parse(excelYamlOption);
                        }
                        
                        // 병합 키 경로 - Excel에서만 로드
                        mergeKeyPaths = excelConfigManager.GetConfigValue(sheetName, "MergeKeyPaths", "");
                        
                        // Flow 스타일 - Excel에서만 로드
                        flowStyleConfig = excelConfigManager.GetConfigValue(sheetName, "FlowStyle", "");
                    }
                    
                    // 키 경로 분리 (ID 경로 | 병합 경로 | 키 경로)
                    string idPath = "";
                    string mergePaths = "";
                        string keyPaths = "";
                        
                    if (!string.IsNullOrEmpty(mergeKeyPaths))
                        {
                            string[] parts = mergeKeyPaths.Split('|');
                        if (parts.Length >= 1) idPath = parts[0];
                        if (parts.Length >= 2) mergePaths = parts[1];
                        if (parts.Length >= 3) keyPaths = parts[2];
                    }
                    
                    // Flow Style 설정 분리 (Flow Style 필드 | Flow Style 항목 필드)
                    string flowStyleFields = "";
                    string flowStyleItemsFields = "";
                    
                    if (!string.IsNullOrEmpty(flowStyleConfig))
                    {
                        string[] parts = flowStyleConfig.Split('|');
                        if (parts.Length >= 1) flowStyleFields = parts[0];
                        if (parts.Length >= 2) flowStyleItemsFields = parts[1];
                    }
                    
                    // 행 추가 준비
                        int rowIndex = dataGridView.Rows.Add();
                        var row = dataGridView.Rows[rowIndex];
                        
                    // 주요 열 채우기
                    row.Cells[0].Value = sheetName;                // 시트 이름
                    row.Cells[1].Value = isEnabled;                // 활성화 여부
                    row.Cells[2].Value = savePath;                 // 저장 경로
                    row.Cells[4].Value = yamlOption;               // YAML 선택적 필드 처리
                    
                    // 고급 옵션 열 채우기
                    if (row.Cells.Count > 5 && showExtraColumns)
                    {
                        row.Cells[5].Value = idPath;               // ID 경로
                        row.Cells[6].Value = mergePaths;           // 병합 경로
                        row.Cells[7].Value = keyPaths;             // 키 경로
                    }
                    
                    // Flow Style 열 채우기
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.OwningColumn.Name == "FlowStyleFieldsColumn")
                        {
                            cell.Value = flowStyleFields;
                        }
                        else if (cell.OwningColumn.Name == "FlowStyleItemsFieldsColumn")
                        {
                            cell.Value = flowStyleItemsFields;
                        }
                    }
                }
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
                
                // 경로 선택 후 즉시 XML 설정에 저장
                Debug.WriteLine($"[SelectPath] 경로 선택 후 즉시 저장: 시트={sheetName}, 경로={selectedPath}");
                UpdateSheetPathForRow(rowIndex);
                
                // 설정 저장 확인
                SheetPathManager.Instance.SaveSettings();
                Debug.WriteLine($"[SelectPath] 설정 저장 완료");
            }
            else if (string.IsNullOrEmpty(currentPath))
            {
                // 사용자가 폴더를 선택하지 않았고 기존 경로도 없으면 체크 해제
                row.Cells[1].Value = false;
            }
        }

        private string ShowFolderBrowserDialog(string title, string initialFolder)
        {
            Debug.WriteLine($"[ShowFolderBrowserDialog] 시작: 제목='{title}', 초기 폴더='{initialFolder}'");
            
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
                    Debug.WriteLine($"[ShowFolderBrowserDialog] 초기 폴더 설정: '{initialFolder}'");
                }
                else
                {
                    string defaultDir = Properties.Settings.Default.LastExportPath;
                    if (!string.IsNullOrEmpty(defaultDir) && Directory.Exists(defaultDir))
                    {
                        folderBrowser.InitialDirectory = defaultDir;
                        Debug.WriteLine($"[ShowFolderBrowserDialog] 기본 폴더 설정: '{defaultDir}'");
                    }
                    else
                    {
                        Debug.WriteLine($"[ShowFolderBrowserDialog] 초기 폴더 없음, 기본 폴더도 없음");
                    }
                }

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    // 선택된 파일이 아닌 선택된 폴더 경로 반환
                    string selectedPath = Path.GetDirectoryName(folderBrowser.FileName);
                    Debug.WriteLine($"[ShowFolderBrowserDialog] 폴더 선택 완료: '{selectedPath}'");
                    return selectedPath;
                }

                Debug.WriteLine($"[ShowFolderBrowserDialog] 폴더 선택 취소됨");
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
                    Debug.WriteLine($"[UpdateSheetPathForRow] 저장: 시트 '{sheetName}', 경로 '{path}', 활성화 상태: {enabled}");
                    // XML에는 경로와 활성화 상태만 저장 (YAML 관련 설정은 제외)
                    pathManager.SetSheetPath(workbookName, sheetName, path);
                    pathManager.SetSheetEnabled(workbookName, sheetName, enabled);
                    
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetSheetPath(fullWorkbookPath, sheetName, path);
                        pathManager.SetSheetEnabled(fullWorkbookPath, sheetName, enabled);
                    }
                    
                    // Excel 설정에 YAML 관련 설정 저장
                    if (excelConfigManager != null)
                    {
                        // YAML 선택적 필드 처리 옵션 저장
                        excelConfigManager.SetConfigValue(sheetName, "YamlEmptyFields", yamlEmptyFields.ToString());
                    
                    // 후처리 키 경로 설정 저장
                    Debug.WriteLine($"[UpdateSheetPathForRow] 후처리 키 경로 저장: 시트 '{sheetName}', 값: '{mergeKeyPaths}'");
                        excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPaths);
                    
                    // Flow Style 설정 저장
                    Debug.WriteLine($"[UpdateSheetPathForRow] Flow Style 설정 저장: 시트 '{sheetName}', 값: '{flowStyleConfig}'");
                        excelConfigManager.SetConfigValue(sheetName, "FlowStyle", flowStyleConfig);
                    }
                }
                else
                {
                    // 경로가 비어있더라도 활성화 상태는 저장
                    Debug.WriteLine($"[UpdateSheetPathForRow] 경로 없음: 시트 '{sheetName}', 활성화 상태만 저장: {enabled}");
                    pathManager.SetSheetEnabled(workbookName, sheetName, enabled);
                    
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.SetSheetEnabled(fullWorkbookPath, sheetName, enabled);
                    }
                    
                    // 경로가 비어있고 활성화되지 않은 경우에만 경로 정보 삭제
                    if (!enabled)
                    {
                        Debug.WriteLine($"[UpdateSheetPathForRow] 제거: 시트 '{sheetName}' (경로가 비어있고 비활성화됨)");
                    pathManager.RemoveSheetPath(workbookName, sheetName);
                    if (workbookName != fullWorkbookPath)
                    {
                        pathManager.RemoveSheetPath(fullWorkbookPath, sheetName);
                        }
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
                string sheetName = dataGridView.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
                string currentPath = dataGridView.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                
                Debug.WriteLine($"[DataGridView_CellContentClick] 폴더 선택 버튼 클릭: 행={e.RowIndex}, 시트='{sheetName}', 현재 경로='{currentPath}'");
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
                // 현재 워크북 설정
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                
                // 워크북 경로 변수를 블록 밖에서 선언
                string workbookPath = "";
                
                if (app.ActiveWorkbook != null)
                {
                    workbookPath = app.ActiveWorkbook.FullName;
                    SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
                    
                    // Excel 설정 관리자 설정
                    if (excelConfigManager != null)
                    {
                        excelConfigManager.SetCurrentWorkbook(workbookPath);
                        excelConfigManager.EnsureConfigSheetExists();
                    }
                }
                
                // 변경사항을 저장
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    DataGridViewRow row = dataGridView.Rows[i];
                    
                    // 원래 시트 이름 가져오기
                    string sheetName = row.Cells["SheetNameColumn"].Tag as string;
                    
                    if (string.IsNullOrEmpty(sheetName))
                        continue;
                    
                    // 활성화 상태 가져오기
                    bool isEnabled = (bool)row.Cells["EnabledColumn"].Value;
                    
                    // 저장 경로 가져오기
                    string sheetPath = row.Cells["PathColumn"].Value as string;
                    
                    // 저장 경로가 없으면 비활성화
                    if (string.IsNullOrEmpty(sheetPath))
                    {
                        isEnabled = false;
                    }
                    
                    // 저장 경로는 계속 XML에 저장 (사용자별 데이터)
                    if (!string.IsNullOrEmpty(sheetPath))
                    {
                        string workbookName = Path.GetFileName(workbookPath);
                        
                        // 워크북 이름과 전체 경로 모두 저장
                        SheetPathManager.Instance.SetSheetPath(workbookName, sheetName, sheetPath);
                        SheetPathManager.Instance.SetSheetPath(workbookPath, sheetName, sheetPath);
                        
                        Debug.WriteLine($"[SaveButton_Click] 경로 저장: 워크북='{workbookPath}', 시트='{sheetName}', 경로='{sheetPath}'");
                    }
                    
                    // 활성화 상태 설정 (XML에만 저장)
                    string activeWorkbookName = Path.GetFileName(workbookPath);
                    SheetPathManager.Instance.SetSheetEnabled(activeWorkbookName, sheetName, isEnabled);
                    
                    // YAML 선택적 필드 처리 옵션 저장
                    bool yamlOption = (bool)row.Cells["YamlEmptyFields"].Value;
                    
                    // Excel에만 저장 (XML에는 저장하지 않음)
                    if (excelConfigManager != null)
                    {
                        // Excel에 저장
                        excelConfigManager.SetConfigValue(sheetName, "YamlEmptyFields", yamlOption.ToString());
                    }
                    
                    // 병합 키 경로 설정 저장
                    string mergeKeyPaths = row.Cells["MergePathsColumn"].Value as string;
                    
                    if (!string.IsNullOrEmpty(mergeKeyPaths))
                    {
                        // Excel에만 저장 (XML에는 저장하지 않음)
                        if (excelConfigManager != null)
                        {
                            excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPaths);
                        }
                    }
                    else
                    {
                        // 빈 값이면 제거
                        if (excelConfigManager != null)
                        {
                            excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", "");
                        }
                    }
                    
                    // Flow 스타일 설정 저장 (필드와 항목 필드 합쳐서 구성)
                    string flowStyleFields = row.Cells["FlowStyleFieldsColumn"].Value as string ?? "";
                    string flowStyleItemsFields = row.Cells["FlowStyleItemsFieldsColumn"].Value as string ?? "";
                    
                    // 두 설정 합치기 (필드|항목필드 형식)
                    string flowStyle = $"{flowStyleFields}|{flowStyleItemsFields}";
                    
                    if (!string.IsNullOrEmpty(flowStyle) && !flowStyle.Equals("|"))
                    {
                        // Excel에만 저장 (XML에는 저장하지 않음)
                        if (excelConfigManager != null)
                        {
                            excelConfigManager.SetConfigValue(sheetName, "FlowStyle", flowStyle);
                        }
                    }
                    else
                    {
                        // 빈 값이면 제거
                        if (excelConfigManager != null)
                        {
                            excelConfigManager.SetConfigValue(sheetName, "FlowStyle", "");
                        }
                    }
                }
                
                // XML 설정 저장
                SheetPathManager.Instance.SaveSettings();

                // 폼 닫기
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"시트별 경로 설정을 저장하는 중 오류가 발생했습니다: {ex.Message}", 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SheetPathSettingsForm_Load(object sender, EventArgs e)
        {
            string configFilePath = SheetPathManager.GetConfigFilePath();
            bool configFileExists = File.Exists(configFilePath);
            
            lblConfigPath.Text = $"설정 파일 경로: {configFilePath}";
            lblConfigPath.Text += $"\n설정 파일 존재 여부: {(configFileExists ? "있음" : "없음")}";
            
            // !Config 시트에 대한 정보 추가
            lblConfigPath.Text += $"\nExcel 내부 설정: {ExcelConfigManager.CONFIG_SHEET_NAME} 시트";
            
            // 설정 파일이 없으면 생성
            if (!configFileExists)
            {
                Debug.WriteLine($"[SheetPathSettingsForm_Load] 설정 파일이 존재하지 않습니다. 초기화 시도");
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine($"[SheetPathSettingsForm_Load] 설정 파일 초기화 완료");
            }
            
            // 초기 DataGridView 크기 조정
            AdjustDataGridViewSize();
        }
    }
}
