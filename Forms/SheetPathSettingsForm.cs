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
using ExcelToYamlAddin.Core;
using ExcelToYamlAddin.Config;

namespace ExcelToYamlAddin.Forms
{
    public partial class SheetPathSettingsForm : Form
    {
        private List<Worksheet> convertibleSheets;
        
        // Excel 설정 관리자 추가
        private ExcelConfigManager excelConfigManager;

        public SheetPathSettingsForm(List<Worksheet> sheets)
        {
            InitializeComponent();
            
            Debug.WriteLine("[SheetPathSettingsForm] 생성자 시작");
            
            // 데이터 그리드 셀 포맷 설정
            ConfigureDataGridCellFormatting();
            
            // 시트 목록 저장
            convertibleSheets = sheets;
            
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
                
                // 디버깅: excel2yamlconfig 시트에서 로드한 값 확인
                Debug.WriteLine("[SheetPathSettingsForm] excel2yamlconfig 시트에서 로드한 값 확인");
                foreach (var sheet in sheets)
                {
                    string sheetName = sheet.Name;
                    
                    // 후처리 관련 값 로그
                    string mergeKeyPaths = excelConfigManager.GetConfigValue(sheetName, "MergeKeyPaths", "");
                    string flowStyle = excelConfigManager.GetConfigValue(sheetName, "FlowStyle", "");
                    bool yamlEmptyFields = excelConfigManager.GetConfigBool(sheetName, "YamlEmptyFields", false);
                    
                    Debug.WriteLine($"[SheetPathSettingsForm] 시트: '{sheetName}'");
                    Debug.WriteLine($"[SheetPathSettingsForm]   - MergeKeyPaths: '{mergeKeyPaths}'");
                    Debug.WriteLine($"[SheetPathSettingsForm]   - FlowStyle: '{flowStyle}'");
                    Debug.WriteLine($"[SheetPathSettingsForm]   - YamlEmptyFields: {yamlEmptyFields}");
                }
            }
            
            Debug.WriteLine("[SheetPathSettingsForm] 생성자 완료");
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
                // 폼 위와 아래의 여백 계산
                int verticalMargin = dataGridView.Top + (this.ClientSize.Height - dataGridView.Bottom);
                
                // 좌우 여백
                int horizontalMargin = 40;
                
                // 데이터그리드뷰의 너비를 폼에 맞게 조정
                dataGridView.Width = this.ClientSize.Width - horizontalMargin;
                
                // 모든 컬럼의 내용에 맞게 너비 자동 조정
                dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                
                // 컬럼 내용이 모두 보이도록 최소 너비 설정
                int totalColumnsWidth = 0;
                foreach (DataGridViewColumn col in dataGridView.Columns)
                {
                    // 최소 80픽셀, 최대 300픽셀로 제한
                    col.Width = Math.Max(80, Math.Min(300, col.Width));
                    totalColumnsWidth += col.Width;
                }
                
                // 폼의 크기를 모든 컬럼이 보이도록 조정
                int newFormWidth = totalColumnsWidth + horizontalMargin + 50; // 스크롤바 공간 추가
                
                // 현재 폼 크기보다 큰 경우에만 조정
                if (newFormWidth > this.Width)
                {
                    this.Width = newFormWidth;
                }
                
                Debug.WriteLine($"[AdjustDataGridViewSize] DataGridView 크기 조정: 너비={dataGridView.Width}, 폼너비={this.ClientSize.Width}, 총컬럼너비={totalColumnsWidth}");
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
            try
            {
                Debug.WriteLine("[LoadSheetPaths] 시트 경로 정보 로드 시작");
                // 시트 경로 정보 로드
                PopulateSheetsList();
                
                // 데이터 그리드 셀 서식 설정
                ConfigureDataGridCellFormatting();
                
                // 그리드 이벤트 핸들러 추가
                dataGridView.CellValueChanged += new DataGridViewCellEventHandler(DataGridView_CellValueChanged);
                dataGridView.CellEndEdit += new DataGridViewCellEventHandler(DataGridView_CellEndEdit);
                dataGridView.KeyDown += new KeyEventHandler(DataGridView_KeyDown);
                
                Debug.WriteLine("[LoadSheetPaths] 시트 경로 정보 로드 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadSheetPaths] 예외 발생: {ex.Message}\n{ex.StackTrace}");
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
                    
                    // 표시용 시트 이름 (UI에 표시할 때는 '!'를 제거)
                    string displayName = sheetName;
                    if (displayName.StartsWith("!"))
                    {
                        displayName = displayName.Substring(1);
                    }

                    // XML에서 설정 로드 (원본 시트 이름 사용)
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
                    string arrayFieldPaths = "";
                        
                    if (!string.IsNullOrEmpty(mergeKeyPaths))
                    {
                        string[] parts = mergeKeyPaths.Split('|');
                        if (parts.Length >= 1) idPath = parts[0];
                        if (parts.Length >= 2) mergePaths = parts[1];
                        if (parts.Length >= 3) keyPaths = parts[2];
                        if (parts.Length >= 4) arrayFieldPaths = parts[3];
                    }
                    
                    Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}'의 병합 키 경로 분리: idPath='{idPath}', mergePaths='{mergePaths}', keyPaths='{keyPaths}', arrayFieldPaths='{arrayFieldPaths}'");
                    
                    // Flow Style 설정 분리 (Flow Style 필드 | Flow Style 항목 필드)
                    string flowStyleFields = "";
                    string flowStyleItemsFields = "";
                    
                    if (!string.IsNullOrEmpty(flowStyleConfig))
                    {
                        string[] parts = flowStyleConfig.Split('|');
                        if (parts.Length >= 1) flowStyleFields = parts[0];
                        if (parts.Length >= 2) flowStyleItemsFields = parts[1];
                    }
                    
                    Debug.WriteLine($"[PopulateSheetsList] 시트 '{sheetName}'의 Flow Style 분리: fields='{flowStyleFields}', itemsFields='{flowStyleItemsFields}'");
                    
                    // 행 추가 준비
                    int rowIndex = dataGridView.Rows.Add();
                    var row = dataGridView.Rows[rowIndex];
                        
                    // 주요 열 채우기
                    row.Cells[0].Value = displayName;                // 시트 이름
                    row.Cells[1].Value = isEnabled;                // 활성화 여부
                    row.Cells[2].Value = savePath;                 // 저장 경로
                    row.Cells[4].Value = yamlOption;               // YAML 선택적 필드 처리
                    
                    // 원본 시트 이름 태그로 저장 (느낌표 포함된 원래 이름)
                    row.Cells[0].Tag = sheet.Name;
                    
                    // 추가 열 채우기 (병합 키 경로, 플로우 스타일)
                    if (showExtraColumns)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.OwningColumn.Name == "IdPathColumn")
                            {
                                cell.Value = idPath;
                            }
                            else if (cell.OwningColumn.Name == "MergePathsColumn")
                            {
                                cell.Value = mergePaths;
                            }
                            else if (cell.OwningColumn.Name == "KeyPathsColumn")
                            {
                                cell.Value = keyPaths;
                            }
                            else if (cell.OwningColumn.Name == "FlowStyleFieldsColumn")
                            {
                                cell.Value = flowStyleFields;
                            }
                            else if (cell.OwningColumn.Name == "FlowStyleItemsFieldsColumn")
                            {
                                cell.Value = flowStyleItemsFields;
                            }
                            else if (cell.OwningColumn.Name == "ArrayFieldPathsColumn")
                            {
                                cell.Value = arrayFieldPaths;
                            }
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

        /// <summary>
        /// DataGridView_CellValueChanged 이벤트 핸들러
        /// </summary>
        private void DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            try
            {
                string sheetName = dataGridView.Rows[e.RowIndex].Cells[0].Tag?.ToString() ?? "";
                
                if (string.IsNullOrEmpty(sheetName))
                {
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 행 {e.RowIndex}의 시트 이름이 없습니다.");
                    return;
                }
                
                Debug.WriteLine($"[DataGridView_CellValueChanged] 행 {e.RowIndex}, 시트 '{sheetName}'의 셀 변경 처리 중...");
                
                // "활성화" 열 처리 (인덱스 1)
                if (e.ColumnIndex == 1)
                {
                    bool isEnabled = Convert.ToBoolean(dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 활성화 상태 변경: {isEnabled}");
                    
                    // SheetPathManager에만 설정을 저장합니다.
                    SheetPathManager.Instance.SetSheetEnabled(sheetName, isEnabled);
                }
                // "경로" 열 처리 (인덱스 2)
                else if (e.ColumnIndex == 2)
                {
                    string path = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 경로 변경: '{path}'");
                    
                    UpdateSheetPathForRow(e.RowIndex);
                }
                // YAML 빈 필드 포함 옵션 처리 (인덱스 4)
                else if (e.ColumnIndex == 4)
                {
                    bool yamlEmptyFields = Convert.ToBoolean(dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 YAML 선택적 필드 설정 변경: {yamlEmptyFields}");
                    
                    // SheetPathManager에 저장하지 않고 ExcelConfigManager에만 저장
                    // SheetPathManager.Instance.SetYamlEmptyFieldsOption(sheetName, yamlEmptyFields);
                    
                    // XML 파일에만 저장하므로 ExcelConfigManager 저장 부분은 제거합니다.
                    // if (excelConfigManager != null)
                    // {
                    //     excelConfigManager.SetConfigValue(sheetName, "YamlEmptyFields", yamlEmptyFields.ToString().ToLower());
                    // }
                }
                // ID 경로 열 처리
                else if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "IdPathColumn")
                {
                    string idPath = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 ID 경로 설정 변경: '{idPath}'");
                    
                    // 설정 값 변경 확인
                    var cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    string cellValue = cell.Value?.ToString() ?? "";
                    
                    if (cellValue != idPath)
                    {
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 경고: 셀 값이 의도한 값과 다릅니다. 설정하려는 값: '{idPath}', 실제 셀 값: '{cellValue}'");
                        cell.Value = idPath;
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 셀 값을 수정했습니다. 새 값: '{idPath}'");
                    }
                    
                    // 전체 MergeKeyPaths 문자열 업데이트 (중복 저장 방지)
                    if (excelConfigManager != null)
                    {
                        // 다른 관련 값 가져오기
                        string mergePaths = dataGridView.Rows[e.RowIndex].Cells["MergePathsColumn"].Value?.ToString() ?? "";
                        string keyPaths = dataGridView.Rows[e.RowIndex].Cells["KeyPathsColumn"].Value?.ToString() ?? "";
                        string arrayFieldPaths = dataGridView.Rows[e.RowIndex].Cells["ArrayFieldPathsColumn"].Value?.ToString() ?? "";
                        
                        // MergeKeyPaths 통합 문자열 업데이트
                        string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                        excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                        
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 MergeKeyPaths 통합 설정 업데이트: '{mergeKeyPathsConfig}'");
                    }
                }
                // 병합 경로 필드 변경 처리 (인덱스 6)
                else if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "MergePathsColumn")
                {
                    string mergePaths = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 병합 경로 변경: '{mergePaths}'");
                    
                    // 값이 제대로 설정되었는지 확인
                    var cell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    string cellValue = cell.Value?.ToString() ?? "";
                    if (cellValue != mergePaths)
                    {
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 경고: 셀 값이 의도한 값과 다릅니다. 설정하려는 값: '{mergePaths}', 실제 셀 값: '{cellValue}'");
                        cell.Value = mergePaths;
                    }
                    
                    // 전체 MergeKeyPaths 문자열 업데이트 (중복 저장 방지)
                    if (excelConfigManager != null)
                    {
                        // 다른 관련 값 가져오기
                        string idPath = dataGridView.Rows[e.RowIndex].Cells["IdPathColumn"].Value?.ToString() ?? "id";
                        string keyPaths = dataGridView.Rows[e.RowIndex].Cells["KeyPathsColumn"].Value?.ToString() ?? "";
                        string arrayFieldPaths = dataGridView.Rows[e.RowIndex].Cells["ArrayFieldPathsColumn"].Value?.ToString() ?? "";
                        
                        // MergeKeyPaths 통합 문자열 업데이트
                        string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                        excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                        
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 MergeKeyPaths 통합 설정 업데이트: '{mergeKeyPathsConfig}'");
                    }
                }
                // 키 경로 필드 변경 처리 (인덱스 7)
                else if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "KeyPathsColumn")
                {
                    string keyPaths = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 키 경로 변경: '{keyPaths}'");
                    
                    // 전체 MergeKeyPaths 문자열 업데이트 (중복 저장 방지)
                    if (excelConfigManager != null)
                    {
                        // 다른 관련 값 가져오기
                        string idPath = dataGridView.Rows[e.RowIndex].Cells["IdPathColumn"].Value?.ToString() ?? "id";
                        string mergePaths = dataGridView.Rows[e.RowIndex].Cells["MergePathsColumn"].Value?.ToString() ?? "";
                        string arrayFieldPaths = dataGridView.Rows[e.RowIndex].Cells["ArrayFieldPathsColumn"].Value?.ToString() ?? "";
                        
                        // MergeKeyPaths 통합 문자열 업데이트
                        string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                        excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                        
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 MergeKeyPaths 통합 설정 업데이트: '{mergeKeyPathsConfig}'");
                    }
                }
                // 배열 필드 경로 필드 변경 처리
                else if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "ArrayFieldPathsColumn")
                {
                    string arrayFieldPaths = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 배열 필드 경로 변경: '{arrayFieldPaths}'");
                    
                    // 전체 MergeKeyPaths 문자열 업데이트 (중복 저장 방지)
                    if (excelConfigManager != null)
                    {
                        // 다른 관련 값 가져오기
                        string idPath = dataGridView.Rows[e.RowIndex].Cells["IdPathColumn"].Value?.ToString() ?? "id";
                        string mergePaths = dataGridView.Rows[e.RowIndex].Cells["MergePathsColumn"].Value?.ToString() ?? "";
                        string keyPaths = dataGridView.Rows[e.RowIndex].Cells["KeyPathsColumn"].Value?.ToString() ?? "";
                        
                        // MergeKeyPaths 통합 문자열 업데이트
                        string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                        excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                        
                        Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 MergeKeyPaths 통합 설정 업데이트: '{mergeKeyPathsConfig}'");
                    }
                }
                // Flow Style 필드 설정 필드 변경 처리
                else if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "FlowStyleFieldsColumn")
                {
                    UpdateSheetPathForRow(e.RowIndex);
                    string flowStyleFields = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 Flow Style 필드 설정 변경: '{flowStyleFields}'");
                }
                
                // Flow Style 항목 필드 설정 필드 변경 처리
                if (e.ColumnIndex >= 0 && dataGridView.Columns[e.ColumnIndex].Name == "FlowStyleItemsFieldsColumn")
                {
                    UpdateSheetPathForRow(e.RowIndex);
                    string flowStyleItemsFields = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Debug.WriteLine($"[DataGridView_CellValueChanged] 시트 '{sheetName}'의 Flow Style 항목 필드 설정 변경: '{flowStyleItemsFields}'");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DataGridView_CellValueChanged] 예외 발생: {ex.Message}");
                Debug.WriteLine($"[DataGridView_CellValueChanged] 스택 추적: {ex.StackTrace}");
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
            if (rowIndex < 0 || rowIndex >= dataGridView.Rows.Count)
                return;
                
            var row = dataGridView.Rows[rowIndex];
            // 원본 시트 이름(느낌표 포함)을 Tag에서 가져옴
            string sheetName = row.Cells["SheetNameColumn"].Tag?.ToString() ?? "";
            string path = row.Cells["PathColumn"].Value?.ToString() ?? "";
            bool enabled = row.Cells["EnabledColumn"].Value != null && (bool)row.Cells["EnabledColumn"].Value;
            bool yamlEmptyFields = row.Cells["YamlEmptyFields"].Value != null && (bool)row.Cells["YamlEmptyFields"].Value;
            string idPath = "";
            string mergePaths = "";
            string keyPaths = "";
            string arrayFieldPaths = "";
            string flowStyleFields = "";
            string flowStyleItemsFields = "";
            
            // 각 설정 값 추출
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.OwningColumn.Name == "IdPathColumn")
                {
                    idPath = cell.Value?.ToString() ?? "";
                }
                else if (cell.OwningColumn.Name == "MergePathsColumn")
                {
                    mergePaths = cell.Value?.ToString() ?? "";
                }
                else if (cell.OwningColumn.Name == "KeyPathsColumn")
                {
                    keyPaths = cell.Value?.ToString() ?? "";
                }
                else if (cell.OwningColumn.Name == "ArrayFieldPathsColumn")
                {
                    arrayFieldPaths = cell.Value?.ToString() ?? "";
                }
                else if (cell.OwningColumn.Name == "FlowStyleFieldsColumn")
                {
                    flowStyleFields = cell.Value?.ToString() ?? "";
                }
                else if (cell.OwningColumn.Name == "FlowStyleItemsFieldsColumn")
                {
                    flowStyleItemsFields = cell.Value?.ToString() ?? "";
                }
            }
            
            // Flow Style 구성 (필드와 항목 합쳐서 구성)
            string flowStyle = !string.IsNullOrEmpty(flowStyleFields) || !string.IsNullOrEmpty(flowStyleItemsFields)
                ? $"{flowStyleFields}|{flowStyleItemsFields}"
                : "";
                
            Debug.WriteLine($"[UpdateSheetPathForRow] 시트 '{sheetName}' 저장: 경로='{path}', 활성화={enabled}, YAML 빈 필드={yamlEmptyFields}");
            Debug.WriteLine($"[UpdateSheetPathForRow] ID 경로='{idPath}', 병합 경로='{mergePaths}', 키 경로='{keyPaths}', 배열 필드 경로='{arrayFieldPaths}'");
            Debug.WriteLine($"[UpdateSheetPathForRow] Flow Style 구성: fields='{flowStyleFields}', itemsFields='{flowStyleItemsFields}'");
            Debug.WriteLine($"[UpdateSheetPathForRow] 최종 flowStyleConfig='{flowStyle}'");

            // Excel 설정에 YAML 관련 설정 저장
            if (excelConfigManager != null)
            {
                // 아래 세 개의 개별 설정 저장 코드를 제거 (MergeKeyPaths만 저장하도록 수정)
                // // ID 경로 설정
                // if (!string.IsNullOrEmpty(idPath))
                // {
                //     excelConfigManager.SetConfigValue(sheetName, "IdPath", idPath);
                // }
                
                // // 병합 경로 설정
                // if (!string.IsNullOrEmpty(mergePaths))
                // {
                //     excelConfigManager.SetConfigValue(sheetName, "MergePaths", mergePaths);
                // }
                
                // // 키 경로 설정
                // if (!string.IsNullOrEmpty(keyPaths))
                // {
                //     excelConfigManager.SetConfigValue(sheetName, "KeyPaths", keyPaths);
                // }
                
                // Flow Style 설정
                excelConfigManager.SetConfigValue(sheetName, "FlowStyle", flowStyle);
                
                // YAML 빈 필드 포함 설정
                excelConfigManager.SetConfigValue(sheetName, "YamlEmptyFields", yamlEmptyFields.ToString().ToLower());
                
                // MergeKeyPaths 설정 추가 - 개별 설정 대신 통합 설정만 저장
                if (!string.IsNullOrEmpty(idPath) || !string.IsNullOrEmpty(mergePaths) || !string.IsNullOrEmpty(keyPaths) || !string.IsNullOrEmpty(arrayFieldPaths))
                {
                    string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                    excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                }
            }
            
            // XML 설정에 저장 (SheetPathManager)
            if (!string.IsNullOrEmpty(path))
            {
                // 현재 워크북 경로 가져오기
                string workbookPath = excelConfigManager != null ? excelConfigManager.WorkbookPath : "";
                if (!string.IsNullOrEmpty(workbookPath))
                {
                    SheetPathManager.Instance.SetSheetPath(workbookPath, sheetName, path);
                }
                else
                {
                    Debug.WriteLine($"[UpdateSheetPathForRow] 워크북 경로가 없어 시트 '{sheetName}'의 경로를 저장할 수 없습니다.");
                }
            }
            
            SheetPathManager.Instance.SetSheetEnabled(sheetName, enabled);
            SheetPathManager.Instance.SetYamlEmptyFieldsOption(sheetName, yamlEmptyFields);
        }

        /// <summary>
        /// DataGridView 셀 내용 클릭 이벤트 핸들러
        /// </summary>
        private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // "찾아보기" 버튼 열의 인덱스 확인
                if (e.ColumnIndex == dataGridView.Columns["BrowseColumn"].Index && e.RowIndex >= 0)
                {
                    // 선택한 행에 대한 경로 선택 다이얼로그 표시
                    SelectPath(e.RowIndex);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DataGridView_CellContentClick] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 취소 버튼 클릭 이벤트 핸들러
        /// </summary>
        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void OpenFolderSelectionDialog(int rowIndex)
        {
            try
            {
                if (rowIndex < 0 || rowIndex >= dataGridView.Rows.Count)
                    return;

                SelectPath(rowIndex);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OpenFolderSelectionDialog] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 폼 로드 이벤트 핸들러
        /// </summary>
        private void SheetPathSettingsForm_Load(object sender, EventArgs e)
        {
            try
            {
                Debug.WriteLine("[SheetPathSettingsForm_Load] 시작됨");
                
                // 워크북 경로 표시
                string workbookPath = ExcelConfigManager.Instance.WorkbookPath;
                lblConfigPath.Text = workbookPath;
                Debug.WriteLine($"[SheetPathSettingsForm_Load] 워크북 경로: {workbookPath}");
                
                // 시트 경로 로드
                LoadSheetPaths();
                
                // DataGridView 컬럼 자동 조정 (콘텐츠에 맞게)
                dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                
                // DataGridView 크기 조정
                AdjustDataGridViewSize();
                
                // 설정 로깅 (디버깅용)
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    string sheetName = row.Cells[0].Value?.ToString() ?? "";
                    
                    // 설정값 로깅
                    string mergeKeyPaths = excelConfigManager.GetConfigValue(sheetName, "MergeKeyPaths", "");
                    string flowStyle = excelConfigManager.GetConfigValue(sheetName, "FlowStyle", "");
                    
                    Debug.WriteLine($"[SheetPathSettingsForm_Load] 시트: {sheetName}, MergeKeyPaths: {mergeKeyPaths}, " +
                                    $"FlowStyle: {flowStyle}");
                }
                
                Debug.WriteLine("[SheetPathSettingsForm_Load] 완료됨");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SheetPathSettingsForm_Load] 예외 발생: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"설정 로드 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 데이터 그리드 셀의 포맷팅 설정
        /// </summary>
        private void ConfigureDataGridCellFormatting()
        {
            try
            {
                // 데이터 그리드 뷰의 CellFormatting 이벤트 핸들러 추가
                dataGridView.CellFormatting += (sender, e) => {
                    if (e.RowIndex < 0 || e.ColumnIndex < 0)
                        return;
                        
                    // 셀 값 로깅 (필요시)
                    if (Debug.Listeners.Count > 0 && e.Value != null)
                    {
                        string columnName = dataGridView.Columns[e.ColumnIndex].Name;
                        string value = e.Value.ToString();
                        Debug.WriteLine($"[CellFormatting] 열 '{columnName}' 값: '{value}'");
                    }
                };
                
                // 셀 스타일 설정
                dataGridView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                
                Debug.WriteLine("[ConfigureDataGridCellFormatting] 데이터 그리드 셀 포맷 설정 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ConfigureDataGridCellFormatting] 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // SaveButton_Click 함수
        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                Debug.WriteLine("[SaveButton_Click] 시작");
                
                // 변경사항을 저장
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    var currentRow = dataGridView.Rows[i];
                    // 원본 시트 이름(느낌표 포함)을 Tag에서 가져옴
                    string sheetName = currentRow.Cells["SheetNameColumn"].Tag?.ToString() ?? "";
                    bool enabled = Convert.ToBoolean(currentRow.Cells["EnabledColumn"].Value);
                    string path = currentRow.Cells["PathColumn"].Value?.ToString() ?? "";
                    bool yamlEmptyFields = Convert.ToBoolean(currentRow.Cells["YamlEmptyFields"].Value);
                    string idPath = currentRow.Cells["IdPathColumn"].Value?.ToString() ?? "";
                    string mergePaths = currentRow.Cells["MergePathsColumn"].Value?.ToString() ?? "";
                    string keyPaths = currentRow.Cells["KeyPathsColumn"].Value?.ToString() ?? "";
                    string arrayFieldPaths = currentRow.Cells["ArrayFieldPathsColumn"].Value?.ToString() ?? "";
                    
                    // Flow 스타일 설정 저장 (필드와 항목 합쳐서 구성)
                    string flowStyleFields = currentRow.Cells["FlowStyleFieldsColumn"].Value?.ToString() ?? "";
                    string flowStyleItemsFields = currentRow.Cells["FlowStyleItemsFieldsColumn"].Value?.ToString() ?? "";
                    
                    Debug.WriteLine($"[SaveButton_Click] 시트: {sheetName}, 경로: {path}, YAML 빈 필드: {yamlEmptyFields}, " +
                                   $"ID 경로: {idPath}, 병합 경로: {mergePaths}, 키 경로: {keyPaths}, " +
                                   $"Flow 스타일 필드: {flowStyleFields}, Flow 스타일 항목 필드: {flowStyleItemsFields}");
                    
                    // 시트 정보 저장
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        // 경로 저장
                        if (!string.IsNullOrEmpty(path))
                        {
                            // 현재 워크북 경로 가져오기
                            string workbookPath = excelConfigManager != null ? excelConfigManager.WorkbookPath : "";
                            if (!string.IsNullOrEmpty(workbookPath))
                            {
                                SheetPathManager.Instance.SetSheetPath(workbookPath, sheetName, path);
                            }
                            else
                            {
                                Debug.WriteLine($"[SaveButton_Click] 워크북 경로가 없어 시트 '{sheetName}'의 경로를 저장할 수 없습니다.");
                            }
                        }
                        
                        // 활성화 상태 저장
                        SheetPathManager.Instance.SetSheetEnabled(sheetName, enabled);
                        
                        // YAML 선택적 필드 설정 - XML에는 저장하지 않음
                        // SheetPathManager.Instance.SetYamlEmptyFieldsOption(sheetName, yamlEmptyFields);
                        
                        // Excel 설정 저장
                        if (excelConfigManager != null)
                        {
                            // 개별 필드 저장 제거 - IdPath, MergePaths, KeyPaths를 개별적으로 저장하지 않음
                            
                            // Flow 스타일 설정 저장
                            string flowStyle = $"{flowStyleFields}|{flowStyleItemsFields}";
                            
                            // 병합 키 경로 저장
                            if (!string.IsNullOrEmpty(idPath) || !string.IsNullOrEmpty(mergePaths) || !string.IsNullOrEmpty(keyPaths) || !string.IsNullOrEmpty(arrayFieldPaths))
                            {
                                Debug.WriteLine($"[SaveButton_Click] 시트 '{sheetName}'의 병합 키 경로 저장: idPath='{idPath}', mergePaths='{mergePaths}', keyPaths='{keyPaths}', arrayFieldPaths='{arrayFieldPaths}'");
                                
                                string mergeKeyPathsConfig = $"{idPath}|{mergePaths}|{keyPaths}|{arrayFieldPaths}";
                                excelConfigManager.SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPathsConfig);
                            }
                            
                            excelConfigManager.SetConfigValue(sheetName, "FlowStyle", flowStyle);
                            // 활성화 상태와 YAML 빈 필드 옵션은 XML 파일에만 저장
                            excelConfigManager.SetConfigValue(sheetName, "YamlEmptyFields", yamlEmptyFields.ToString().ToLower());
                        }
                    }
                }
                
                // 설정 저장
                SheetPathManager.Instance.SaveSettings();
                
                Debug.WriteLine("[SaveButton_Click] 완료");
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SaveButton_Click] 예외 발생: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

