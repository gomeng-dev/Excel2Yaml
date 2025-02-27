namespace ExcelToJsonAddin.Forms
{
    partial class SheetPathSettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.SheetNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EnabledColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.PathColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BrowseColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.YamlEmptyFields = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.IdPathColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MergePathsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KeyPathsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FlowStyleFieldsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FlowStyleItemsFieldsColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.saveButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblConfigPath = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SheetNameColumn,
            this.EnabledColumn,
            this.PathColumn,
            this.BrowseColumn,
            this.YamlEmptyFields,
            this.IdPathColumn,
            this.MergePathsColumn,
            this.KeyPathsColumn,
            this.FlowStyleFieldsColumn,
            this.FlowStyleItemsFieldsColumn});
            this.dataGridView.Location = new System.Drawing.Point(17, 70);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 62;
            this.dataGridView.Size = new System.Drawing.Size(1295, 629);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView_CellContentClick);
            // 
            // SheetNameColumn
            // 
            this.SheetNameColumn.HeaderText = "시트 이름";
            this.SheetNameColumn.MinimumWidth = 8;
            this.SheetNameColumn.Name = "SheetNameColumn";
            this.SheetNameColumn.ReadOnly = true;
            this.SheetNameColumn.Width = 150;
            // 
            // EnabledColumn
            // 
            this.EnabledColumn.HeaderText = "활성화";
            this.EnabledColumn.MinimumWidth = 8;
            this.EnabledColumn.Name = "EnabledColumn";
            this.EnabledColumn.Width = 70;
            // 
            // PathColumn
            // 
            this.PathColumn.HeaderText = "저장 경로";
            this.PathColumn.MinimumWidth = 8;
            this.PathColumn.Name = "PathColumn";
            this.PathColumn.Width = 300;
            // 
            // BrowseColumn
            // 
            this.BrowseColumn.HeaderText = "폴더 선택";
            this.BrowseColumn.MinimumWidth = 8;
            this.BrowseColumn.Name = "BrowseColumn";
            this.BrowseColumn.Text = "...";
            this.BrowseColumn.UseColumnTextForButtonValue = true;
            this.BrowseColumn.Width = 80;
            // 
            // YamlEmptyFields
            // 
            this.YamlEmptyFields.HeaderText = "YAML 선택적 필드";
            this.YamlEmptyFields.MinimumWidth = 8;
            this.YamlEmptyFields.Name = "YamlEmptyFields";
            this.YamlEmptyFields.ToolTipText = "YAML 파일에 선택적 필드를 빈 값으로 포함할지 여부";
            this.YamlEmptyFields.Width = 120;
            // 
            // IdPathColumn
            // 
            this.IdPathColumn.HeaderText = "ID 경로";
            this.IdPathColumn.MinimumWidth = 8;
            this.IdPathColumn.Name = "IdPathColumn";
            this.IdPathColumn.ToolTipText = "ID가 있는 경로 (기본값: id)";
            this.IdPathColumn.Width = 90;
            // 
            // MergePathsColumn
            // 
            this.MergePathsColumn.HeaderText = "병합 경로";
            this.MergePathsColumn.MinimumWidth = 8;
            this.MergePathsColumn.Name = "MergePathsColumn";
            this.MergePathsColumn.ToolTipText = "병합할 경로들 (기본값: events)";
            this.MergePathsColumn.Width = 90;
            // 
            // KeyPathsColumn
            // 
            this.KeyPathsColumn.HeaderText = "키 경로 전략";
            this.KeyPathsColumn.MinimumWidth = 8;
            this.KeyPathsColumn.Name = "KeyPathsColumn";
            this.KeyPathsColumn.ToolTipText = "키 경로:전략 문자열 (예: level:merge;achievement:append)";
            this.KeyPathsColumn.Width = 120;
            // 
            // FlowStyleFieldsColumn
            // 
            this.FlowStyleFieldsColumn.HeaderText = "Flow 필드";
            this.FlowStyleFieldsColumn.MinimumWidth = 8;
            this.FlowStyleFieldsColumn.Name = "FlowStyleFieldsColumn";
            this.FlowStyleFieldsColumn.ToolTipText = "Flow 스타일로 변환할 필드 (예: details,info)";
            // 
            // FlowStyleItemsFieldsColumn
            // 
            this.FlowStyleItemsFieldsColumn.HeaderText = "Flow 항목 필드";
            this.FlowStyleItemsFieldsColumn.MinimumWidth = 8;
            this.FlowStyleItemsFieldsColumn.Name = "FlowStyleItemsFieldsColumn";
            this.FlowStyleItemsFieldsColumn.ToolTipText = "Flow 스타일로 변환할 항목 필드 (예: triggers,events)";
            this.FlowStyleItemsFieldsColumn.Width = 110;
            // 
            // saveButton
            // 
            this.saveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveButton.Location = new System.Drawing.Point(1090, 733);
            this.saveButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(107, 34);
            this.saveButton.TabIndex = 1;
            this.saveButton.Text = "저장";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Location = new System.Drawing.Point(1205, 733);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(107, 34);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "취소";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(382, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "각 시트별 JSON 파일 저장 경로를 설정합니다.";
            // 
            // lblConfigPath
            // 
            this.lblConfigPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblConfigPath.AutoSize = true;
            this.lblConfigPath.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblConfigPath.Location = new System.Drawing.Point(17, 703);
            this.lblConfigPath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblConfigPath.Name = "lblConfigPath";
            this.lblConfigPath.Size = new System.Drawing.Size(134, 18);
            this.lblConfigPath.TabIndex = 3;
            this.lblConfigPath.Text = "설정 파일 경로:";
            // 
            // SheetPathSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1325, 780);
            this.Controls.Add(this.lblConfigPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.dataGridView);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SheetPathSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "시트별 경로 설정";
            this.Load += new System.EventHandler(this.SheetPathSettingsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn SheetNameColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn EnabledColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn PathColumn;
        private System.Windows.Forms.DataGridViewButtonColumn BrowseColumn;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblConfigPath;
        private System.Windows.Forms.DataGridViewCheckBoxColumn YamlEmptyFields;
        private System.Windows.Forms.DataGridViewTextBoxColumn IdPathColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn MergePathsColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn KeyPathsColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FlowStyleFieldsColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn FlowStyleItemsFieldsColumn;
    }
}
