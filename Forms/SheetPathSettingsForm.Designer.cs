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
            this.BrowseColumn});
            this.dataGridView.Location = new System.Drawing.Point(12, 47);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(676, 323);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView_CellContentClick);
            // 
            // SheetNameColumn
            // 
            this.SheetNameColumn.HeaderText = "시트 이름";
            this.SheetNameColumn.Name = "SheetNameColumn";
            this.SheetNameColumn.ReadOnly = true;
            this.SheetNameColumn.Width = 150;
            // 
            // EnabledColumn
            // 
            this.EnabledColumn.HeaderText = "활성화";
            this.EnabledColumn.Name = "EnabledColumn";
            this.EnabledColumn.Width = 70;
            // 
            // PathColumn
            // 
            this.PathColumn.HeaderText = "저장 경로";
            this.PathColumn.Name = "PathColumn";
            this.PathColumn.ReadOnly = false;
            this.PathColumn.Width = 300;
            // 
            // BrowseColumn
            // 
            this.BrowseColumn.HeaderText = "폴더 선택";
            this.BrowseColumn.Name = "BrowseColumn";
            this.BrowseColumn.Text = "...";
            this.BrowseColumn.UseColumnTextForButtonValue = true;
            this.BrowseColumn.Width = 80;
            // 
            // saveButton
            // 
            this.saveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.saveButton.Location = new System.Drawing.Point(532, 376);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(75, 23);
            this.saveButton.TabIndex = 1;
            this.saveButton.Text = "저장";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Location = new System.Drawing.Point(613, 376);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "취소";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.TabIndex = 2;
            this.label1.Text = "각 시트별 JSON 파일 저장 경로를 설정합니다.";
            // 
            // lblConfigPath
            // 
            this.lblConfigPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblConfigPath.AutoSize = true;
            this.lblConfigPath.ForeColor = System.Drawing.SystemColors.GrayText;
            this.lblConfigPath.Location = new System.Drawing.Point(12, 382);
            this.lblConfigPath.Name = "lblConfigPath";
            this.lblConfigPath.Size = new System.Drawing.Size(85, 12);
            this.lblConfigPath.TabIndex = 3;
            this.lblConfigPath.Text = "설정 파일 경로:";
            // 
            // SheetPathSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 411);
            this.Controls.Add(this.lblConfigPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.dataGridView);
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
    }
}
