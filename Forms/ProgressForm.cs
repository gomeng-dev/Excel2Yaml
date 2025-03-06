using System;
using System.ComponentModel;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace ExcelToYamlAddin.Forms
{
    public partial class ProgressForm : Form
    {
        private BackgroundWorker worker;
        private Action<IProgress<ProgressInfo>> workAction;
        private bool cancelRequested = false;

        public class ProgressInfo
        {
            public int Percentage { get; set; }
            public string StatusMessage { get; set; }
            public bool IsCompleted { get; set; }
            public bool HasError { get; set; }
            public string ErrorMessage { get; set; }
        }

        public ProgressForm()
        {
            InitializeComponent();
            
            // 모던 스타일 적용
            ApplyModernStyle();
        }

        private void InitializeComponent()
        {
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar = new ModernProgressBar();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(12, 42);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(460, 8);
            this.progressBar.TabIndex = 0;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 18);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(99, 15);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "변환 준비 중...";
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(397, 82);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "취소";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ProgressForm
            // 
            this.ClientSize = new System.Drawing.Size(484, 117);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel 변환 작업 진행 중";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProgressForm_FormClosing);
            this.Load += new System.EventHandler(this.ProgressForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private ModernProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnCancel;
        
        // 작업 실행 및 진행 상태 업데이트를 처리하는 메서드
        public void RunOperation(Action<IProgress<ProgressInfo>> work, string title = null)
        {
            if (!string.IsNullOrEmpty(title))
                this.Text = title;
                
            workAction = work;
            
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            
            // 폼이 로드될 때 작업 시작
        }
        
        private void ProgressForm_Load(object sender, EventArgs e)
        {
            // 애니메이션 효과 적용
            this.Opacity = 0;
            
            // 타이머를 사용한 페이드인 효과
            System.Windows.Forms.Timer fadeTimer = new System.Windows.Forms.Timer();
            fadeTimer.Interval = 10;
            fadeTimer.Tick += (s, args) => {
                if (this.Opacity < 1)
                {
                    this.Opacity += 0.05;
                }
                else
                {
                    fadeTimer.Stop();
                    fadeTimer.Dispose();
                    
                    // 페이드인 완료 후 작업 시작
                    if (worker != null && !worker.IsBusy)
                    {
                        worker.RunWorkerAsync();
                    }
                }
            };
            fadeTimer.Start();
            
            // 버튼 위치 재조정 (안전성 확보)
            StyleCancelButton();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var progress = new Progress<ProgressInfo>(info => 
            {
                int percentage = Math.Min(Math.Max(info.Percentage, 0), 100);
                worker.ReportProgress(percentage, info);
            });

            try
            {
                workAction(progress);
            }
            catch (Exception ex)
            {
                var errorInfo = new ProgressInfo
                {
                    IsCompleted = true,
                    HasError = true,
                    ErrorMessage = ex.Message,
                    StatusMessage = "오류 발생: " + ex.Message
                };
                worker.ReportProgress(100, errorInfo);
                e.Result = errorInfo;
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var info = e.UserState as ProgressInfo;
            if (info != null)
            {
                progressBar.Value = e.ProgressPercentage;
                lblStatus.Text = info.StatusMessage ?? "";
                
                // 진행 상태가 100%이고 완료되었다면 폼을 닫을 준비
                if (e.ProgressPercentage >= 100 && info.IsCompleted)
                {
                    // 에러가 있으면 메시지 표시
                    if (info.HasError)
                    {
                        MessageBox.Show(info.ErrorMessage, "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    
                    // 잠시 후 폼 닫기
                    System.Windows.Forms.Timer closeTimer = new System.Windows.Forms.Timer();
                    closeTimer.Interval = 500; // 0.5초 후 닫기
                    closeTimer.Tick += (s, args) => 
                    {
                        closeTimer.Stop();
                        this.DialogResult = info.HasError ? DialogResult.Cancel : DialogResult.OK;
                        this.Close();
                    };
                    closeTimer.Start();
                }
            }
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 작업이 취소되었거나 예외가 발생한 경우
            if (e.Cancelled)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
            else if (e.Error != null)
            {
                MessageBox.Show("작업 중 오류가 발생했습니다: " + e.Error.Message, 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (worker != null && worker.IsBusy && worker.WorkerSupportsCancellation)
            {
                cancelRequested = true;
                lblStatus.Text = "작업 취소 중...";
                worker.CancelAsync();
                btnCancel.Enabled = false;
            }
        }

        private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 작업 중이라면 취소 처리
            if (worker != null && worker.IsBusy)
            {
                e.Cancel = true; // 폼 닫기 취소
                cancelRequested = true;
                lblStatus.Text = "작업 취소 중...";
                worker.CancelAsync();
                btnCancel.Enabled = false;
            }
        }
        
        // 취소 여부 확인 속성
        public bool IsCancellationRequested
        {
            get { return cancelRequested; }
        }

        private void ApplyModernStyle()
        {
            try
            {
                // 폼 스타일 설정
                this.BackColor = Color.White;
                this.Font = new Font("Segoe UI", 9F);
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = true;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.ShowIcon = true;
                
                // 프로그레스 바 스타일 설정
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Height = 8;
                progressBar.ForeColor = Color.FromArgb(0, 120, 215);  // 파란색 프로그레스 바
                
                // 라벨 스타일
                lblStatus.Font = new Font("Segoe UI", 9.5F);
                lblStatus.ForeColor = Color.FromArgb(40, 40, 40);
                lblStatus.AutoSize = false;
                lblStatus.Width = this.ClientSize.Width - 24;
                lblStatus.Height = 20;
                lblStatus.TextAlign = ContentAlignment.MiddleLeft;
                
                // 취소 버튼 스타일링
                StyleCancelButton();
                
                // 타이틀 스타일
                this.Text = "변환 작업 진행 중...";
                
                // 툴팁 추가
                ToolTip toolTip = new ToolTip();
                toolTip.SetToolTip(btnCancel, "작업을 취소하고 돌아갑니다");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ApplyModernStyle] 예외 발생: {ex.Message}");
            }
        }

        /// <summary>
        /// 취소 버튼에 모던한 스타일을 적용합니다.
        /// </summary>
        private void StyleCancelButton()
        {
            // 취소 버튼 스타일
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.BackColor = Color.FromArgb(245, 245, 245);
            btnCancel.ForeColor = Color.FromArgb(60, 60, 60);
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.Font = new Font("Segoe UI", 9F);
            btnCancel.Cursor = Cursors.Hand;
            btnCancel.Text = "취소";
            btnCancel.Size = new Size(80, 28);
            
            // 버튼 이벤트 - 마우스 오버 효과
            btnCancel.MouseEnter += (sender, e) => {
                btnCancel.BackColor = Color.FromArgb(230, 230, 230);
            };
            
            btnCancel.MouseLeave += (sender, e) => {
                btnCancel.BackColor = Color.FromArgb(245, 245, 245);
            };
            
            // 버튼 위치 조정
            btnCancel.Location = new Point(
                this.ClientSize.Width - btnCancel.Width - 12,
                this.ClientSize.Height - btnCancel.Height - 12
            );
        }

        // 프로그레스 바 렌더링 재정의를 위한 클래스 추가
        public class ModernProgressBar : ProgressBar
        {
            public ModernProgressBar()
            {
                this.SetStyle(ControlStyles.UserPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                Rectangle rect = new Rectangle(0, 0, this.Width, this.Height);
                e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(240, 240, 240)), rect);

                if (this.Value > 0)
                {
                    Rectangle progressRect = new Rectangle(0, 0, (int)((float)this.Value / this.Maximum * this.Width), this.Height);
                    LinearGradientBrush brush = new LinearGradientBrush(
                        progressRect,
                        Color.FromArgb(0, 120, 215),
                        Color.FromArgb(0, 130, 230),
                        LinearGradientMode.Vertical);
                    e.Graphics.FillRectangle(brush, progressRect);
                }
            }
        }
    }
} 