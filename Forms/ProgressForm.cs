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
        private Action<IProgress<ProgressInfo>, CancellationToken> workAction;
        private bool cancelRequested = false;
        private bool isClosing = false; // 폼 닫기 중복 방지를 위한 플래그
        private bool operationCompleted = false; // 작업 완료 플래그
        private CancellationTokenSource cancellationTokenSource;
        private readonly object lockObject = new object(); // 스레드 동기화를 위한 락 객체
        private IProgress<ProgressInfo> progressReporter; // 전역 Progress 객체

        // 작업 취소 상태 확인 속성 추가 (Ribbon.cs에서 호출)
        public bool IsCancellationRequested 
        { 
            get 
            { 
                // 취소 요청 상태 또는 폼 닫기 상태만 확인 (작업 완료는 취소로 간주하지 않음)
                return cancelRequested || (isClosing && !operationCompleted); 
            } 
        }

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
        public void RunOperation(Action<IProgress<ProgressInfo>, CancellationToken> work, string title = null)
        {
            if (!string.IsNullOrEmpty(title))
                this.Text = title;
                
            workAction = work;
            
            // 취소 토큰 초기화
            cancellationTokenSource = new CancellationTokenSource();
            
            // 전역 Progress 객체 초기화 (ReportProgress를 Invoke로 처리)
            progressReporter = new Progress<ProgressInfo>(info => 
            {
                try 
                {
                    if (this.InvokeRequired)
                    {
                        // UI 스레드에서 진행 상태 업데이트 (동기 호출)
                        this.Invoke(new Action(() => UpdateProgress(info)));
                    }
                    else
                    {
                        UpdateProgress(info);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[progressReporter] 예외: {ex.Message}");
                }
            });
            
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = false; // 직접 ReportProgress를 사용하지 않음
            worker.WorkerSupportsCancellation = true;
            
            worker.DoWork += Worker_DoWork;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            
            // 폼이 로드될 때 작업 시작
        }

        // 이전 버전과의 호환성을 위한 오버로드
        public void RunOperation(Action<IProgress<ProgressInfo>> work, string title = null)
        {
            RunOperation((progress, token) => 
            {
                // 주기적으로 취소 토큰을 확인하는 래퍼 구현
                work(progress);
            }, title);
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

        // 진행 상태 업데이트 분리 (UI 스레드에서 실행)
        private void UpdateProgress(ProgressInfo info)
        {
            if (this.IsDisposed || !this.IsHandleCreated || isClosing || cancelRequested || operationCompleted)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateProgress] 업데이트 무시: IsDisposed={this.IsDisposed}, !IsHandleCreated={!this.IsHandleCreated}, isClosing={isClosing}, cancelRequested={cancelRequested}, operationCompleted={operationCompleted}");
                return;
            }
                
            try
            {
                if (info != null)
                {
                    progressBar.Value = Math.Min(Math.Max(info.Percentage, 0), 100);
                    lblStatus.Text = info.StatusMessage ?? "";
                    
                    // 진행 상태가 100%이고 완료되었다면
                    if (info.Percentage >= 100 && info.IsCompleted && !operationCompleted)
                    {
                        System.Diagnostics.Debug.WriteLine($"[UpdateProgress] 작업 완료 감지: Percentage={info.Percentage}, IsCompleted={info.IsCompleted}, HasError={info.HasError}");
                        
                        lock (lockObject)
                        {
                            if (operationCompleted) // 이중 체크
                            {
                                System.Diagnostics.Debug.WriteLine("[UpdateProgress] 이미 operationCompleted=true, 리턴");
                                return;
                            }
                                
                            operationCompleted = true; // 작업 완료 표시
                            System.Diagnostics.Debug.WriteLine("[UpdateProgress] operationCompleted=true 설정 완료");
                        }
                        
                        // 에러가 있으면 메시지 표시
                        if (info.HasError)
                        {
                            System.Diagnostics.Debug.WriteLine($"[UpdateProgress] 오류 감지: {info.ErrorMessage}");
                            // 메시지 박스는 동기적으로 처리
                            MessageBox.Show(info.ErrorMessage, "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        
                        // 작업 결과 저장 (Worker_RunWorkerCompleted에서 사용)
                        DialogResult result = info.HasError ? DialogResult.Cancel : DialogResult.OK;
                        this.Tag = result;
                        this.DialogResult = result;
                        System.Diagnostics.Debug.WriteLine($"[UpdateProgress] DialogResult={result} 설정 (Tag에도 저장)");
                        
                        // 작업이 완료되었음을 표시하고 폼을 닫는 처리는 RunWorkerCompleted에서 수행
                        System.Diagnostics.Debug.WriteLine("[UpdateProgress] 작업 완료 - RunWorkerCompleted에서 폼 닫기 예정");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateProgress] 예외: {ex.Message}");
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // 작업 실행 (CancellationToken 전달)
                workAction(progressReporter, cancellationTokenSource.Token);
                
                // 작업이 완료된 후 상태 정보 기록
                System.Diagnostics.Debug.WriteLine($"[Worker_DoWork] 작업 완료 - CancellationPending={worker.CancellationPending}, cancelRequested={cancelRequested}, Token.IsCancellationRequested={cancellationTokenSource.Token.IsCancellationRequested}");
                
                // 작업 중 취소 요청이 있었는지 확인
                if (worker.CancellationPending || cancelRequested)
                {
                    System.Diagnostics.Debug.WriteLine($"[Worker_DoWork] 취소 요청으로 e.Cancel=true 설정 - CancellationPending={worker.CancellationPending}, cancelRequested={cancelRequested}");
                    e.Cancel = true;
                }
            }
            catch (OperationCanceledException)
            {
                // 취소 예외는 정상적인 취소로 처리
                System.Diagnostics.Debug.WriteLine("[Worker_DoWork] OperationCanceledException 발생, e.Cancel=true 설정");
                e.Cancel = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Worker_DoWork] 예외 발생: {ex.Message}");
                
                if (!isClosing && !cancelRequested) // 취소 중이면 예외 처리하지 않음
                {
                    var errorInfo = new ProgressInfo
                    {
                        IsCompleted = true,
                        HasError = true,
                        ErrorMessage = ex.Message,
                        StatusMessage = "오류 발생: " + ex.Message
                    };
                    
                    try
                    {
                        // 진행 상태 직접 업데이트 (Progress 객체 사용)
                        progressReporter.Report(errorInfo);
                    }
                    catch
                    {
                        // 무시
                    }
                    
                    e.Result = errorInfo;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[Worker_DoWork] 폼 닫기 중 또는 취소 요청됨, e.Cancel=true 설정 - isClosing={isClosing}, cancelRequested={cancelRequested}");
                    e.Cancel = true;
                }
            }
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] 진입 (isClosing={isClosing}, operationCompleted={operationCompleted}, e.Cancelled={e.Cancelled}, e.Error={(e.Error != null ? "있음" : "없음")})");
                
                lock (lockObject)
                {
                    // 이미 폼이 닫히는 중이면 추가 처리 없음
                    if (isClosing)
                    {
                        System.Diagnostics.Debug.WriteLine("[Worker_RunWorkerCompleted] 이미 폼 닫는 중, 리턴");
                        return;
                    }
                    
                    isClosing = true; // 폼이 닫히기 시작함을 표시
                    
                    // 아직 완료 처리되지 않았다면 여기서 처리
                    if (!operationCompleted)
                    {
                        operationCompleted = true;
                        
                        // 작업이 취소되었거나 예외가 발생한 경우
                        if (e.Cancelled)
                        {
                            System.Diagnostics.Debug.WriteLine("[Worker_RunWorkerCompleted] e.Cancelled=true, DialogResult.Cancel 설정");
                            this.DialogResult = DialogResult.Cancel;
                        }
                        else if (e.Error != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] e.Error 발생: {e.Error.Message}");
                            try
                            {
                                MessageBox.Show("작업 중 오류가 발생했습니다: " + e.Error.Message, 
                                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            catch
                            {
                                // 무시
                            }
                            
                            this.DialogResult = DialogResult.Cancel;
                        }
                        else if (e.Result is ProgressInfo errorInfo && errorInfo.HasError)
                        {
                            System.Diagnostics.Debug.WriteLine("[Worker_RunWorkerCompleted] ProgressInfo에 에러 있음, DialogResult.Cancel 설정");
                            this.DialogResult = DialogResult.Cancel;
                        }
                        else 
                        {
                            System.Diagnostics.Debug.WriteLine("[Worker_RunWorkerCompleted] 정상 완료, DialogResult.OK 설정");
                            this.DialogResult = DialogResult.OK;
                        }
                    }
                    else
                    {
                        // operationCompleted가 이미 true인 경우, UpdateProgress에서 설정한 DialogResult를 유지
                        // 여기서는 Tag에 저장된 DialogResult 확인
                        System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] operationCompleted=true, Tag={this.Tag}, 현재 DialogResult={this.DialogResult}");
                        
                        // Tag에 저장된 값이 있으면 해당 값으로 DialogResult 유지
                        if (this.Tag is DialogResult savedResult)
                        {
                            if (e.Cancelled)
                            {
                                System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] e.Cancelled=true이지만 Tag 값 {savedResult}로 복원");
                            }
                            this.DialogResult = savedResult;
                            System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] DialogResult={savedResult} (Tag에서 복원)");
                        }
                        else if (e.Cancelled)
                        {
                            // Tag에 값이 없고 취소된 경우에만 취소로 처리
                            System.Diagnostics.Debug.WriteLine("[Worker_RunWorkerCompleted] e.Cancelled=true, Tag 없음, DialogResult.Cancel 설정");
                            this.DialogResult = DialogResult.Cancel;
                        }
                    }
                }
                
                // 이벤트 핸들러 제거
                UnregisterEvents();
                
                // 리소스 정리
                CleanupResources();
                
                // UI 스레드에서 안전하게 폼 닫기
                CloseFormSafely();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Worker_RunWorkerCompleted] 예외 발생: {ex.Message}");
                
                try
                {
                    // 마지막 시도로 폼 닫기
                    if (!this.IsDisposed && this.IsHandleCreated)
                    {
                        this.Close();
                    }
                }
                catch
                {
                    // 무시
                }
            }
        }

        // 이벤트 핸들러 제거 메서드
        private void UnregisterEvents()
        {
            try
            {
                if (worker != null)
                {
                    worker.DoWork -= Worker_DoWork;
                    worker.RunWorkerCompleted -= Worker_RunWorkerCompleted;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UnregisterEvents] 예외: {ex.Message}");
            }
        }
        
        // 리소스 정리 메서드
        private void CleanupResources()
        {
            try
            {
                // 취소 토큰 해제
                if (cancellationTokenSource != null)
                {
                    if (!cancellationTokenSource.IsCancellationRequested)
                    {
                        cancellationTokenSource.Cancel();
                    }
                    cancellationTokenSource.Dispose();
                    cancellationTokenSource = null;
                }
                
                // worker 객체 해제
                if (worker != null)
                {
                    worker.Dispose();
                    worker = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CleanupResources] 예외: {ex.Message}");
            }
        }
        
        // 안전하게 폼을 닫는 메서드
        private void CloseFormSafely()
        {
            try
            {
                if (this.IsDisposed || !this.IsHandleCreated)
                    return;
                    
                if (this.InvokeRequired)
                {
                    try
                    {
                        // UI 스레드에서 폼 닫기 (동기 호출로 변경)
                        this.Invoke(new Action(() => {
                            if (!this.IsDisposed && this.IsHandleCreated)
                            {
                                try
                                {
                                    System.Diagnostics.Debug.WriteLine("[CloseFormSafely] UI 스레드에서 폼 닫기 시작");
                                    this.Close();
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[CloseFormSafely] UI 스레드에서 폼 닫기 예외: {ex.Message}");
                                }
                            }
                        }));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseFormSafely] Invoke 예외: {ex.Message}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[CloseFormSafely] 직접 폼 닫기");
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CloseFormSafely] 예외 발생: {ex.Message}");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[btnCancel_Click] 취소 버튼 클릭");
            
            lock (lockObject)
            {
                if (isClosing || cancelRequested || operationCompleted)
                {
                    System.Diagnostics.Debug.WriteLine("[btnCancel_Click] 이미 진행 중, 무시");
                    return;
                }
                    
                // 취소 처리 시작
                cancelRequested = true;
                lblStatus.Text = "작업 즉시 중단 중...";
                btnCancel.Enabled = false;
                
                try
                {
                    // 취소 토큰 신호 보내기
                    cancellationTokenSource?.Cancel();
                    
                    // 백그라운드 워커에도 취소 요청
                    if (worker != null && worker.IsBusy && worker.WorkerSupportsCancellation)
                    {
                        worker.CancelAsync();
                    }
                    
                    // 2초 후에도 완료되지 않으면 강제 종료
                    System.Windows.Forms.Timer forceCloseTimer = new System.Windows.Forms.Timer();
                    forceCloseTimer.Interval = 2000;
                    forceCloseTimer.Tick += (s, args) => 
                    {
                        forceCloseTimer.Stop();
                        forceCloseTimer.Dispose();
                        
                        System.Diagnostics.Debug.WriteLine("[forceCloseTimer_Tick] 강제 종료 타이머 실행");
                        
                        lock (lockObject)
                        {
                            if (!isClosing) 
                            {
                                isClosing = true;
                                operationCompleted = true;
                                
                                // 이벤트 핸들러 제거
                                UnregisterEvents();
                                
                                // 리소스 정리
                                CleanupResources();
                                
                                this.DialogResult = DialogResult.Cancel;
                                CloseFormSafely();
                            }
                        }
                    };
                    forceCloseTimer.Start();
                }
                catch (Exception ex) 
                {
                    System.Diagnostics.Debug.WriteLine($"[btnCancel_Click] 예외: {ex.Message}");
                    
                    // 취소 처리 중 예외 발생 시 강제 종료
                    isClosing = true;
                    operationCompleted = true;
                    
                    // 이벤트 핸들러 제거
                    UnregisterEvents();
                    
                    // 리소스 정리
                    CleanupResources();
                    
                    // 폼 닫기 직접 실행
                    this.DialogResult = DialogResult.Cancel;
                    CloseFormSafely();
                }
            }
        }

        private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[ProgressForm_FormClosing] 진입 (isClosing={isClosing}, operationCompleted={operationCompleted})");
            
            lock (lockObject)
            {
                if (isClosing) // 이미 폼이 닫히는 중이면 중복 처리 방지
                {
                    System.Diagnostics.Debug.WriteLine("[ProgressForm_FormClosing] 이미 닫는 중, 리턴");
                    return;
                }
                    
                // 작업 완료되었으면 정상적으로 닫기 진행
                if (operationCompleted)
                {
                    System.Diagnostics.Debug.WriteLine("[ProgressForm_FormClosing] 작업 완료됨, 정상 닫기");
                    isClosing = true;
                    return;
                }
                    
                // 작업 중이라면 취소 처리
                if (worker != null && worker.IsBusy)
                {
                    System.Diagnostics.Debug.WriteLine("[ProgressForm_FormClosing] 작업 중, 닫기 취소하고 취소 처리");
                    // 작업 완료되지 않았으면 작업 취소 진행
                    e.Cancel = true; // 폼 닫기 취소
                    cancelRequested = true;
                    isClosing = true; // 폼이 닫히는 중임을 표시
                    lblStatus.Text = "작업 취소 중...";
                    
                    // 취소 토큰 신호 보내기
                    cancellationTokenSource?.Cancel();
                    
                    // 백그라운드 워커에도 취소 요청
                    worker.CancelAsync();
                    btnCancel.Enabled = false;
                }
            }
        }
        
        // 폼 닫기 확실하게 수행
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[OnFormClosing] 호출됨");
            
            // 이벤트 핸들러 명시적 제거
            UnregisterEvents();
            
            // 리소스 정리
            CleanupResources();
            
            base.OnFormClosing(e);
        }

        // Dispose에서도 리소스 정리
        protected override void Dispose(bool disposing)
        {
            System.Diagnostics.Debug.WriteLine($"[Dispose] 호출됨 (disposing={disposing})");
            
            if (disposing)
            {
                // 명시적 이벤트 핸들러 제거
                UnregisterEvents();
                
                // 리소스 정리
                CleanupResources();
            }
            
            base.Dispose(disposing);
        }
        
        // 모던 스타일 적용
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