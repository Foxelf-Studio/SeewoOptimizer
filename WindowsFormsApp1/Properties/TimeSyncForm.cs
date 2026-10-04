using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using SeewoOpt.Services;
using Microsoft.VisualBasic;

namespace SeewoOpt
{
    public class TimeSyncForm : Form
    {
        // 系统托盘图标
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Thread syncThread;
        // 跨线程共享状态：以下字段在 UI 线程与 syncThread 之间双向读写，
        // 必须声明为 volatile，否则 JIT 可能缓存寄存器值导致退出/完成信号无法及时生效。
        private volatile bool syncCompleted = false;
        private volatile bool isSyncing = false;
        private volatile bool isPermissionError = false;  // 权限错误标志

        // 取消令牌：替代原先的 volatile bool forceExit。
        // 原实现用 Thread.Sleep + 轮询标志实现"可取消的等待"，
        // 问题是：用户点退出后，最多要等 sleep 结束（最长 5 秒）才会响应。
        // 改为 CancellationToken 后，等待可被立即打断，退出响应从秒级降到毫秒级。
        private CancellationTokenSource _syncCts;
        private bool IsSyncCancelled
        {
            get
            {
                CancellationTokenSource cts = _syncCts;
                return cts != null && cts.IsCancellationRequested;
            }
        }

        // 设置字段（与属性对应）
        private bool _autoStart = false;
        private bool _silentStart = false;
        private int _volumeLevel = 50;
        private bool _autoVolume = false;
        private bool _killWps = false;

        // 程序版本号（从程序集读取，统一管理）
        private readonly string _versionString;

        // 公共属性（供 SettingsForm 访问）
        public bool AutoStart
        {
            get { return _autoStart; }
            set { _autoStart = value; }
        }

        public bool SilentStart
        {
            get { return _silentStart; }
            set { _silentStart = value; }
        }

        public int VolumeLevel
        {
            get { return _volumeLevel; }
            set { _volumeLevel = value; }
        }

        public bool AutoVolume
        {
            get { return _autoVolume; }
            set { _autoVolume = value; }
        }

        public bool KillWps
        {
            get { return _killWps; }
            set { _killWps = value; }
        }

        private const int MAX_RETRIES = 3;
        private const int SECOND_STAGE_RETRIES = 3;
        private const int RETRY_DELAY_MS = 2000;
        private const int FINAL_FAILURE_DELAY_MS = 5000;
        private const int STAGE_DISPLAY_DELAY_MS = 2000;
        private const int SERVER_SWITCH_DELAY_MS = 1000;

        // 时间同步成功后，等待更新检查结束的最长时间。
        // 超过则强制退出，避免网络异常时程序永远卡在托盘。
        private const int UPDATE_CHECK_WAIT_LIMIT_MS = 60000;

        // 前 4 个为主要服务器，全部失败后才启用备用服务器。
        // 原代码用字面量 4 与 i-3 表达这个阶段划分，容易看错，现提取为常量。
        private const int PRIMARY_SERVER_COUNT = 4;

        // 定义多个NTP服务器
        //
        // 顺序调整记录：原列表把 time.windows.com / time.apple.com / time.google.com /
        // time-a.nist.gov 排在前 4 位，国内源全部垫底。实测日志显示这 4 个境外源
        // 在国内网络下极不可靠——time.google.com 连续 3 个包全部超时，
        // 每次超时白等 5 秒，前 4 个服务器合计浪费 20 秒才轮到自己人。
        // 现改为国内可直连的源优先，境外源保留作兜底。
        private static readonly string[] NtpServers = {
            "ntp.aliyun.com", "ntp1.aliyun.com", "ntp.ntsc.ac.cn", "cn.pool.ntp.org",
            "time.windows.com", "time.apple.com", "time.google.com",
            "time-a.nist.gov", "time-b.nist.gov", "pool.ntp.org"
        };

        // 用于线程安全的UI更新
        private delegate void UpdateStatusDelegate(string text, Color color);
        private delegate void UpdateProgressBarDelegate(bool visible, int? value = null);
        private delegate void UpdateButtonDelegate(bool enabled);

        // UI控件
        private RichTextBox logTextBox;
        private ProgressBar progressBar;
        private Label statusLabel;
        private Button showConsoleButton;
        private Button hideConsoleButton;
        private Button retryButton;

        public TimeSyncForm()
        {
            // 获取程序集版本号
            _versionString = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            try
            {
                WriteLog("构造函数开始");

                // 订阅更新检测事件，用于显示气泡提示
                Program.UpdateDetected += (version, message) =>
                {
                    if (trayIcon != null && !this.IsDisposed)
                    {
                        // 确保在 UI 线程上操作
                        if (this.InvokeRequired)
                        {
                            this.Invoke(new MethodInvoker(() =>
                            {
                                trayIcon.ShowBalloonTip(5000, "软件更新", $"Yeah, 检测到新版本 {version}，{message}", ToolTipIcon.Info);
                            }));
                        }
                        else
                        {
                            trayIcon.ShowBalloonTip(5000, "软件更新", $"Yeah, 检测到新版本 {version}，{message}", ToolTipIcon.Info);
                        }
                    }
                };

                // 从当前 exe 文件加载图标作为窗体图标（任务栏图标）
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                LoadSettings();
                SetupForm();
                InitializeTrayIcon();
                this.Load += TimeSyncForm_Load;
                WriteLog("构造函数完成");
            }
            catch (Exception ex)
            {
                WriteLog($"构造函数异常：{ex}");
                string errorMsg = $"初始化失败：{ex.Message}\n\n程序将关闭。";
                MessageBox.Show(errorMsg, "致命错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Program.Shutdown(1);
            }
        }

        private void SetupForm()
        {
            this.Text = "川中计算机协会 - 陈叔叔系统优化工具";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.DoubleBuffered = true;

            // 创建菜单栏
            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem fileMenu = new ToolStripMenuItem("文件(&F)");
            ToolStripMenuItem settingsItem = new ToolStripMenuItem("设置(&S)");
            settingsItem.Click += (s, e) => OpenSettings();
            fileMenu.DropDownItems.Add(settingsItem);

            ToolStripMenuItem exitItem = new ToolStripMenuItem("退出(&X)");
            exitItem.Click += FileExitItem_Click;
            fileMenu.DropDownItems.Add(exitItem);

            ToolStripMenuItem viewMenu = new ToolStripMenuItem("视图(&V)");
            ToolStripMenuItem showLogItem = new ToolStripMenuItem("打开日志(&L)");
            showLogItem.Click += (s, e) => OpenLogFolder();
            viewMenu.DropDownItems.Add(showLogItem);

            menuStrip.Items.Add(fileMenu);
            menuStrip.Items.Add(viewMenu);

            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);

            statusLabel = new Label
            {
                Text = "陈叔叔正在启动系统优化...",
                Location = new Point(10, 30),
                Size = new Size(760, 25),
                Font = new Font("Microsoft YaHei", 12, FontStyle.Bold),
                ForeColor = Color.Blue
            };
            this.Controls.Add(statusLabel);

            logTextBox = new RichTextBox
            {
                Location = new Point(10, 60),
                Size = new Size(760, 400),
                Font = new Font("Consolas", 10),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(20, 20, 20),
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both
            };
            this.Controls.Add(logTextBox);

            progressBar = new ProgressBar
            {
                Location = new Point(10, 470),
                Size = new Size(760, 25),
                Style = ProgressBarStyle.Marquee,
                Visible = true
            };
            this.Controls.Add(progressBar);

            int buttonY = 510;
            int buttonWidth = 105;
            int buttonHeight = 35;
            int buttonSpacing = 12;

            retryButton = new Button
            {
                Text = "重新同步",
                Location = new Point(10, buttonY),
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 10),
                Enabled = false
            };
            retryButton.Click += RetryButton_Click;
            this.Controls.Add(retryButton);

            showConsoleButton = new Button
            {
                Text = "显示窗口",
                Location = new Point(10 + buttonWidth + buttonSpacing, buttonY),
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 10),
                Visible = false
            };
            showConsoleButton.Click += (s, e) => ShowMainWindow();
            this.Controls.Add(showConsoleButton);

            hideConsoleButton = new Button
            {
                Text = "隐藏到托盘",
                Location = new Point(10 + (buttonWidth + buttonSpacing) * 2, buttonY),
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 10)
            };
            hideConsoleButton.Click += (s, e) => MinimizeToTray();
            this.Controls.Add(hideConsoleButton);

            // GitHub 仓库按钮
            Button githubButton = new Button
            {
                Text = "GitHub 仓库",
                Location = new Point(10 + (buttonWidth + buttonSpacing) * 3, buttonY),
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 10),
               
            };
            githubButton.Click += (s, e) =>
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://github.com/Foxelf-Studio/SeewoOptimizer",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    WriteLog($"打开 GitHub 页面失败：{ex.Message}");
                    MessageBox.Show("无法打开浏览器，请手动访问：\nhttps://github.com/Foxelf-Studio/SeewoOptimizer",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            this.Controls.Add(githubButton);

            // 卸载程序按钮
            Button uninstallButton = new Button
            {
                Text = "卸载程序",
                Location = new Point(10 + (buttonWidth + buttonSpacing) * 4, buttonY),
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 10),
                ForeColor = Color.Maroon,
                UseVisualStyleBackColor = true
            };
            uninstallButton.Click += UninstallButton_Click;
            this.Controls.Add(uninstallButton);

            this.FormClosing += MainForm_FormClosing;
            this.FormClosed += (s, e) =>
            {
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }
            };
        }

        private void OpenLogFolder()
        {
            try
            {
                string folderPath = Path.GetDirectoryName(Program.LogFilePath);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                Process.Start("explorer.exe", $"/select,\"{Program.LogFilePath}\"");
            }
            catch (Exception ex)
            {
                WriteLog($"打开日志文件夹失败：{ex.Message}");
                MessageBox.Show("无法打开日志文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void TimeSyncForm_Load(object sender, EventArgs e)
        {
            try
            {
                WriteLog("Load 事件开始");
                WriteLog($"SilentStart 当前值: {SilentStart}");

                EnsureAutoStartConsistency();

                if (Program.PendingUpdateInfo != null && trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(5000, "软件更新",
                        $"Yeah, 检测到新版本 {Program.PendingUpdateInfo.Item1}，{Program.PendingUpdateInfo.Item2}",
                        ToolTipIcon.Info);
                    Program.PendingUpdateInfo = null;
                }

                if (SilentStart)
                {
                    WriteLog("静默启动，隐藏窗口到托盘");
                    this.WindowState = FormWindowState.Minimized;
                    this.ShowInTaskbar = false;
                    MinimizeToTray();
                }
                else
                {
                    WriteLog("非静默启动，显示窗口");
                    this.Show();
                    this.WindowState = FormWindowState.Normal;
                    this.Activate();
                    this.BringToFront();
                }

                System.Windows.Forms.Timer delayTimer = new System.Windows.Forms.Timer();
                delayTimer.Interval = 2000;
                delayTimer.Tick += (s, args) =>
                {
                    delayTimer.Stop();
                    WriteLog("开始启动同步进程");
                    StartSyncProcess();
                };
                delayTimer.Start();

                WriteLog("Load 事件完成");
            }
            catch (Exception ex)
            {
                WriteLog($"Load 事件异常：{ex}");
                MessageBox.Show($"加载窗体时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WriteLog(string message)
        {
            LogService.Write(message);
        }

        private void EnsureAutoStartConsistency()
        {
            try
            {
                AutoStartResult result = AutoStartService.EnsureConsistency(
                    _autoStart, Application.ExecutablePath);

                WriteLog("EnsureAutoStartConsistency: " + result.Detail
                    + (result.Success ? "" : " (失败: " + result.ErrorMessage + ")"));
            }
            catch (Exception ex)
            {
                WriteLog("EnsureAutoStartConsistency 异常：" + ex);
            }
        }

        private void FileExitItem_Click(object sender, EventArgs e)
        {
            if (isSyncing)
            {
                DialogResult result = MessageBox.Show("同步正在进行中，宁真的要强制退出吗？",
                    "你真的要退出吗？", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    CancelSync();
                    syncCompleted = true;

                    if (syncThread != null && syncThread.IsAlive)
                    {
                        if (!syncThread.Join(3000))
                            WriteLog("等待同步线程退出超时");
                    }

                    if (trayIcon != null)
                    {
                        trayIcon.Visible = false;
                        trayIcon.Dispose();
                    }

                    Application.Exit();
                }
            }
            else
            {
                CancelSync();
                syncCompleted = true;
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }
                Application.Exit();
            }
        }

        private void RetryButton_Click(object sender, EventArgs e)
        {
            if (!isSyncing)
                StartSyncProcess();
        }

        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();

            ToolStripMenuItem showItem = new ToolStripMenuItem("显示窗口");
            showItem.Click += (s, e) => ShowMainWindow();
            trayMenu.Items.Add(showItem);

            trayMenu.Items.Add(new ToolStripSeparator());

            ToolStripMenuItem exitItem = new ToolStripMenuItem("退出程序");
            exitItem.Click += (s, e) => ExitProgram();
            trayMenu.Items.Add(exitItem);

            // 从当前 exe 文件加载图标作为托盘图标
            Icon trayIconIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            trayIcon = new NotifyIcon
            {
                Icon = trayIconIcon ?? SystemIcons.Information,
                Text = "川中计算机协会 - 系统优化工具",
                ContextMenuStrip = trayMenu,
                Visible = false
            };

            trayIcon.DoubleClick += (s, e) => ShowMainWindow();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsSyncCancelled || isPermissionError || syncCompleted)
            {
                if (trayIcon != null)
                {
                    try
                    {
                        trayIcon.Visible = false;
                        trayIcon.Dispose();
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"清理托盘图标时出错：{ex.Message}");
                    }
                    finally
                    {
                        trayIcon = null;
                    }
                }
                return;
            }

            if (!syncCompleted)
            {
                e.Cancel = true;
                MinimizeToTray();
            }
        }

        private void MinimizeToTray()
        {
            WriteLog("MinimizeToTray 被调用");
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(MinimizeToTray));
                return;
            }

            this.Hide();
            showConsoleButton.Visible = true;

            if (trayIcon != null)
            {
                trayIcon.Visible = true;
                trayIcon.ShowBalloonTip(3000, "系统优化提示",
                    "川中计协提醒您：这是在优化系统，不必关闭，为不打扰您使用电脑，已隐藏窗口！",
                    ToolTipIcon.Info);
            }
        }

        private void ShowMainWindow()
        {
            WriteLog("ShowMainWindow 被调用");
            if (this.InvokeRequired)
            {
                this.Invoke(new MethodInvoker(ShowMainWindow));
                return;
            }

            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
            showConsoleButton.Visible = false;

            if (trayIcon != null)
                trayIcon.Visible = false;
        }

        private void ExitProgram()
        {
            DialogResult result = MessageBox.Show("宁确定要退出系统优化工具吗？",
                "确认退出", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                syncCompleted = true;
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }
                Application.Exit();
            }
        }

        /// <summary>
        /// 可被取消的等待。原先用 Thread.Sleep + 轮询 forceExit，
        /// 用户点退出后最长要等 5 秒才响应；现在取消后立即返回 true。
        /// </summary>
        private bool WaitOrCancel(int milliseconds)
        {
            if (IsSyncCancelled) return true;

            try
            {
                return _syncCts.Token.WaitHandle.WaitOne(milliseconds);
            }
            catch (ObjectDisposedException)
            {
                // CTS 已被释放，说明流程已结束
                return true;
            }
        }

        /// <summary>请求取消同步流程。可从任意线程调用，立即生效。</summary>
        private void CancelSync()
        {
            try
            {
                _syncCts?.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // CTS 已被释放，说明流程已结束，忽略
            }
        }

        /// <summary>同步被取消时写日志并返回</summary>
        private bool AbortIfCancelled(string stage)
        {
            if (!IsSyncCancelled) return false;
            WriteLog("同步被用户取消" + (string.IsNullOrEmpty(stage) ? "" : "（" + stage + "）"));
            return true;
        }

        private void StartSyncProcess()
        {
            WriteLog("StartSyncProcess 被调用");
            if (isSyncing) return;

            isSyncing = true;
            syncCompleted = false;
            isPermissionError = false;

            // 每次同步开始都重建 CTS，保证上一轮的取消状态不残留
            if (_syncCts != null) { try { _syncCts.Dispose(); } catch { } }
            _syncCts = new CancellationTokenSource();

            UpdateButton(false);

            logTextBox.Clear();
            logTextBox.AppendText($"PEACE & LOVE 川中计算机协会 陈叔叔希沃系统优化工具 ver.{_versionString}\n");
            logTextBox.AppendText(new string('-', 50) + "\n");

            syncThread = new Thread(new ThreadStart(() =>
            {
                try
                {
                    WriteLog("同步线程开始");
                    UpdateStatus("正在初始化...", Color.Yellow);
                    UpdateProgressBar(true);

                    AddLog($"PEACE & LOVE 川中计算机协会 陈叔叔希沃系统优化工具 ver.{_versionString}\n", Color.DarkBlue);
                    AddLog($"可用时间服务器: {NtpServers.Length} 个\n", Color.DarkBlue);
                    AddLog(new string('-', 50) + "\n", Color.DarkGray);

                    if (AutoVolume)
                    {
                        WriteLog("开始音量调节");
                        UpdateStatus($"正在调节系统音量至 {VolumeLevel}%...", Color.DarkBlue);
                        AddLog($"正在调节系统音量至 {VolumeLevel}%... ", Color.Black);

                        bool volumeSet = VolumeController.SetMasterVolumePercent(VolumeLevel);
                        if (volumeSet)
                        {
                            AddLog("√ 完成\n", Color.Green);
                            AddLog($"系统音量已调节至 {VolumeLevel}%\n", Color.DarkGreen);
                        }
                        else
                        {
                            AddLog("× 失败\n", Color.Red);
                            AddLog("音量调节失败，但将继续进行时间同步\n", Color.DarkRed);
                        }
                        WriteLog("音量调节完成");
                    }
                    else
                    {
                        AddLog("音量调节已禁用，跳过\n", Color.Gray);
                    }

                    if (KillWps)
                    {
                        WriteLog("开始结束WPS进程");
                        UpdateStatus("正在结束WPS进程...", Color.DarkBlue);
                        AddLog("\n正在结束WPS相关进程... ", Color.Black);

                        // ProcessCleaner 逐个回调进度，界面只负责显示
                        int killedCount = ProcessCleaner.KillWpsProcesses(r =>
                        {
                            if (r.ProcessId == 0)
                            {
                                AddLog($"× {r.ProcessName}: {r.Message}\n", Color.Red);
                            }
                            else
                            {
                                AddLog($"结束进程: {r.ProcessName} (ID: {r.ProcessId})... ", Color.DarkGray);
                                AddLog(r.Success ? "√ 成功\n" : $"× {r.Message}\n",
                                       r.Success ? Color.Green : Color.Red);
                            }
                        });

                        if (killedCount > 0)
                        {
                            AddLog($"√ 完成 (已结束 {killedCount} 个WPS进程)\n", Color.Green);
                            AddLog("WPS进程已结束，继续进行时间同步\n", Color.DarkGreen);
                        }
                        else
                        {
                            AddLog($"√ 完成 (未发现WPS进程)\n", Color.Green);
                            AddLog("未发现WPS进程，继续进行时间同步\n", Color.DarkGreen);
                        }
                        WriteLog("结束WPS进程完成");
                    }
                    else
                    {
                        AddLog("\n跳过结束WPS进程（用户设置）\n", Color.Gray);
                    }

                    bool success = false;
                    bool adminPermissionError = false;

                    UpdateStatus("开始时间同步进程...", Color.DarkCyan);
                    AddLog("开始时间同步进程...\n", Color.DarkCyan);
                    AddLog(new string('=', 60) + "\n", Color.DarkGray);

                    UpdateStatus("第一阶段: 主要时间服务器", Color.DarkCyan);
                    AddLog("\n第一阶段: 主要时间服务器\n", Color.DarkCyan);

                    for (int i = 0; i < PRIMARY_SERVER_COUNT && !success && !adminPermissionError; i++)
                    {
                        if (AbortIfCancelled("第一阶段")) return;
                        WriteLog($"尝试服务器 {i + 1}: {NtpServers[i]}");
                        UpdateStatus($"尝试服务器 {i + 1}/{PRIMARY_SERVER_COUNT}: {NtpServers[i]}", Color.DarkBlue);
                        AddLog($"尝试服务器 {i + 1}/{PRIMARY_SERVER_COUNT}: {NtpServers[i]}\n", Color.Black);
                        success = SyncTimeWithServer(NtpServers[i], ref adminPermissionError, false);

                        if (!success && !adminPermissionError && i < PRIMARY_SERVER_COUNT - 1)
                        {
                            AddLog($"\n{new string('-', 40)}\n", Color.DarkGray);
                            AddLog("尝试下一个服务器...\n", Color.DarkGray);
                            if (WaitOrCancel(SERVER_SWITCH_DELAY_MS)) return;
                        }
                    }

                    if (!success && !adminPermissionError)
                    {
                        WriteLog("第一阶段失败，进入第二阶段");
                        UpdateStatus("第一阶段失败，启动备用服务器", Color.DarkOrange);
                        AddLog("\n" + new string('=', 60) + "\n", Color.DarkOrange);
                        AddLog("第一阶段同步失败，启动备用服务器\n", Color.DarkOrange);
                        AddLog(new string('=', 60) + "\n", Color.DarkOrange);

                        if (AbortIfCancelled(null)) return;
                        if (WaitOrCancel(STAGE_DISPLAY_DELAY_MS)) return;

                        UpdateStatus("第二阶段: 备用时间服务器", Color.DarkCyan);
                        AddLog("第二阶段: 备用时间服务器\n", Color.DarkCyan);

                        int totalSecondStage = NtpServers.Length - PRIMARY_SERVER_COUNT;
                        for (int i = PRIMARY_SERVER_COUNT; i < NtpServers.Length && !success && !adminPermissionError; i++)
                        {
                            if (AbortIfCancelled("第二阶段")) return;
                            int attemptNumber = i - PRIMARY_SERVER_COUNT + 1;
                            WriteLog($"尝试备用服务器 {attemptNumber}: {NtpServers[i]}");
                            UpdateStatus($"尝试备用服务器 {attemptNumber}/{totalSecondStage}: {NtpServers[i]}", Color.DarkBlue);
                            AddLog($"尝试备用服务器 {attemptNumber}/{totalSecondStage}: {NtpServers[i]}\n", Color.Black);
                            success = SyncTimeWithServer(NtpServers[i], ref adminPermissionError, true);

                            if (!success && !adminPermissionError && i < NtpServers.Length - 1)
                            {
                                AddLog($"\n{new string('-', 40)}\n", Color.DarkGray);
                                AddLog("尝试下一个备用服务器...\n", Color.DarkGray);
                                if (WaitOrCancel(SERVER_SWITCH_DELAY_MS)) return;
                            }
                        }
                    }

                    AddLog("\n" + new string('=', 60) + "\n", Color.DarkGray);

                    if (success)
                    {
                        WriteLog("时间同步成功");
                        UpdateStatus("时间同步成功！", Color.Green);
                        AddLog("√ 时间同步成功！\n", Color.Green);
                        AddLog($"\n和时间的同步率达到100%！程序将隐藏到托盘继续检查更新，完成后自动退出。\n", Color.DarkGreen);

                        if (trayIcon != null)
                            trayIcon.ShowBalloonTip(3000, "时间同步完成", "系统时间已成功同步！后台更新检测中...", ToolTipIcon.Info);

                        // 必须在下面检查 UpdateCheckCompleted 之前重试。
                        // 本工具的使用场景就是"时钟错误"，而时钟错误会让
                        // 启动时那次更新检查必然因证书未生效而失败。
                        // 若不重置标志，UI 会读到上一轮残留的 true 直接退出，
                        // 用户永远收不到更新——开机时钟总是不对就永久失效。
                        Program.RetryUpdateCheckAfterSync();

                        if (WaitOrCancel(2000)) return;
                        this.Invoke(new MethodInvoker(() => MinimizeToTray()));

                        this.Invoke(new MethodInvoker(() =>
                        {
                            if (Program.UpdateCheckCompleted)
                            {
                                syncCompleted = true;
                                Application.Exit();
                                return;
                            }

                            // 兜底上限：更新检查的网络请求若因断网等原因一直不返回，
                            // 没有这个上限程序会永远卡在托盘不退出。
                            // 校时本身已经完成，更新没查到也不该拖着用户。
                            int elapsedTicks = 0;

                            System.Windows.Forms.Timer updateCheckTimer = new System.Windows.Forms.Timer();
                            updateCheckTimer.Interval = 1000;
                            updateCheckTimer.Tick += (s, args) =>
                            {
                                elapsedTicks += updateCheckTimer.Interval;

                                if (Program.UpdateCheckCompleted || elapsedTicks >= UPDATE_CHECK_WAIT_LIMIT_MS)
                                {
                                    updateCheckTimer.Stop();

                                    if (!Program.UpdateCheckCompleted)
                                        WriteLog($"等待更新检查超过 {UPDATE_CHECK_WAIT_LIMIT_MS / 1000} 秒，强制退出");

                                    syncCompleted = true;
                                    Application.Exit();
                                }
                            };
                            updateCheckTimer.Start();
                        }));
                    }
                    else if (adminPermissionError)
                    {
                        WriteLog("权限错误");
                        UpdateStatus("权限错误", Color.Red);
                        AddLog("× 权限错误\n", Color.Red);
                        AddLog("\n（悲报！） 由于权限问题，时间同步失败。\n", Color.DarkRed);
                        AddLog("请右键点击程序，选择以管理员身份运行\n", Color.DarkRed);

                        DialogResult result = MessageBox.Show("需要管理员权限才能修改系统时间。是否以管理员身份重新启动程序？",
                            "权限不足", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            ProcessStartInfo startInfo = new ProcessStartInfo(Application.ExecutablePath)
                            {
                                Verb = "runas",
                                UseShellExecute = true
                            };
                            try
                            {
                                Process.Start(startInfo);
                            }
                            catch { }
                            Program.Shutdown(0);
                        }

                        UpdateButton(true);
                        UpdateProgressBar(false);
                        isPermissionError = true;
                    }
                    else
                    {
                        WriteLog("同步失败");
                        UpdateStatus("同步失败", Color.Red);
                        AddLog("× 同步失败\n", Color.Red);
                        AddLog($"\n（悲报！） 经过 {NtpServers.Length} 个服务器的尝试后，时间同步失败。\n", Color.DarkRed);
                        AddLog($"程序将在 {FINAL_FAILURE_DELAY_MS / 1000} 秒后自动关闭...\n", Color.DarkRed);

                        WaitOrCancel(FINAL_FAILURE_DELAY_MS);
                        syncCompleted = true;
                        this.Invoke(new MethodInvoker(this.Close));
                    }
                    WriteLog("同步线程正常结束");
                }
                catch (Exception ex)
                {
                    WriteLog($"同步线程异常：{ex}");
                    try
                    {
                        UpdateStatus($"错误: {ex.Message}", Color.Red);
                        AddLog($"错误: {ex.Message}\n", Color.Red);
                    }
                    catch { }
                    this.Invoke(new MethodInvoker(this.Close));
                }
                finally
                {
                    WriteLog("同步线程 finally 块");
                    if (!syncCompleted)
                    {
                        try
                        {
                            UpdateProgressBar(false);
                            UpdateButton(true);
                        }
                        catch { }
                    }
                    isSyncing = false;
                }
            }));

            syncThread.IsBackground = true;
            syncThread.Start();
        }

        private void UpdateStatus(string text, Color color)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.Invoke(new UpdateStatusDelegate(UpdateStatus), text, color);
                }
                catch { }
                return;
            }
            statusLabel.Text = text;
            statusLabel.ForeColor = color;
            statusLabel.Refresh();
        }

        private void UpdateProgressBar(bool visible, int? value = null)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.Invoke(new UpdateProgressBarDelegate(UpdateProgressBar), visible, value);
                }
                catch { }
                return;
            }
            progressBar.Visible = visible;
            if (value.HasValue)
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = Math.Min(Math.Max(value.Value, 0), 100);
            }
            else
            {
                progressBar.Style = ProgressBarStyle.Marquee;
            }
        }

        private void UpdateButton(bool enabled)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.Invoke(new UpdateButtonDelegate(UpdateButton), enabled);
                }
                catch { }
                return;
            }
            retryButton.Enabled = enabled;
        }

        private void AddLog(string text, Color? color = null)
        {
            if (this.InvokeRequired)
            {
                try
                {
                    this.Invoke(new Action<string, Color?>(AddLog), text, color);
                }
                catch { }
                return;
            }
            if (color.HasValue)
                logTextBox.SelectionColor = color.Value;

            logTextBox.AppendText(text);
            logTextBox.SelectionStart = logTextBox.Text.Length;
            logTextBox.ScrollToCaret();

            if (color.HasValue)
                logTextBox.SelectionColor = logTextBox.ForeColor;

            logTextBox.Refresh();

        }
private static bool SyncTimeWithServer(string ntpServer, ref bool adminPermissionError, bool isSecondStage = false)
        {
            int maxRetries = isSecondStage ? SECOND_STAGE_RETRIES : MAX_RETRIES;
            int retryCount = 0;
            bool success = false;

            while (retryCount <= maxRetries && !success && !adminPermissionError)
            {
                try
                {
                    NtpResult ntp = NtpClient.Query(ntpServer);

                    if (!ntp.IsValid)
                    {
                        // 协议校验不通过（stratum / mode / LI 异常），重试下一个服务器
                        LogService.Write($"NTP 响应校验失败: {ntp.ValidationError}");
                        HandleRetry(ref retryCount,
                            ntp.ValidationError ?? $"无法从 {ntpServer} 获取时间", maxRetries);
                        continue;
                    }

                    LogService.Write(ntp.Describe());

                    // 直接把 NTP 的 UTC 交给系统时钟。绝不在此处做时区换算：
                    // SetSystemTime 收的就是 UTC，Windows 显示时会自行应用当前时区。
                    // 曾经在这里转成本地时间再写入，导致系统时间整整多出一个时区（+8 小时）。
                    SetTimeResult setResult = SystemTimeSetter.SetToUtc(ntp.UtcTime);

                    if (setResult.Success)
                    {
                        success = true;
                        // 日志仍打本地时间，方便人读；但写进系统的是上面那个 UTC 值。
                        LogService.Write($"系统时钟已设置为 UTC {ntp.UtcTime:yyyy-MM-dd HH:mm:ss}"
                                        + $"（本地时间 {SystemTimeSetter.ToLocalTime(ntp.UtcTime):yyyy-MM-dd HH:mm:ss}）");
                    }
                    else if (setResult.PermissionDenied)
                    {
                        LogService.Write("设置系统时间需要管理员权限");
                        adminPermissionError = true;
                    }
                    else
                    {
                        HandleRetry(ref retryCount,
                            setResult.ErrorMessage ?? "设置系统时间失败", maxRetries);
                    }
                }
                catch (Exception ex)
                {
                    if (SystemTimeSetter.IsPermissionException(ex))
                        adminPermissionError = true;
                    else
                        HandleRetry(ref retryCount, ex.Message, maxRetries);
                }
            }
            return success;
        }

        private static void HandleRetry(ref int retryCount, string errorMessage, int maxRetries)
        {
            retryCount++;

            // 修正记录：errorMessage 原先被接收后直接丢弃，
            // 导致"设置系统时间失败"之类的关键错误在日志中完全不可见，
            // 排查时只能看到"NTP 响应有效"却不知下一步为何失败。
            LogService.Write($"  第 {retryCount} 次尝试失败: {errorMessage}");

            if (retryCount <= maxRetries)
                Thread.Sleep(RETRY_DELAY_MS);
        }

        private void LoadSettings()
        {
            // 读取与容错已下沉到 SettingsStore，此处只负责把结果映射到界面状态
            AppSettings settings = SettingsStore.Load();

            _autoStart = settings.AutoStart;
            _silentStart = settings.SilentStart;
            _autoVolume = settings.AutoVolume;
            _volumeLevel = settings.VolumeLevel;
            _killWps = settings.KillWps;
        }

        public void SaveSettings()
        {
            try
            {
                SettingsStore.Save(new AppSettings
                {
                    AutoStart = _autoStart,
                    SilentStart = _silentStart,
                    AutoVolume = _autoVolume,
                    VolumeLevel = _volumeLevel,
                    KillWps = _killWps
                });

                UpdateAutoStartTask();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateAutoStartTask()
        {
            WriteLog("UpdateAutoStartTask 开始, _autoStart=" + _autoStart);

            AutoStartResult result = AutoStartService.Sync(_autoStart, Application.ExecutablePath);

            if (!string.IsNullOrEmpty(result.Detail))
                WriteLog(result.Detail);

            if (result.Success)
            {
                WriteLog("UpdateAutoStartTask 完成");
            }
            else
            {
                WriteLog("UpdateAutoStartTask 失败: " + result.ErrorMessage);
                if (result.ShouldNotifyUser)
                {
                    MessageBox.Show("设置开机自启失败，请尝试以管理员身份运行程序。",
                        "任务计划创建失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void OpenSettings()
        {
            using (var settingsForm = new SettingsForm(this))
            {
                settingsForm.ShowDialog();
            }
        }

        /// <summary>
        /// 卸载程序按钮点击事件
        /// </summary>
        private void UninstallButton_Click(object sender, EventArgs e)
        {
            // 使用 UI 线程执行，避免跨线程问题
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action(() => UninstallButton_Click(sender, e)));
                return;
            }

            // 弹出输入框，要求输入暗号
            string input = Microsoft.VisualBasic.Interaction.InputBox(
                "请输入卸载暗号以确认卸载程序： （暗号请在github仓库readme中获取）",
                "卸载确认",
                "",
                -1, -1);

            if (input != "0319")
            {
                MessageBox.Show("暗号错误，卸载已取消。可在github仓库readme找到暗号。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 确认卸载
            DialogResult confirm = MessageBox.Show(
                "确定要卸载本程序吗？\n\n这将删除：\n" +
                "• 程序生成的所有日志文件\n" +
                "• 更新缓存目录\n" +
                "• 注册表中的设置项\n" +
                "• 开机自启任务计划\n\n" +
                "注意：程序文件本身和dll文件不会被删除，请手动删除。",
                "确认卸载",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes)
                return;

            // 开始卸载，并在主窗口显示日志（确保在 UI 线程）
            AddLog("\n========== 开始卸载程序 ==========\n", Color.DarkRed);
            AddLog("正在清理程序生成的数据...\n", Color.DarkRed);


            bool success = true;

            // 停止所有日志写入。日志目录即将被删除，若继续写入，
            // 后续 WriteLog 会因目录不存在而静默失败，导致日志永久丢失。
            LogService.Enabled = false;

            // 1. 删除日志目录
            try
            {
                string logDir = Path.GetDirectoryName(Program.LogFilePath);
                if (Directory.Exists(logDir))
                {
                    Directory.Delete(logDir, true);
                    AddLog($"√ 已删除日志目录: {logDir}\n", Color.Green);
                }
                else
                {
                    AddLog("日志目录不存在，跳过\n", Color.Gray);
                }
            }
            catch (Exception ex)
            {
                AddLog($"× 删除日志目录失败: {ex.Message}\n", Color.Red);
                success = false;
            }


            // 2. 删除更新目录
            try
            {
                string updateDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TimeSyncTool", "updates");
                if (Directory.Exists(updateDir))
                {
                    Directory.Delete(updateDir, true);
                    AddLog($"√ 已删除更新目录: {updateDir}\n", Color.Green);
                }
                else
                {
                    AddLog("更新目录不存在，跳过\n", Color.Gray);
                }
            }
            catch (Exception ex)
            {
                AddLog($"× 删除更新目录失败: {ex.Message}\n", Color.Red);
                success = false;
            }


            // 3. 删除注册表项
            try
            {
                SettingsStore.Delete();
                AddLog("√ 已删除注册表项: HKCU\\Software\\TimeSyncTool\n", Color.Green);
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("找不到"))
                    AddLog($"× 删除注册表项失败: {ex.Message}\n", Color.Red);
                else
                    AddLog("注册表项不存在，跳过\n", Color.Gray);
            }


            // 4. 删除开机自启任务计划
            try
            {
                if (AutoStartService.Remove())
                {
                    AddLog("√ 已删除开机自启任务计划\n", Color.Green);
                }
                else
                {
                    AddLog("删除任务计划失败\n", Color.Red);
                    success = false;
                }
            }
            catch (Exception ex)
            {
                AddLog($"× 删除任务计划失败: {ex.Message}\n", Color.Red);
                success = false;
            }

            // 5. 提示删除完成
            AddLog("\n========== 卸载完成 ==========\n", success ? Color.Green : Color.Red);
            if (success)
            {
                AddLog("程序生成的数据已清理完毕。\n", Color.DarkGreen);
                AddLog("您可以手动删除本程序文件（SeewoOpt.exe）及其所在文件夹。\n", Color.DarkGreen);
            }
            else
            {
                AddLog("部分项目清理失败，请手动检查。\n", Color.DarkRed);
            }

            // 询问是否立即退出程序
            DialogResult exit = MessageBox.Show(
                "卸载操作已完成。是否立即退出程序？",
                "卸载完成",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (exit == DialogResult.Yes)
            {
                // 清理托盘图标（确保在 UI 线程）
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                    trayIcon = null;
                }

                // 强制终止进程，避免任何残留
                Program.Shutdown(0);
            }
            else
            {
                // 用户选择继续使用：日志目录已被删除，重新开启写入会自动重建，
                // 否则本次会话剩余时间将完全没有日志。
                LogService.Enabled = true;
            }
        }

    }
}