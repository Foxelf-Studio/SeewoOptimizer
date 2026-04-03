using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using WindowsFormsApp1;

namespace TimeSyncTool
{
    public class TimeSyncForm : Form
    {
        // 系统托盘图标
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private Thread syncThread;
        private bool syncCompleted = false;
        private bool isSyncing = false;
        private bool forceExit = false;          // 强制退出标志
        private bool isPermissionError = false;  // 权限错误标志

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

        private const string SETTINGS_REGISTRY_PATH = @"Software\TimeSyncTool";
        private const string TASK_NAME = "TimeSyncTool";

        // 日志文件路径
        private readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool", "startup.log");

        // Windows API 用于设置系统时间
        [StructLayout(LayoutKind.Sequential)]
        public struct SYSTEMTIME
        {
            public short wYear;
            public short wMonth;
            public short wDayOfWeek;
            public short wDay;
            public short wHour;
            public short wMinute;
            public short wSecond;
            public short wMilliseconds;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetLocalTime(ref SYSTEMTIME st);

        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();

        // Core Audio API for volume control
        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        internal class MMDeviceEnumerator { }

        internal enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count
        }

        internal enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        internal interface IMMDeviceEnumerator
        {
            int NotImpl1();
            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        internal interface IMMDevice
        {
            [PreserveSig]
            int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
        internal interface IAudioEndpointVolume
        {
            [PreserveSig]
            int RegisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int UnregisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int GetChannelCount(out int pnChannelCount);
            [PreserveSig]
            int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            [PreserveSig]
            int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
            [PreserveSig]
            int GetMasterVolumeLevel(out float pfLevelDB);
            [PreserveSig]
            int GetMasterVolumeLevelScalar(out float pfLevel);
            [PreserveSig]
            int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
            [PreserveSig]
            int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
            [PreserveSig]
            int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            [PreserveSig]
            int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
            [PreserveSig]
            int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            [PreserveSig]
            int GetMute(out bool pbMute);
            [PreserveSig]
            int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
            [PreserveSig]
            int VolumeStepUp(Guid pguidEventContext);
            [PreserveSig]
            int VolumeStepDown(Guid pguidEventContext);
            [PreserveSig]
            int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            [PreserveSig]
            int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
        }

        private const int MAX_RETRIES = 3;
        private const int SECOND_STAGE_RETRIES = 3;
        private const int RETRY_DELAY_MS = 2000;
        private const int FINAL_FAILURE_DELAY_MS = 5000;
        private const int STAGE_DISPLAY_DELAY_MS = 2000;

        // 定义多个NTP服务器
        private static readonly string[] NtpServers = {
            "time.windows.com", "time.apple.com", "time.google.com", "time-a.nist.gov",
            "time-b.nist.gov", "pool.ntp.org", "cn.pool.ntp.org", "ntp.aliyun.com",
            "ntp1.aliyun.com", "ntp2.aliyun.com"
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
                Environment.Exit(1);
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

            retryButton = new Button
            {
                Text = "重新同步",
                Location = new Point(10, buttonY),
                Size = new Size(120, 35),
                Font = new Font("Microsoft YaHei", 10),
                Enabled = false
            };
            retryButton.Click += RetryButton_Click;
            this.Controls.Add(retryButton);

            showConsoleButton = new Button
            {
                Text = "显示窗口",
                Location = new Point(140, buttonY),
                Size = new Size(120, 35),
                Font = new Font("Microsoft YaHei", 10),
                Visible = false
            };
            showConsoleButton.Click += (s, e) => ShowMainWindow();
            this.Controls.Add(showConsoleButton);

            hideConsoleButton = new Button
            {
                Text = "隐藏到托盘",
                Location = new Point(270, buttonY),
                Size = new Size(120, 35),
                Font = new Font("Microsoft YaHei", 10)
            };
            hideConsoleButton.Click += (s, e) => MinimizeToTray();
            this.Controls.Add(hideConsoleButton);

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
                string folderPath = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                Process.Start("explorer.exe", $"/select,\"{LogFilePath}\"");
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
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                File.AppendAllText(LogFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { }
        }

        private void EnsureAutoStartConsistency()
        {
            try
            {
                bool taskExists = false;
                using (TaskService ts = new TaskService())
                {
                    taskExists = ts.GetTask(TASK_NAME) != null;
                }
                WriteLog($"当前任务计划自启项存在：{taskExists}，设置值为：{_autoStart}");
                if (taskExists != _autoStart)
                {
                    WriteLog($"不一致，执行同步");
                    UpdateAutoStartTask();
                }
            }
            catch (Exception ex)
            {
                WriteLog($"EnsureAutoStartConsistency 异常：{ex}");
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
                    forceExit = true;
                    syncCompleted = true;

                    try
                    {
                        if (syncThread != null && syncThread.IsAlive)
                            syncThread.Abort();
                    }
                    catch { }

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
                forceExit = true;
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

            trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Information,
                Text = "川中计算机协会 - 系统优化工具",
                ContextMenuStrip = trayMenu,
                Visible = false
            };

            trayIcon.DoubleClick += (s, e) => ShowMainWindow();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (forceExit || isPermissionError || syncCompleted)
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

        private void StartSyncProcess()
        {
            WriteLog("StartSyncProcess 被调用");
            if (isSyncing) return;

            isSyncing = true;
            syncCompleted = false;
            isPermissionError = false;

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
                        float targetVolume = VolumeLevel / 100.0f;
                        UpdateStatus($"正在调节系统音量至 {VolumeLevel}%...", Color.DarkBlue);
                        AddLog($"正在调节系统音量至 {VolumeLevel}%... ", Color.Black);

                        bool volumeSet = SetSystemVolume(targetVolume);
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

                        int killedCount = KillWpsProcesses();
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

                    for (int i = 0; i < 4 && !success && !adminPermissionError; i++)
                    {
                        WriteLog($"尝试服务器 {i + 1}: {NtpServers[i]}");
                        UpdateStatus($"尝试服务器 {i + 1}/4: {NtpServers[i]}", Color.DarkBlue);
                        AddLog($"尝试服务器 {i + 1}/4: {NtpServers[i]}\n", Color.Black);
                        success = SyncTimeWithServer(NtpServers[i], ref adminPermissionError, false, i + 1, 4);

                        if (!success && !adminPermissionError && i < 3)
                        {
                            AddLog($"\n{new string('-', 40)}\n", Color.DarkGray);
                            AddLog("尝试下一个服务器...\n", Color.DarkGray);
                            Thread.Sleep(1000);
                        }
                    }

                    if (!success && !adminPermissionError)
                    {
                        WriteLog("第一阶段失败，进入第二阶段");
                        UpdateStatus("第一阶段失败，启动备用服务器", Color.DarkOrange);
                        AddLog("\n" + new string('=', 60) + "\n", Color.DarkOrange);
                        AddLog("第一阶段同步失败，启动备用服务器\n", Color.DarkOrange);
                        AddLog(new string('=', 60) + "\n", Color.DarkOrange);

                        Thread.Sleep(STAGE_DISPLAY_DELAY_MS);

                        UpdateStatus("第二阶段: 备用时间服务器", Color.DarkCyan);
                        AddLog("第二阶段: 备用时间服务器\n", Color.DarkCyan);

                        int totalSecondStage = NtpServers.Length - 4;
                        for (int i = 4; i < NtpServers.Length && !success && !adminPermissionError; i++)
                        {
                            int attemptNumber = i - 3;
                            WriteLog($"尝试备用服务器 {attemptNumber}: {NtpServers[i]}");
                            UpdateStatus($"尝试备用服务器 {attemptNumber}/{totalSecondStage}: {NtpServers[i]}", Color.DarkBlue);
                            AddLog($"尝试备用服务器 {attemptNumber}/{totalSecondStage}: {NtpServers[i]}\n", Color.Black);
                            success = SyncTimeWithServer(NtpServers[i], ref adminPermissionError, true, attemptNumber, totalSecondStage);

                            if (!success && !adminPermissionError && i < NtpServers.Length - 1)
                            {
                                AddLog($"\n{new string('-', 40)}\n", Color.DarkGray);
                                AddLog("尝试下一个备用服务器...\n", Color.DarkGray);
                                Thread.Sleep(1000);
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

                        Thread.Sleep(2000);
                        this.Invoke(new MethodInvoker(() => MinimizeToTray()));

                        this.Invoke(new MethodInvoker(() =>
                        {
                            if (Program.UpdateCheckCompleted)
                            {
                                syncCompleted = true;
                                Application.Exit();
                                return;
                            }

                            System.Windows.Forms.Timer updateCheckTimer = new System.Windows.Forms.Timer();
                            updateCheckTimer.Interval = 1000;
                            updateCheckTimer.Tick += (s, args) =>
                            {
                                if (Program.UpdateCheckCompleted)
                                {
                                    updateCheckTimer.Stop();
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
                            Environment.Exit(0);
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

                        Thread.Sleep(FINAL_FAILURE_DELAY_MS);
                        syncCompleted = true;
                        this.Invoke(new MethodInvoker(this.Close));
                    }
                    WriteLog("同步线程正常结束");
                }
                catch (ThreadAbortException)
                {
                    WriteLog("同步线程被中止");
                    AddLog("\n同步已被用户中断\n", Color.DarkOrange);
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

        private int KillWpsProcesses()
        {
            int killedCount = 0;
            try
            {
                string[] wpsProcessNames = { "wps", "wpp", "et", "wpscloudsvr", "ksolaunch" };

                foreach (string processName in wpsProcessNames)
                {
                    Process[] processes = Process.GetProcessesByName(processName);
                    foreach (Process process in processes)
                    {
                        try
                        {
                            AddLog($"结束进程: {process.ProcessName} (ID: {process.Id})... ", Color.DarkGray);
                            process.Kill();
                            process.WaitForExit(3000);
                            if (process.HasExited)
                            {
                                AddLog("√ 成功\n", Color.Green);
                                killedCount++;
                            }
                            else
                            {
                                AddLog("× 超时\n", Color.Red);
                            }
                        }
                        catch (Exception ex)
                        {
                            AddLog($"× 失败: {ex.Message}\n", Color.Red);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog($"结束WPS进程时发生错误: {ex.Message}\n", Color.Red);
            }
            return killedCount;
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
            Application.DoEvents();
        }

        private static bool SetSystemVolume(float volumeLevel)
        {
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            IAudioEndpointVolume volume = null;
            try
            {
                if (volumeLevel < 0.0f || volumeLevel > 1.0f)
                    return false;

                enumerator = new MMDeviceEnumerator() as IMMDeviceEnumerator;
                if (enumerator == null) return false;

                int hr = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device);
                if (hr != 0 || device == null) return false;

                Guid IID_IAudioEndpointVolume = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
                hr = device.Activate(IID_IAudioEndpointVolume, 0, IntPtr.Zero, out object obj);
                if (hr != 0 || obj == null) return false;

                volume = obj as IAudioEndpointVolume;
                if (volume == null) return false;

                hr = volume.SetMasterVolumeLevelScalar(volumeLevel, Guid.Empty);
                return hr == 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (volume != null) Marshal.ReleaseComObject(volume);
                if (device != null) Marshal.ReleaseComObject(device);
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }
        }

        private static bool SyncTimeWithServer(string ntpServer, ref bool adminPermissionError, bool isSecondStage = false, int currentAttempt = 1, int totalAttempts = 1)
        {
            int maxRetries = isSecondStage ? SECOND_STAGE_RETRIES : MAX_RETRIES;
            int retryCount = 0;
            bool success = false;

            while (retryCount <= maxRetries && !success && !adminPermissionError)
            {
                try
                {
                    DateTime ntpTime = GetNetworkTime(ntpServer);
                    if (ntpTime != DateTime.MinValue)
                    {
                        DateTime beijingTime = ntpTime.AddHours(8);
                        if (SetLocalTime(beijingTime))
                        {
                            success = true;
                        }
                        else
                        {
                            uint errorCode = GetLastError();
                            if (IsPermissionError(errorCode))
                            {
                                adminPermissionError = true;
                            }
                            else
                            {
                                HandleRetry(ref retryCount, GetErrorMessage(errorCode), maxRetries);
                            }
                        }
                    }
                    else
                    {
                        HandleRetry(ref retryCount, $"无法从 {ntpServer} 获取时间", maxRetries);
                    }
                }
                catch (Exception ex)
                {
                    if (IsPermissionException(ex))
                        adminPermissionError = true;
                    else
                        HandleRetry(ref retryCount, ex.Message, maxRetries);
                }
            }
            return success;
        }

        private static bool IsPermissionError(uint errorCode)
        {
            return errorCode == 5 || errorCode == 1314;
        }

        private static bool IsPermissionException(Exception ex)
        {
            return ex is UnauthorizedAccessException ||
                   ex.Message.IndexOf("access", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   ex.Message.Contains("权限") ||
                   ex.Message.IndexOf("privilege", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   ex.Message.IndexOf("admin", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   ex.Message.Contains("管理员");
        }

        private static void HandleRetry(ref int retryCount, string errorMessage, int maxRetries)
        {
            retryCount++;
            if (retryCount <= maxRetries)
                Thread.Sleep(RETRY_DELAY_MS);
        }

        private static string GetErrorMessage(uint errorCode)
        {
            switch (errorCode)
            {
                case 5: return "拒绝访问！ - 需要管理员权限";
                case 1314: return "客户端没有所需的特权 - 需要管理员权限";
                case 13: return "数据无效";
                case 87: return "参数错误";
                default: return $"错误代码: {errorCode}";
            }
        }

        private static DateTime GetNetworkTime(string ntpServer)
        {
            try
            {
                var ntpData = new byte[48];
                ntpData[0] = 0x1B;

                IPAddress ipAddress = null;
                var addresses = Dns.GetHostEntry(ntpServer).AddressList;

                foreach (var addr in addresses)
                {
                    if (addr.AddressFamily == AddressFamily.InterNetwork)
                    {
                        ipAddress = addr;
                        break;
                    }
                }

                if (ipAddress == null && addresses.Length > 0)
                    ipAddress = addresses[0];

                if (ipAddress == null)
                    throw new Exception($"无法解析NTP服务器: {ntpServer}");

                var ipEndPoint = new IPEndPoint(ipAddress, 123);

                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                {
                    socket.ReceiveTimeout = 5000;
                    socket.SendTimeout = 5000;

                    IAsyncResult connectResult = socket.BeginConnect(ipEndPoint, null, null);
                    if (!connectResult.AsyncWaitHandle.WaitOne(5000, false))
                        throw new Exception("连接NTP服务器超时");
                    socket.EndConnect(connectResult);

                    IAsyncResult sendResult = socket.BeginSend(ntpData, 0, ntpData.Length, SocketFlags.None, null, null);
                    if (!sendResult.AsyncWaitHandle.WaitOne(5000, false))
                        throw new Exception("发送NTP请求超时");
                    socket.EndSend(sendResult);

                    IAsyncResult receiveResult = socket.BeginReceive(ntpData, 0, ntpData.Length, SocketFlags.None, null, null);
                    if (!receiveResult.AsyncWaitHandle.WaitOne(5000, false))
                        throw new Exception("接收NTP响应超时");
                    int bytesReceived = socket.EndReceive(receiveResult);

                    if (bytesReceived < 48)
                        throw new Exception("接收的NTP数据不完整");
                }

                ulong intPart = (ulong)ntpData[40] << 24 | (ulong)ntpData[41] << 16 | (ulong)ntpData[42] << 8 | ntpData[43];
                ulong fractPart = (ulong)ntpData[44] << 24 | (ulong)ntpData[45] << 16 | (ulong)ntpData[46] << 8 | ntpData[47];

                var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);
                var networkDateTime = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMilliseconds((long)milliseconds);

                return networkDateTime;
            }
            catch (Exception ex)
            {
                throw new Exception($"获取NTP时间失败: {ex.Message}");
            }
        }

        private static bool SetLocalTime(DateTime newTime)
        {
            try
            {
                SYSTEMTIME st = new SYSTEMTIME
                {
                    wYear = (short)newTime.Year,
                    wMonth = (short)newTime.Month,
                    wDay = (short)newTime.Day,
                    wHour = (short)newTime.Hour,
                    wMinute = (short)newTime.Minute,
                    wSecond = (short)newTime.Second,
                    wMilliseconds = (short)newTime.Millisecond
                };
                return SetLocalTime(ref st);
            }
            catch (Exception ex)
            {
                throw new Exception($"设置系统时间时发生错误: {ex.Message}");
            }
        }

        private void LoadSettings()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(SETTINGS_REGISTRY_PATH))
                {
                    object val = key.GetValue("AutoStart", true);
                    _autoStart = val is bool ? (bool)val : Convert.ToBoolean(val);

                    val = key.GetValue("SilentStart", false);
                    _silentStart = val is bool ? (bool)val : Convert.ToBoolean(val);

                    val = key.GetValue("AutoVolume", true);
                    _autoVolume = val is bool ? (bool)val : Convert.ToBoolean(val);

                    val = key.GetValue("VolumeLevel", 60);
                    _volumeLevel = val is int ? (int)val : Convert.ToInt32(val);

                    val = key.GetValue("KillWps", true);
                    _killWps = val is bool ? (bool)val : Convert.ToBoolean(val);
                }
            }
            catch
            {
                _autoStart = true;
                _silentStart = false;
                _autoVolume = true;
                _volumeLevel = 60;
                _killWps = true;
            }
        }

        public void SaveSettings()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(SETTINGS_REGISTRY_PATH))
                {
                    key.SetValue("AutoStart", _autoStart);
                    key.SetValue("SilentStart", _silentStart);
                    key.SetValue("AutoVolume", _autoVolume);
                    key.SetValue("VolumeLevel", _volumeLevel);
                    key.SetValue("KillWps", _killWps);
                }
                UpdateAutoStartTask();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateAutoStartTask()
        {
            try
            {
                WriteLog($"========== UpdateAutoStartTask 开始 ==========");
                WriteLog($"_autoStart 当前值: {_autoStart}");

                using (TaskService ts = new TaskService())
                {
                    if (_autoStart)
                    {
                        if (!File.Exists(Application.ExecutablePath))
                        {
                            WriteLog($"错误：程序文件不存在");
                            return;
                        }

                        Microsoft.Win32.TaskScheduler.Task existingTask = ts.GetTask(TASK_NAME);
                        if (existingTask != null)
                        {
                            ts.RootFolder.DeleteTask(TASK_NAME, false);
                            WriteLog("旧任务已删除");
                        }

                        TaskDefinition td = ts.NewTask();
                        td.RegistrationInfo.Description = "川中计算机协会 - 陈叔叔系统优化工具";
                        td.RegistrationInfo.Author = "TimeSyncTool";

                        td.Triggers.Add(new LogonTrigger
                        {
                            UserId = null,
                            Delay = TimeSpan.FromSeconds(5)
                        });

                        td.Actions.Add(new Microsoft.Win32.TaskScheduler.ExecAction(Application.ExecutablePath, null, null));

                        ts.RootFolder.RegisterTaskDefinition(TASK_NAME, td);
                        WriteLog($"√ 任务计划创建成功");
                    }
                    else
                    {
                        Microsoft.Win32.TaskScheduler.Task existingTask = ts.GetTask(TASK_NAME);
                        if (existingTask != null)
                        {
                            ts.RootFolder.DeleteTask(TASK_NAME, false);
                            WriteLog("√ 任务已删除");
                        }
                    }
                }
                WriteLog($"========== UpdateAutoStartTask 完成 ==========\n");
            }
            catch (Exception ex)
            {
                WriteLog($"UpdateAutoStartTask 异常：{ex.Message}");
                MessageBox.Show("设置开机自启失败，请尝试以管理员身份运行程序。",
                    "任务计划创建失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OpenSettings()
        {
            using (var settingsForm = new SettingsForm(this))
            {
                settingsForm.ShowDialog();
            }
        }
    }
}