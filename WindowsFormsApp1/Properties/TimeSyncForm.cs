using System;
using System.Collections.Generic;
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

        /// <summary>
        /// 当前存活的窗体实例，供 <see cref="WaitForBackgroundWork"/> 找到同步线程。
        ///
        /// 【为什么需要它】同步线程是窗体的私有字段，而"写日志结束标记"发生在
        /// Program.Main 的 finally 里。要让结束标记压在同步线程最后一条日志之后，
        /// Main 就得能拿到这个线程。用静态引用而不是把 thread 提为公共字段，
        /// 是为了让同步线程的生命周期仍归窗体自己管。
        ///
        /// 【为什么不能在 FormClosed 里清空它——这是个踩过的坑】
        /// 曾把它写成"窗体关闭时置 null"，结果 Main 的 finally 拿到的是 null，
        /// WaitForBackgroundWork 直接 return，Join 根本没执行，
        /// 同步线程的最后两条日志照样落在结束标记之下。
        ///
        /// 原因：窗体关闭（FormClosed）发生得比 Main 的 finally **更早**——
        /// Application.Run 要等窗体彻底销毁才返回。所以"清引用"这件事
        /// 必须发生在收尾之后，由 <see cref="ClearActiveInstance"/> 显式调用，
        /// 不能挂在窗体自己的关闭事件上。
        /// </summary>
        private static TimeSyncForm _activeInstance;

        /// <summary>
        /// 收尾完成后清掉静态引用。由 Program 在写日志结束标记之后调用，
        /// 保证引用在整个收尾窗口内都有效。
        /// </summary>
        public static void ClearActiveInstance()
        {
            _activeInstance = null;
        }

        /// <summary>
        /// 等待本窗体的后台工作线程（同步线程）结束。
        ///
        /// 【为什么必须等】日志分段用"======== RUN 结束 ========"当分界尺，
        /// 它必须压在本次运行所有日志之下。但同步线程是独立线程，
        /// 主线程走完 Application.Run 返回、写完结束标记之后，它可能还在跑，
        /// 于是它的最后两条日志（"同步线程正常结束"、"同步线程 finally 块"）
        /// 就落到了结束标记**之下**，看上去像属于下一次运行——实测确实如此。
        ///
        /// 【死锁风险，务必看清】绝不能在消息循环已经停止之后阻塞等待。
        /// 同步线程收尾时会调用 UpdateProgressBar / UpdateButton，
        /// 这两者内部用的是**同步** this.Invoke —— 它会把工作项投进 UI 队列
        /// 并等 UI 线程处理。而 Main 的 finally 执行时 Application.Run 已经返回、
        /// 消息循环已经停了，UI 线程一旦停在这里 Join，那些 Invoke 就永远
        /// 等不到人来处理，双方互锁。
        ///
        /// 所以同步线程收尾时的 UI 调用已改为 BeginInvoke（异步投递、投递即返回），
        /// 这里再配合带超时的 Join，双重保险。
        /// </summary>
        /// <param name="timeoutMs">最长等待毫秒数</param>
        public static void WaitForBackgroundWork(int timeoutMs)
        {
            TimeSyncForm form = _activeInstance;
            if (form == null) return;

            Thread t = form.syncThread;
            if (t == null || !t.IsAlive) return;

            // 先请求取消，让线程尽快从 WaitOrCancel 之类的等待中醒来，
            // 否则它可能睡满一整个延迟周期，白白耗掉等待窗口。
            try { form.CancelSync(); }
            catch { }

            // 取消之后线程若已结束，就不必等了——这一步挡掉了绝大多数情况，
            // 也让"同步线程已收尾"的正常路径完全不碰 Join 的死锁风险。
            if (!t.IsAlive) return;

            if (!t.Join(timeoutMs))
                LogService.Write($"等待同步线程收尾超时（{timeoutMs}ms），继续退出");
        }

        // 跨线程共享状态：以下字段在 UI 线程与 syncThread 之间双向读写，
        // 必须声明为 volatile，否则 JIT 可能缓存寄存器值导致退出/完成信号无法及时生效。
        //
        // 【2026-10-10 常驻改造后的状态说明】
        // 这两个字段在"同步完成即退出"的时代是控制流的关键（决定窗口能不能关、
        // 进程要不要退）。改成常驻托盘后，它们不再被任何分支读取——
        // 因此编译器会给出 CS0414 "已赋值但从未使用"。这里**保留**它们：
        //
        //   · 它们是"本次运行走到了哪一步"的唯一记录，排查用户反馈时
        //     需要从日志之外确认进程内部状态；
        //   · 卸载/重启/退出三条路径仍然要显式置位，表达"这一次流程到此为止"；
        //   · 删掉它们会让那几处赋值语句变成无主赋值，反而更难读。
        //
        // 若将来简化控制流时要把它们去掉，请连带清理上文所有赋值点，
        // 并确认 FormClosing 与退出路径已完全改用 _exitingProgram。
#pragma warning disable 0414
        private volatile bool syncCompleted = false;
        private volatile bool isPermissionError = false;  // 权限错误标志
#pragma warning restore 0414

        /// <summary>
        /// 同步线程是否正在运行。StartSyncProcess 用它防重入，
        /// StartSyncProcess / SetSyncFinishedUi 双方读写，必须 volatile。
        /// </summary>
        private volatile bool isSyncing = false;

        /// <summary>
        /// 用户是否明确要求退出程序（区别于"关闭窗口"）。
        ///
        /// 常驻模式下关窗只是缩到托盘，只有这个标志为 true 时
        /// FormClosing 才真正放行。三个入口会置位：
        /// 菜单"选项→退出"、"托盘右键→退出程序"、以及权限不足时的
        /// 重启流程（那条路径本身就是 Program.Shutdown）。
        /// </summary>
        private volatile bool _exitingProgram = false;

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

        /// <summary>启动时是否自动开始执行既定任务（需求 1）</summary>
        private bool _autoStartTask = false;

        /// <summary>首次运行日期，用于"已守护本电脑 x 天"</summary>
        private DateTime _firstRunDate = DateTime.MinValue;

        /// <summary>关机时间点（最多 5 条）</summary>
        private List<ShutdownTimeRule> _shutdownRules = new List<ShutdownTimeRule>();

        /// <summary>
        /// 常驻托盘时的巡检定时器：盯着关机时间点。
        ///
        /// 【为什么用 System.Windows.Forms.Timer 而不是 System.Threading.Timer】
        /// 命中后要弹模态对话框，而模态框必须在 UI 线程上打开。
        /// WinForms 的 Timer 本来就在 UI 线程回调，不需要再 Invoke 一次；
        /// 换成线程池定时器反而要多一层跨线程编组，还要小心和消息循环的时序。
        /// 关机判断本身是微秒级操作，占用 UI 线程完全可接受。
        /// </summary>
        private System.Windows.Forms.Timer shutdownTimer;

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

        /// <summary>启动时是否自动开始执行既定任务</summary>
        public bool AutoStartTask
        {
            get { return _autoStartTask; }
            set { _autoStartTask = value; }
        }

        /// <summary>
        /// 关机时间点。返回的是**副本**，防止 SettingsForm 在编辑过程中
        /// 直接改到正在生效的规则——若用户点了取消，改动也已经落到了运行中的配置上。
        /// </summary>
        public List<ShutdownTimeRule> ShutdownRules
        {
            get
            {
                var copy = new List<ShutdownTimeRule>();
                foreach (ShutdownTimeRule r in _shutdownRules)
                    if (r != null) copy.Add(r.Clone());
                return copy;
            }
            set
            {
                _shutdownRules = value ?? new List<ShutdownTimeRule>();
                _shutdownRules = new List<ShutdownTimeRule>(_shutdownRules);

                // 设置变更后立刻重排巡检，避免刚设好的时间点要等到
                // 下一个轮询周期才生效（最长 20 秒的延迟，用户会觉得"没反应"）
                if (shutdownTimer != null)
                    ShutdownService.ResetRuntimeState();
            }
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
        private Label statusLabel;

        /// <summary>上半屏居中的"已守护本电脑 x 天"</summary>
        private Label guardLabel;

        /// <summary>左下角"开始执行"按钮</summary>
        private Button startButton;

        private Button showConsoleButton;
        private Button hideConsoleButton;
        private Button retryButton;

        public TimeSyncForm()
        {
            // 记录当前实例，供 Program.Main 在写日志结束标记前等待同步线程收尾
            _activeInstance = this;

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
                Program.Shutdown(1, $"主窗体初始化失败：{ex.Message}");
            }
        }

        private void SetupForm()
        {
            this.Text = "川中计算机协会 - 陈叔叔系统优化工具";
            // 【窗口尺寸】900 x 600。
            //
            // 这是最初的 800x600 略加宽后的尺寸。中间经历过一次"高度减半到
            // 300"的尝试（想让窗口不挡教室大屏），但 300 高把日志区压得只剩
            // 三四行，翻日志要一直滚；而 900 宽配上 600 高，日志区才有足够
            // 纵向空间。
            //
            // 上半屏仍留给"已守护本电脑 x 天"这句主信息，日志在下半屏。
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.DoubleBuffered = true;

            // 创建菜单栏
            //
            // 【为什么"文件"改叫"选项"、"视图"改叫"日志"】
            // 这两个菜单里分别只有"设置"和"打开日志"一项，旧名字
            // （文件 / 视图）是从通用模板抄来的，对这个程序名不符实——
            // 没有"新建/打开/保存"这类文件操作，也没有任何"视图"切换。
            // 直接按内容命名，用户不必展开就能猜到里面是什么。
            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem optionMenu = new ToolStripMenuItem("选项(&O)");
            ToolStripMenuItem settingsItem = new ToolStripMenuItem("设置(&S)");
            settingsItem.Click += (s, e) => OpenSettings();
            optionMenu.DropDownItems.Add(settingsItem);

            ToolStripMenuItem exitItem = new ToolStripMenuItem("退出(&X)");
            exitItem.Click += FileExitItem_Click;
            optionMenu.DropDownItems.Add(exitItem);

            ToolStripMenuItem logMenu = new ToolStripMenuItem("日志(&L)");
            ToolStripMenuItem showLogItem = new ToolStripMenuItem("打开日志(&O)");
            showLogItem.Click += (s, e) => OpenLogFolder();
            logMenu.DropDownItems.Add(showLogItem);

            menuStrip.Items.Add(optionMenu);
            menuStrip.Items.Add(logMenu);

            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);

            // ---------------------------------------------------------------
            // 上半屏：守护天数（居中大字）
            //
            // 【为什么用 Dock=Fill 的 Label 而不是绝对坐标】
            // 这是唯一需要随窗口宽度自适应的元素，用 Dock 让 WinForms 自己
            // 处理停靠与居中，比手算坐标更稳，也避免改字号后位置错位。
            // 日志区则放在下方的 Panel 里，两者互不干扰。
            // ---------------------------------------------------------------
            guardLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = "陈叔叔希沃优化助手\n正在确认守护天数...",
                Font = new Font("Microsoft YaHei", 20, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 90, 160),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.Transparent
            };

            // 上半屏容器：高度约窗口的 1/4，把 guardLabel 撑满并居中。
            //
            // 118 这个值是配 300 高窗口时的（约占 40%）。窗口改回 600 高后
            // 若仍用 118，这块会缩成一条窄带、显得局促，故按比例放大到 170。
            Panel guardPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 170,
                BackColor = Color.FromArgb(245, 250, 255)
            };
            guardPanel.Controls.Add(guardLabel);

            // ---------------------------------------------------------------
            // 下半屏：执行过程日志
            // ---------------------------------------------------------------
            statusLabel = new Label
            {
                Text = "陈叔叔正在等待开始...",
                Dock = DockStyle.Top,
                Height = 22,
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold),
                ForeColor = Color.Blue,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0)
            };

            logTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(20, 20, 20),
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Both,
                BorderStyle = BorderStyle.None
            };

            Panel logPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(6, 0, 6, 0)
            };
            logPanel.Controls.Add(logTextBox);
            logPanel.Controls.Add(statusLabel);

            // ---------------------------------------------------------------
            // 底部按钮栏
            //
            // 【为什么按钮不用绝对坐标，而要在布局时按宽度均分】
            // 原先六个按钮从 x=8 起按固定间距一路右排，在 800 宽窗口里
            // 恰好铺满。窗口加宽到 1200 后，右侧会空出一大截——按钮全挤在
            // 左边，看着左重右轻。改为按当前宽度均分，窗口怎么变都铺满。
            // ---------------------------------------------------------------
            int buttonWidth = 105;
            int buttonHeight = 32;

            startButton = new Button
            {
                Text = "开始执行",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9),
                // 【启动时不再自动执行】按钮默认可用，但不会自动触发——
                // 由设置里的"程序启动时自动启动执行既定任务"决定是否代用户点一下。
                Enabled = true
            };
            startButton.Click += (s, e) => StartSyncProcess();

            Panel buttonBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = buttonHeight + 10,
                BackColor = Color.Transparent
            };
            buttonBar.Controls.Add(startButton);

            // 其余按钮沿用原有功能，位置统一由 LayoutButtons() 按宽度均分
            retryButton = new Button
            {
                Text = "重新同步",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9),
                Enabled = false
            };
            retryButton.Click += RetryButton_Click;
            buttonBar.Controls.Add(retryButton);

            showConsoleButton = new Button
            {
                Text = "显示窗口",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9),
                Visible = false
            };
            showConsoleButton.Click += (s, e) => ShowMainWindow();
            buttonBar.Controls.Add(showConsoleButton);

            hideConsoleButton = new Button
            {
                Text = "隐藏到托盘",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9)
            };
            hideConsoleButton.Click += (s, e) => MinimizeToTray();
            buttonBar.Controls.Add(hideConsoleButton);

            // GitHub 仓库按钮
            Button githubButton = new Button
            {
                Text = "GitHub 仓库",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9)
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
            buttonBar.Controls.Add(githubButton);

            // 卸载程序按钮
            Button uninstallButton = new Button
            {
                Text = "卸载程序",
                Size = new Size(buttonWidth, buttonHeight),
                Font = new Font("Microsoft YaHei", 9),
                ForeColor = Color.Maroon,
                UseVisualStyleBackColor = true
            };
            uninstallButton.Click += UninstallButton_Click;
            buttonBar.Controls.Add(uninstallButton);

            // 按按钮栏实际宽度均分排布，窗口缩放时重算。
            //
            // 间距 = 剩余空间 / (按钮数 + 1)，两端各留一份，左右对称。
            // 但间距设上限 MaxGap：1200 宽时若按均分算，六个按钮之间会拉开
            // 约 80px，按钮像被撒开一样、彼此失去关联。超过上限就改为
            // "固定间距 + 整组居中"，视觉上仍是一个紧凑的整体。
            buttonBar.Resize += (s, e) =>
            {
                int count = 0;
                foreach (Control c in buttonBar.Controls)
                    if (c is Button) count++;
                if (count == 0) return;

                const int MaxGap = 22;
                int totalBtnWidth = buttonWidth * count;
                int room = buttonBar.ClientSize.Width - totalBtnWidth;

                int gap = room / (count + 1);
                if (gap > MaxGap) gap = MaxGap;
                if (gap < 4) gap = 4;   // 窗口极窄时保底，避免重叠

                // 以整组居中为基准起排
                int groupWidth = totalBtnWidth + gap * (count - 1);
                int x = (buttonBar.ClientSize.Width - groupWidth) / 2;
                if (x < 2) x = 2;

                int y = (buttonBar.ClientSize.Height - buttonHeight) / 2;
                if (y < 0) y = 0;
                foreach (Control c in buttonBar.Controls)
                {
                    if (c is Button)
                    {
                        c.Location = new Point(x, y);
                        x += buttonWidth + gap;
                    }
                }
            };

            // 加控件的顺序决定停靠布局：先加的在上。
            // 这里刻意"倒序"添加，让最终自上而下是 上半屏 → 日志 → 按钮栏。
            //（Dock 的 Fill 会吃掉剩余空间，必须最后加才不被挤掉）
            this.Controls.Add(logPanel);
            this.Controls.Add(guardPanel);
            this.Controls.Add(menuStrip);
            this.Controls.Add(buttonBar);

            this.FormClosing += MainForm_FormClosing;
            this.FormClosed += (s, e) =>
            {
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }
                // 注意：这里**不能**清 _activeInstance。
                // FormClosed 早于 Application.Run 返回，也就早于 Main 的 finally；
                // 一旦在这里置 null，收尾时就拿不到同步线程，Join 会静默跳过。
                // 清理由 Program 在收尾完成后调 ClearActiveInstance 负责。
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

                // 先记首次运行日期——"守护天数"要立刻显示，不能等到同步之后。
                RecordFirstRunDateIfNeeded();

                // 【需求 3 的"时间同步成功后再确认今天是第几天"】
                // 启动时先按当前（可能错误的）系统时间显示一次，让用户
                // 一眼有东西看；同步成功后会再刷新一遍。之所以要"再确认"，
                // 是因为本机时钟可能停在 2000 年，那时算出来的天数是错的。
                RefreshGuardDays();

                StartShutdownWatcher();

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

                // 【需求 1：启动时不再默认自动执行】
                //
                // 原来固定 2 秒后无条件 StartSyncProcess()。现在改成先看设置：
                // 只有用户勾了"程序启动时自动开始执行既定任务"才代他点一下
                // 左下角那个按钮，否则就停在"等待开始"状态。
                //
                // 仍保留 2 秒延迟：窗口要先画出来，否则一键到底用户看不到
                // 界面出现的过程，会以为程序没启动。
                if (AutoStartTask)
                {
                    WriteLog("已启用\"启动时自动执行既定任务\"，2 秒后自动开始");
                    System.Windows.Forms.Timer delayTimer = new System.Windows.Forms.Timer();
                    delayTimer.Interval = 2000;
                    delayTimer.Tick += (s, args) =>
                    {
                        delayTimer.Stop();
                        WriteLog("开始启动同步进程（设置项自动触发）");
                        StartSyncProcess();
                    };
                    delayTimer.Start();
                }
                else
                {
                    WriteLog("未启用自动执行，等待用户点击\"开始执行\"");
                    UpdateStatus("已就绪，点击左下角\"开始执行\"", Color.FromArgb(30, 90, 160));
                    logTextBox.Clear();
                    logTextBox.AppendText($"PEACE & LOVE 川中计算机协会 陈叔叔希沃系统优化工具 ver.{_versionString}\n");
                    logTextBox.AppendText(new string('-', 50) + "\n");
                    logTextBox.AppendText("陈叔叔已在此守候。\n");
                    logTextBox.AppendText("点击左下角「开始执行」按钮，即可开始音量调节、结束 WPS 进程与时间同步。\n");
                }

                WriteLog("Load 事件完成");
            }
            catch (Exception ex)
            {
                WriteLog($"Load 事件异常：{ex}");
                MessageBox.Show($"加载窗体时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 首次运行时把当天写进设置，作为"已守护 x 天"的起点。
        ///
        /// 【为什么在这里写而不是在 SettingsStore.Load 里顺手写】
        /// Load 是个纯读取函数，它写注册表会让"读设置"带上副作用——
        /// 单测、诊断工具调它一次就改了用户数据。写入收敛到这一处显式调用。
        /// </summary>
        private void RecordFirstRunDateIfNeeded()
        {
            try
            {
                AppSettings settings = SettingsStore.Load();
                if (settings.FirstRunDate != DateTime.MinValue)
                {
                    _firstRunDate = settings.FirstRunDate;
                    return;
                }

                _firstRunDate = DateTime.Now.Date;
                settings.FirstRunDate = _firstRunDate;
                // 只补这一个字段，其余原样回写，避免把界面上还没加载的值覆盖掉
                settings.AutoStartTask = _autoStartTask;
                settings.ShutdownRules = _shutdownRules;
                SettingsStore.Save(settings);

                WriteLog($"首次运行，记录守护起始日期：{_firstRunDate:yyyy-MM-dd}");
            }
            catch (Exception ex)
            {
                WriteLog($"记录首次运行日期失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 刷新上半屏的"已守护本电脑 x 天"。
        ///
        /// 同步成功后会被再调一次：启动时系统时钟可能停在 2000 年，
        /// 那时算出来的天数是错的（可能是个天文数字或 1 天）。
        /// 时间校准之后再算一次，显示的才是真实天数。
        /// </summary>
        private void RefreshGuardDays()
        {
            int days = 1;
            try
            {
                AppSettings settings = SettingsStore.Load();
                DateTime first = settings.FirstRunDate != DateTime.MinValue
                    ? settings.FirstRunDate : DateTime.Now.Date;
                days = settings.GuardedDays(DateTime.Now);
            }
            catch
            {
                // 读设置失败也要有个可显示的数字，不能因为一次 I/O 异常
                // 让上半屏一片空白——那看起来像程序坏了
                days = 1;
            }

            string text = string.Format(
                "陈叔叔希沃优化助手\n已守护本电脑 {0} 天", days);

            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action(() => { guardLabel.Text = text; })); }
                catch { }
                return;
            }
            guardLabel.Text = text;
            WriteLog($"守护天数已刷新：第 {days} 天");
        }

        /// <summary>
        /// 启动关机巡检定时器。
        ///
        /// 【为什么常驻托盘期间也要一直跑】
        /// 用户可能在启动后就把窗口隐藏到托盘去上课了，而关机时间点
        /// 恰恰是在那之后才到。定时器挂在窗体上，与窗口是否可见无关——
        /// 只要进程活着、消息循环在转，它就会准时触发。
        /// </summary>
        private void StartShutdownWatcher()
        {
            if (shutdownTimer != null) return;

            if (_shutdownRules == null || _shutdownRules.Count == 0)
            {
                WriteLog("未配置关机时间点，不启动关机巡检");
                return;
            }

            shutdownTimer = new System.Windows.Forms.Timer();
            shutdownTimer.Interval = ShutdownService.PollIntervalMs;
            shutdownTimer.Tick += (s, e) => PollShutdown();
            shutdownTimer.Start();

            WriteLog($"关机巡检已启动，共 {_shutdownRules.Count} 条规则，"
                   + $"每 {ShutdownService.PollIntervalMs / 1000} 秒检查一次");
            foreach (ShutdownTimeRule r in _shutdownRules)
                WriteLog("  · " + r.Describe());
        }

        /// <summary>
        /// 一次关机巡检：提醒或执行。全程跑在 UI 线程（WinForms Timer 的特性）。
        /// </summary>
        private void PollShutdown()
        {
            try
            {
                // 规则为空时不做任何事——包括不能因为"上一轮有规则、
                // 这一轮用户全删了"而残留状态
                if (_shutdownRules == null || _shutdownRules.Count == 0) return;

                DateTime now = DateTime.Now;
                ShutdownService.PollResult result = ShutdownService.Poll(_shutdownRules, now);

                if (result.ShouldWarn)
                {
                    HandleShutdownWarning(result, now);
                }
                else if (result.ShouldShutdown)
                {
                    HandleShutdownExecution(result, now);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"关机巡检异常：{ex.Message}");
            }
        }

        /// <summary>
        /// 弹出"5 分钟后关机"提醒，两个选项：本次不关机 / 确认。
        ///
        /// 【为什么不用 MessageBox】
        /// MessageBox 的按钮文字由系统决定（中文 Windows 上渲染成"是(Y)/否(N)"），
        /// 无法自定义。用户明确要求按钮写「本次不关机」「确认」，
        /// 因此这里改用自绘对话框 ShutdownWarningDialog。
        ///
        /// 【去重与承诺已在 Poll 内完成】
        /// Poll 在产出 ShouldWarn 的同时就把"这一次关机已提醒"和"承诺时刻"
        /// 都记好了（必须在弹框之前定死——弹框是模态的，用户不点时
        /// UI 线程会一直阻塞在这里，不能依赖"弹框返回后再记账"）。
        /// 所以本方法不再需要 MarkWarned / PromiseShutdown。
        /// </summary>
        private void HandleShutdownWarning(ShutdownService.PollResult result, DateTime now)
        {
            WriteLog($"触发关机提醒：规则 {result.Rule.Describe()}，"
                   + $"预定于 {result.DueAt:HH:mm} 关机");

            DialogResult choice;
            using (var dlg = new ShutdownWarningDialog(
                (int)ShutdownScheduleLogic.WarnMinutesAhead, result.DueAt))
            {
                choice = dlg.ShowDialog(this);
            }

            if (choice == DialogResult.Yes)
            {
                // 用户点了"确认"：承诺时刻已在 Poll 内写入，这里什么都不用做——
                // 到点由 PollShutdown 直接执行，不再二次弹框。
                WriteLog("用户在提醒中选择\"确认\"，到点将直接关机");
            }
            else
            {
                // SkipOnce 内部会连带撤销任何已下发的系统关机倒计时
                // （见其注释：只改内存标记挡不住已经发出去的关机命令）。
                ShutdownService.SkipOnce(result.DueAt);
                WriteLog("用户在提醒中选择\"本次不关机\"，已跳过本次并尝试撤销系统关机");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "已跳过本次关机",
                        $"本次 {result.DueAt:HH:mm} 不再关机，下次仍会提醒。", ToolTipIcon.Info);
                }
            }
        }

        /// <summary>
        /// 到点执行关机。
        ///
        /// 【这里不弹任何确认框——提醒框（提前 5 分钟）是唯一的确认窗口】
        /// 用户没点"本次不关机"就等于默许，到点就干脆地关掉。
        /// 到点再问一次会让"配了 23:00 却要点两次才关"变成常态，
        /// 而且无人值守的教室一体机上，那个框会一直挂着，关机永远不执行。
        /// 真正需要撤销时，用户应该在提前 5 分钟的那个框里点"本次不关机"。
        ///
        /// 【为什么不再带系统倒计时】
        /// 见 ShutdownService.SystemCountdownSeconds 的注释：
        /// 5 分钟系统倒计时结束时会由 Windows 自己弹出"您将要被注销"的框，
        /// 与我们的提醒框叠在一起，用户无所适从。
        /// 现在改为 /t 0 立即执行，界面上只留我们自己的提示。
        ///
        /// result.Rule 在"重启补弹后顺延"的情形下可能为 null
        /// （承诺时刻已不等于任何规则时刻），文案需容忍。
        /// </summary>
        private void HandleShutdownExecution(ShutdownService.PollResult result, DateTime now)
        {
            ShutdownService.MarkShutdown(now);

            string label = result.Rule != null
                ? result.Rule.Describe()
                : $"{result.DueAt:HH:mm}";
            WriteLog($"到达关机时间点：{label}，执行关机");

            string error;
            if (ShutdownService.ExecuteShutdown(out error))
            {
                AddLog("\n× 已执行关机命令\n", Color.DarkRed);
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "正在关机",
                        "已到达预定时间，电脑正在关机。", ToolTipIcon.Warning);
                }
            }
            else
            {
                WriteLog($"关机命令执行失败：{error}");
                MessageBox.Show("无法执行关机命令：" + error, "关机失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            // 菜单里的"退出"是明确的退出意图，置位后 FormClosing 才放行
            _exitingProgram = true;

            if (isSyncing)
            {
                DialogResult result = MessageBox.Show("同步正在进行中，宁真的要强制退出吗？",
                    "你真的要退出吗？", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    // 用户反悔：必须把标志撤回，否则此后点 × 会真的关掉程序，
                    // 而用户以为自己只是取消了这次退出
                    _exitingProgram = false;
                    return;
                }

                CancelSync();
                syncCompleted = true;

                if (syncThread != null && syncThread.IsAlive)
                {
                    if (!syncThread.Join(3000))
                        WriteLog("等待同步线程退出超时");
                }
            }
            else
            {
                CancelSync();
                syncCompleted = true;
            }

            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }

            // 走 Program.Shutdown 而不是 Application.Exit：
            // Shutdown 会先等后台线程收尾、再写日志结束标记，保证
            // "同步线程正常结束"之类的收尾日志压在分隔线之内。
            // 直接 Application.Exit 会让那些日志落到结束标记之后。
            Program.Shutdown(0, "用户从菜单退出");
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

            ToolStripMenuItem syncItem = new ToolStripMenuItem("立即同步时间");
            syncItem.Click += (s, e) =>
            {
                if (!isSyncing) StartSyncProcess();
            };
            trayMenu.Items.Add(syncItem);

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
            // 【常驻模式下的关窗语义变了】
            //
            // 原逻辑：只在 syncCompleted（任务完成）或同步被取消/权限错误时
            // 才真正关闭窗口，其余情况一律取消关闭、缩到托盘。
            //
            // 现在程序要长期常驻守着关机时间点，所以"点右上角的 ×"这个动作
            // 应当**永远**是"缩到托盘"——窗口关掉但程序继续工作。
            // 真正退出只有一条路：托盘右键菜单里的"退出程序"
            // （或菜单栏 选项→退出），它们会先设 _exitingProgram = true。
            //
            // 【为什么用独立标志而不是复用 syncCompleted】
            // 若复用，就得在"任务完成"时置位，而那会让用户此后每次点 ×
            // 都直接把程序关掉——偏偏那正是最需要它继续守着的时刻。
            if (_exitingProgram)
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

            e.Cancel = true;
            MinimizeToTray();
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
                // 【文案变更】原来说"不必关闭"——那是"同步完就退出"时代的说法。
                // 现在程序是长期常驻的，用户有权知道它一直在这里。
                trayIcon.ShowBalloonTip(3000, "陈叔叔希沃优化助手",
                    "程序已驻留托盘，将持续守候关机时间点。右键图标可立即同步或退出。",
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
            // 【常驻模式下必须把退出的后果讲清楚】
            // 用户可能没意识到"退出"意味着关机守候也没了。
            // 若配了关机时间点，提示里明确说出来——否则会静默失效，
            // 而失效的那天正好是需要它关机的日子。
            string extra = "";
            if (_shutdownRules != null && _shutdownRules.Count > 0)
            {
                extra = "\n\n注意：退出后关机时间点将不再生效。\n" +
                        "（已配置 " + _shutdownRules.Count + " 个关机时间点）";
            }

            DialogResult result = MessageBox.Show(
                "宁确定要退出系统优化工具吗？" + extra,
                "确认退出", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.Yes)
            {
                _exitingProgram = true;
                syncCompleted = true;
                if (trayIcon != null)
                {
                    trayIcon.Visible = false;
                    trayIcon.Dispose();
                }
                Program.Shutdown(0, "用户从托盘退出");
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
                        AddLog($"\n和时间的同步率达到100%！\n", Color.DarkGreen);

                        if (trayIcon != null)
                            trayIcon.ShowBalloonTip(3000, "时间同步完成", "系统时间已成功同步！", ToolTipIcon.Info);

                        // 【需求 3：时间同步成功后再确认今天是第几天】
                        // 启动时那次刷新用的是当时（可能是 2000 年）的系统时间，
                        // 算出来的天数不对。此刻系统时钟刚刚被校正，必须重算一次。
                        RefreshGuardDays();

                        // 必须在下面检查 UpdateCheckCompleted 之前重试。
                        // 本工具的使用场景就是"时钟错误"，而时钟错误会让
                        // 启动时那次更新检查必然因证书未生效而失败。
                        // 若不重置标志，UI 会读到上一轮残留的 true 直接退出，
                        // 用户永远收不到更新——开机时钟总是不对就永久失效。
                        Program.RetryUpdateCheckAfterSync();

                        if (WaitOrCancel(2000)) return;

                        // ------------------------------------------------------------------
                        // 【本版最重要的行为变更：同步成功后不再自动退出】
                        //
                        // 原实现在更新检查完成后调 Application.Exit()，理由是"任务已完成"。
                        // 但现在程序还担着另一件事——盯着关机时间点。一旦退出，
                        // 那个功能就死了：用户设了 23:00 关机，而程序 8:30 就自己退了。
                        //
                        // 所以改为：同步完成 → 隐藏到托盘 → 常驻待命。
                        // 真正退出只有两条路：用户在菜单里点"退出"，
                        // 或系统关机时进程被终止。
                        // ------------------------------------------------------------------
                        this.Invoke(new MethodInvoker(() =>
                        {
                            MinimizeToTray();

                            if (Program.UpdateCheckCompleted)
                            {
                                WriteLog("更新检查已完成，程序转入常驻托盘待命");
                                AddLog("\n√ 优化完成。程序已在托盘常驻，"
                                     + "将持续守候关机时间点。\n", Color.DarkGreen);
                                AddLog("如需退出，请右键托盘图标选择\"退出程序\"。\n", Color.DarkGray);
                                SetSyncFinishedUi();
                                return;
                            }

                            // 兜底上限：更新检查的网络请求若因断网等原因一直不返回，
                            // 没有这个上限会一直等下去。但注意——这里**不再退出**，
                            // 只是停止等待、宣告同步阶段结束，程序照样常驻。
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
                                        WriteLog($"等待更新检查超过 {UPDATE_CHECK_WAIT_LIMIT_MS / 1000} 秒，"
                                               + "不再等待；程序继续常驻待命");

                                    AddLog("\n√ 优化完成。程序已在托盘常驻，"
                                         + "将持续守候关机时间点。\n", Color.DarkGreen);
                                    AddLog("如需退出，请右键托盘图标选择\"退出程序\"。\n", Color.DarkGray);
                                    SetSyncFinishedUi();
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

                        DialogResult result = MessageBox.Show(
                            "需要管理员权限才能修改系统时间。\n\n" +
                            "是否以管理员身份重新启动程序？\n" +
                            "（重启后关机时间点仍会继续生效）",
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

                            // 重启流程本身就是"结束本进程、换一个提权的进程"，
                            // 必须置位，否则 Program.Shutdown 走到的 FormClosing
                            // 会被"取消关闭、缩到托盘"拦下来，旧进程留在托盘，
                            // 与新启动的进程争抢单实例互斥体。
                            _exitingProgram = true;
                            Program.Shutdown(0, "以管理员权限重新启动");
                        }

                        // 用户选择不重启：留在常驻状态，等他自己处理权限问题
                        SetSyncFinishedUi();
                        isPermissionError = true;
                    }
                    else
                    {
                        WriteLog("同步失败");
                        UpdateStatus("同步失败", Color.Red);
                        AddLog("× 同步失败\n", Color.Red);
                        AddLog($"\n（悲报！） 经过 {NtpServers.Length} 个服务器的尝试后，时间同步失败。\n", Color.DarkRed);

                        // 【行为变更】原实现在这里等 5 秒后关闭窗口退出程序。
                        // 现在程序要常驻守着关机时间点，所以失败也不退出——
                        // 只把界面恢复成"可以重试"的状态，让用户自己决定。
                        WaitOrCancel(FINAL_FAILURE_DELAY_MS);
                        AddLog("\n可以在网络恢复后点击「重新同步」再试一次。\n", Color.DarkOrange);
                        AddLog("程序将常驻托盘继续守候关机时间点。\n", Color.DarkGray);
                        SetSyncFinishedUi();
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

                    // 【行为变更】原实现 this.Close() 直接关窗退出。
                    // 常驻模式下异常也不能把程序带走——一次同步失败不等于
                    // 这个守护程序就该消失，用户可能正是靠它守着关机时间点。
                    SetSyncFinishedUi();
                }
                finally
                {
                    WriteLog("同步线程 finally 块");

                    // 无论走哪条分支，都要把界面恢复成"可以再次点开始"。
                    //
                    // 【为什么用 BeginInvoke 而不是同步 Invoke】
                    // 退出时主线程会在 WaitForBackgroundWork 里 Join 本线程，
                    // 而此刻消息循环可能已经停止。同步 Invoke 会一直等 UI 线程
                    // 来取这个工作项，但 UI 线程正卡在 Join 上——双方互锁，
                    // 只能等 Join 超时才解开。改用 BeginInvoke 投递后立即返回。
                    //
                    // isSyncing 的复位放在 SetSyncFinishedUi 里，两条路径共用，
                    // 避免"成功那条复位了、异常那条忘了"这种漏网。
                    if (this.IsHandleCreated && !this.IsDisposed)
                    {
                        try
                        {
                            this.BeginInvoke(new MethodInvoker(() =>
                            {
                                try { UpdateProgressBar(false); } catch { }
                                try { SetSyncFinishedUi(); } catch { }
                            }));
                        }
                        catch { }
                    }

                    // 双保险：即使 UI 线程已经取不到（进程正在退出），
                    // 这个标志也必须是 false，否则线程内的后续判断会误判
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
            // 【窗口减半后进度条被移除了】原实现用一个 Marquee 风格的长条
            // 表示"正在忙"，它占掉约 25px 高度，而窗口减半后这点高度
            // 要留给日志区。现在"是否在忙"由 statusLabel 的文字和
            // "开始执行"按钮的可用状态共同表达，信息量不减。
            //
            // 这里保留成空方法而不是删掉调用点：UpdateProgressBar 在
            // 同步线程里有 4 处调用，删掉容易漏；留一个明确说明原因的空实现，
            // 比让后来者在多处判断"这里该不该调"更清楚。
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

            // 同步进行中：两个按钮都禁用，避免重复点。
            // "开始执行"与"重新同步"做的是同一件事（StartSyncProcess 内部
            // 有 isSyncing 兜底），但按钮上的字不一样，用户会以为能做两件事，
            // 所以必须一起灰掉。
            startButton.Enabled = enabled;
            retryButton.Enabled = enabled;
        }

        /// <summary>
        /// 同步阶段结束（成功或失败）后的界面收尾。
        ///
        /// 【为什么单独抽出来】原实现把这段散落在同步线程里，且结尾都是
        /// "关闭窗口 + Application.Exit"。本版改成常驻后，两条路径
        /// （成功 / 失败）都收敛到这里，只有一处决定界面回到什么状态。
        /// 抽出来也顺便保证 isSyncing 一定被复位——散着写就容易漏掉某条分支。
        /// </summary>
        private void SetSyncFinishedUi()
        {
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new MethodInvoker(SetSyncFinishedUi)); }
                catch { }
                return;
            }

            isSyncing = false;
            UpdateButton(true);

            // 注意：这里**不设 syncCompleted = true**。
            // syncCompleted 的语义是"本次运行已完成、允许窗口关闭"，
            // 而常驻模式下的窗口是随时可以关的（关掉只是隐藏到托盘）。
            // 若在这里置 true，FormClosing 会跳过"取消关闭、最小化到托盘"
            // 那段逻辑，用户点关闭按钮时窗口会真的消失而不是缩进托盘。
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
            _autoStartTask = settings.AutoStartTask;
            _firstRunDate = settings.FirstRunDate;
            _shutdownRules = settings.ShutdownRules ?? new List<ShutdownTimeRule>();
        }

        public void SaveSettings()
        {
            try
            {
                // 【必须先读回再改】FirstRunDate 是启动时写入、之后只读的字段。
                // 若这里用 new AppSettings 重新构造，FirstRunDate 会是 MinValue，
                // 而 Save 里那条"仅在未记录时写入"的保护会让注册表里的
                // 旧值再也写不回去——本次运行之后的守护天数就全错了。
                AppSettings settings = SettingsStore.Load();

                settings.AutoStart = _autoStart;
                settings.SilentStart = _silentStart;
                settings.AutoVolume = _autoVolume;
                settings.VolumeLevel = _volumeLevel;
                settings.KillWps = _killWps;
                settings.AutoStartTask = _autoStartTask;
                settings.ShutdownRules = _shutdownRules;
                if (settings.FirstRunDate == DateTime.MinValue)
                    settings.FirstRunDate = _firstRunDate;

                SettingsStore.Save(settings);

                UpdateAutoStartTask();

                // 关机设置可能刚被改动，立刻重排巡检——
                // 否则新的时间点要等到下一个轮询周期才生效
                RestartShutdownWatcher();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 设置变更后重启关机巡检：清空记账状态，按新的规则重新开始。
        ///
        /// 【为什么要清记账】若用户把一个时间点从 23:00 改成 23:05，
        /// 而程序刚刚已经在 23:00 提醒过并记了账，不清空的话
        /// "已提醒过的分钟"会残留，可能把新的提醒点一起吞掉。
        /// </summary>
        private void RestartShutdownWatcher()
        {
            ShutdownService.ResetRuntimeState();

            if (shutdownTimer != null)
            {
                shutdownTimer.Stop();
                shutdownTimer.Dispose();
                shutdownTimer = null;
            }
            StartShutdownWatcher();
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
                // （同样要置位退出标志，否则会被常驻的"关窗即缩托盘"拦下）
                _exitingProgram = true;
                syncCompleted = true;
                Program.Shutdown(0, "卸载完成，程序退出");
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