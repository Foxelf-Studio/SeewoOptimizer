using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SeewoOpt.Services;

namespace SeewoOpt
{
    static class Program
    {
        // 日志路径统一定义在 LogService，此处仅做转发以兼容既有调用点
        public static readonly string LogFilePath = LogService.LogFilePath;

        // GitHub 配置
        private const string GITHUB_API = "https://api.github.com/repos/Foxelf-Studio/SeewoOptimizer/releases/latest";
        private const string GITHUB_USER_AGENT = "SeewoOptimizer-Updater";
        private static readonly string UPDATE_DIR = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool", "updates");

        // 更新完成标志（由后台更新线程写、UI 线程读，必须 volatile，
        // 否则 UI 侧 Timer 可能读到缓存值导致窗口一直卡在托盘不退出）
        public static volatile bool UpdateCheckCompleted = false;

        // 本次运行中更新检查是否失败过。
        //
        // 【为什么需要它】本工具的核心场景就是"本机时钟错误"，而时钟错误
        // 会让每一次 HTTPS 都因证书 notBefore 尚未到达而失败。
        // 时序上更新检查先于时间同步启动（Main 里就 Task.Run 了），
        // 于是每次开机时钟都不对时，更新检查必然失败，
        // 而 finally 又会把 UpdateCheckCompleted 置 true，
        // 同步成功后 UI 看到"已完成"就直接退出——自动更新永久失效。
        // 时间同步成功后据此重试，把失效的更新能力救回来。
        private static volatile bool _updateCheckFailed = false;

        // 更新检查是否正在运行。用 Interlocked 做无锁保护，
        // 防止首次检查尚未收尾时又发起一次重试。
        private static int _updateCheckRunning = 0;

        // 更新检测事件（用于通知 TimeSyncForm 显示气泡）
        public static event Action<string, string> UpdateDetected;

        // 缓存更新信息，防止事件错过
        public static Tuple<string, string> PendingUpdateInfo { get; set; }

        // 更新文件的预期 SHA256 哈希（从发布说明中提取）
        private static string _expectedUpdateHash = null;

        // 互斥体对象
        private static Mutex singleInstanceMutex;

        /// <summary>
        /// 本次运行的结束原因，写进日志分段的结尾标记。
        /// 各退出路径设置它，finally 里统一读取——这样不必在每个 return 前
        /// 重复写收尾日志，也不会漏掉某条退出路径。
        /// </summary>
        private static string _exitReason = "未标明";

        [STAThread]
        static void Main()
        {
            try
            {
                // 开始本次运行的日志分段：裁剪旧运行 + 写分隔头。
                // 放在最前面（早于单实例检查）是为了让每一次 Main 进入都
                // 有完整的一段——否则"已有实例运行"那条退出路径只会在
                // 上一次的日志段末尾甩一个孤立的结束标记。
                LogService.BeginSession();

                // 检查是否有待应用的更新（更新后重启）- 只处理不弹窗
                HandlePendingUpdate();

                // 检查是否已有实例在运行
                if (!IsSingleInstance())
                {
                    _exitReason = "已有实例在运行";
                    WriteLog("已有实例运行，退出");
                    MessageBox.Show("程序已在运行中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 程序集解析事件（必须在加载任何外部程序集前注册）
                AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

        // 确保必需的 DLL 存在（必须在异步更新检查之前，否则任务计划库加载失败）
        EnsureRequiredDllsExist();

                // 异步检查新版本（不阻塞主线程）
                StartUpdateCheck("启动时");

                WriteLog(BuildInfo.Describe());
                WriteLog($"命令行参数: {string.Join(" ", Environment.GetCommandLineArgs())}");
                WriteLog($"当前目录: {Environment.CurrentDirectory}");
                WriteLog($"程序路径: {Application.ExecutablePath}");
                WriteLog($"程序目录: {Path.GetDirectoryName(Application.ExecutablePath)}");
                WriteLog($"当前版本: {Assembly.GetExecutingAssembly().GetName().Version}");

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Application.ThreadException += (sender, e) =>
                {
                    WriteLog($"========== 线程异常 ==========");
                    WriteLog($"异常消息: {e.Exception.Message}");
                    WriteLog($"异常堆栈: {e.Exception.StackTrace}");
                    MessageBox.Show($"发生未处理的异常：{e.Exception.Message}\n\n程序将关闭。",
                        "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };

                AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
                {
                    Exception ex = e.ExceptionObject as Exception;
                    WriteLog($"========== 致命异常 ==========");
                    WriteLog($"异常消息: {ex?.Message ?? "未知错误"}");
                    WriteLog($"异常堆栈: {ex?.StackTrace ?? ""}");
                    MessageBox.Show($"发生致命错误：{ex?.Message ?? "未知错误"}\n\n程序将关闭。",
                        "致命错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };

                try
                {
                    WriteLog("开始运行主窗体");
                    Application.Run(new TimeSyncForm());
                    WriteLog("主窗体运行结束");
                    _exitReason = "用户退出（主窗体关闭）";
                }
                catch (Exception ex)
                {
                    WriteLog($"========== Run 异常 ==========");
                    WriteLog($"异常消息: {ex.Message}");
                    _exitReason = $"主窗体异常：{ex.Message}";
                    MessageBox.Show($"程序运行错误：{ex.Message}", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                WriteLog("程序正常退出");
            }
            catch (Exception ex)
            {
                WriteLog($"========== Main 顶级异常 ==========");
                WriteLog($"异常类型: {ex.GetType().Name}");
                WriteLog($"异常消息: {ex.Message}");
                WriteLog($"异常堆栈: {ex.StackTrace}");
                _exitReason = $"启动阶段异常（{ex.GetType().Name}）：{ex.Message}";
                MessageBox.Show($"程序启动时发生致命错误：{ex.Message}\n\n程序将关闭。",
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 【顺序】必须先释放互斥体，再写结束标记。
                // ReleaseMutex 自身会写一条"互斥体已释放"日志，若放在
                // EndSession 之后，这条日志就会落在"======== RUN 结束 ========"
                // 之下、跑到本次运行的分段之外——看上去像属于下一次运行，
                // 实则不是。凡是本次运行产生的日志都应落在结束标记之前。
                ReleaseMutex();

                // 收尾日志分段。放在 finally 里是因为它是唯一保证会执行的路径：
                // 正常退出、Main 抛异常、Run 抛异常都会走到这里。
                // 被强杀（任务管理器结束进程、断电）时不会执行——那种情况
                // 日志里就没有结束标记，下次启动一眼能看出上次是异常终止的。
                // 重复调用是安全的：LogService.EndSession 自身幂等。
                LogService.EndSession(_exitReason);
            }
        }

        /// <summary>
        /// 检查是否只有一个实例在运行
        /// </summary>
        private static bool IsSingleInstance()
        {
            try
            {
                // 尝试创建互斥体
                singleInstanceMutex = new Mutex(true, "TimeSyncTool_UniqueMutex", out bool createdNew);
                return createdNew;
            }
            catch (Exception ex)
            {
                WriteLog($"检查实例时出错：{ex.Message}");
                return true;
            }
        }

        /// <summary>
        /// 释放互斥体
        /// </summary>
        private static void ReleaseMutex()
        {
            try
            {
                if (singleInstanceMutex != null)
                {
                    singleInstanceMutex.ReleaseMutex();
                    singleInstanceMutex.Close();
                    singleInstanceMutex = null;
                    WriteLog("互斥体已释放");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"释放互斥体时出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 检查是否有待应用的更新（更新后重启）- 静默处理，不弹窗
        /// </summary>
        private static void HandlePendingUpdate()
        {
            try
            {
                string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
                if (File.Exists(pendingFile))
                {
                    WriteLog("检测到待应用的更新");

                    string newVersionPath = File.ReadAllText(pendingFile).Trim();
                    string currentExe = Application.ExecutablePath;

                    if (File.Exists(newVersionPath))
                    {
                        // 注意：这里不替换文件，因为已经在更新脚本中完成了
                        // 只清理标志文件
                        File.Delete(pendingFile);
                        WriteLog("更新标志已清除");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"处理待更新文件失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 发起一次更新检查。统一入口，便于在时间同步成功后重试。
        ///
        /// 【重试的必要性】本工具用于修复错误时钟，而错误时钟会让 HTTPS
        /// 证书校验必然失败。启动时的检查发生在同步之前，因此每次开机
        /// 时钟都不对时，这次检查必定失败并被标记为"已完成"，
        /// 程序随后退出，用户永远收不到更新——除非在这里重试。
        /// </summary>
        public static void StartUpdateCheck(string reason)
        {
            // 防止首次检查尚未收尾时重复发起
            if (Interlocked.CompareExchange(ref _updateCheckRunning, 1, 0) != 0)
            {
                WriteLog($"更新检查已在进行中，跳过本次触发（{reason}）");
                return;
            }

            WriteLog($"触发更新检查（{reason}）");

            Task.Run(() => CheckForUpdatesAsync()).ContinueWith(t =>
            {
                Interlocked.Exchange(ref _updateCheckRunning, 0);

                if (t.IsFaulted)
                {
                    _updateCheckFailed = true;
                    WriteLog($"更新检查任务异常: {t.Exception?.Message}");
                    WriteLog($"异常详情: {t.Exception}");
                }
                else
                {
                    WriteLog("更新检查任务正常完成");
                }
            });
        }

        /// <summary>
        /// 时间同步成功后调用，重试此前失败的更新检查。
        ///
        /// 只有确实失败过才重试——正常路径下首次检查已经拿到结果，
        /// 再跑一遍纯属浪费请求，也会让 UI 多等一轮。
        /// </summary>
        public static void RetryUpdateCheckAfterSync()
        {
            if (!_updateCheckFailed)
            {
                WriteLog("首次更新检查成功，无需重试");
                return;
            }

            WriteLog("时间已修正，重试更新检查");
            _updateCheckFailed = false;

            // 关键：重置完成标志，否则 TimeSyncForm 会因读到上一轮的 true
            // 而直接 Application.Exit()，重试还没开始程序就退没了。
            UpdateCheckCompleted = false;
            StartUpdateCheck("时间同步后重试");
        }

        /// <summary>
        /// 异步检查 GitHub 上的新版本
        /// </summary>
        private static async Task CheckForUpdatesAsync()
        {
            WriteLog("进入 CheckForUpdatesAsync 方法");
            try
            {
                // 强制使用 TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                WriteLog("开始后台检查更新...");

                Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                WriteLog($"当前版本: {currentVersion}");

                if (!Directory.Exists(UPDATE_DIR))
                    Directory.CreateDirectory(UPDATE_DIR);

                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("User-Agent", GITHUB_USER_AGENT);
                    client.Headers.Add("Accept", "application/vnd.github.v3+json");
                    client.Encoding = Encoding.UTF8;

                    string jsonResponse = await client.DownloadStringTaskAsync(GITHUB_API);
                    WriteLog("GitHub API 响应成功");
                    WriteLog($"响应内容长度: {jsonResponse.Length}");

                    // 用框架自带的 JSON 序列化器解析，正确处理转义与多资产场景
                    if (!ReleaseInfoParser.TryParse(jsonResponse, out ReleaseInfo release, out string parseError))
                    {
                        WriteLog($"解析 Release 信息失败: {parseError}");
                        return;
                    }

                    WriteLog($"提取到 tag_name: {release.TagName}");

                    // 从资产列表中挑选可执行文件——不再简单取第一个
                    ReleaseAsset asset = ReleaseInfoParser.SelectExecutableAsset(release.Assets);
                    if (asset == null)
                    {
                        WriteLog("未在 Release 资产中找到可执行文件（.exe/.msi）");
                        return;
                    }

                    string downloadUrl = asset.BrowserDownloadUrl;
                    WriteLog($"选定资产: {asset.Name}（{asset.Size} 字节）");
                    WriteLog($"下载地址: {downloadUrl}");

                    // 从发布说明（body）中提取预期的 SHA256 哈希
                    _expectedUpdateHash = ReleaseInfoParser.TryExtractExpectedSha256(release.Body);
                    if (_expectedUpdateHash != null)
                    {
                        WriteLog($"从发布说明中提取到 SHA256: {_expectedUpdateHash}");
                    }
                    else
                    {
                        WriteLog("发布说明中未找到 SHA256 哈希，将仅依赖签名校验");
                    }

                    // 从资产名提取文件名（比从 URL 解析更可靠）
                    string fileName = asset.Name;
                    WriteLog($"提取到文件名: {fileName}");

                    Version latestVersion = ReleaseInfoParser.ParseVersion(release.TagName);
                    if (latestVersion == null)
                    {
                        WriteLog($"无法解析版本号: {release.TagName}，跳过本次更新检查");
                        return;
                    }

                    WriteLog($"GitHub 最新版本: {latestVersion}");

                    if (latestVersion.CompareTo(currentVersion) > 0)
                    {
                        WriteLog("发现新版本，准备自动更新");

                        // 缓存更新信息
                        PendingUpdateInfo = Tuple.Create(latestVersion.ToString(), "正在自动更新...");
                        UpdateDetected?.Invoke(latestVersion.ToString(), "正在自动更新...");

                        // 带重试的下载
                        await DownloadUpdateWithRetryAsync(downloadUrl, fileName, latestVersion, _expectedUpdateHash);
                    }
                    else
                    {
                        WriteLog("已是最新版本");
                    }
                }
            }
            catch (Exception ex)
            {
                // 无论何种原因失败都要置位：时间同步成功后会重试，
                // 而重试是恢复"时钟错误导致更新永久失效"的唯一机会。
                _updateCheckFailed = true;

                // 本机时钟严重错误时，几乎所有 HTTPS 都会以"证书无效"失败：
                // 证书的 notBefore/notAfter 是绝对时间，1980 年的本机时钟会让
                // 每一张证书都变成"尚未生效"。这不是网络故障，但日志里表现为
                // trust relationship 失败，极易被误判成网络问题而查错方向。
                // 这里识别出来并说明真实原因——时间同步成功后会重试。
                if (IsCertificateTimeFailure(ex))
                {
                    WriteLog("更新检查失败：本机时间与证书有效期不匹配，" +
                             $"当前 {DateTime.Now:yyyy-MM-dd}，HTTPS 证书校验必然失败。" +
                             "时间同步成功后会自动重试。");
                }
                else
                {
                    WriteLog($"后台检查更新失败：{ex.Message}");
                    WriteLog($"异常详情：{ex}");
                }
            }
            finally
            {
                UpdateCheckCompleted = true;
                WriteLog("更新检查完成");
            }
        }

        /// <summary>
        /// 判断异常是否源于"本机时钟错误导致证书不在有效期内"。
        ///
        /// 这类失败的特征是异常链里同时出现 TLS 握手失败与证书校验失败，
        /// 而底层往往是同一个"remote certificate is invalid"。
        /// </summary>
        private static bool IsCertificateTimeFailure(Exception ex)
        {
            for (Exception e = ex; e != null; e = e.InnerException)
            {
                string msg = e.Message ?? string.Empty;
                if (msg.IndexOf("certificate", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("trust relationship", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 带重试机制的下载更新文件
        /// </summary>
        private static async Task<bool> DownloadUpdateWithRetryAsync(string downloadUrl, string fileName, Version newVersion, string expectedHash = null)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 2000;

            for (int retry = 0; retry < maxRetries; retry++)
            {
                try
                {
                    await DownloadUpdateAsync(downloadUrl, fileName, newVersion, expectedHash);
                    return true; // 成功
                }
                catch (Exception ex)
                {
                    WriteLog($"下载失败 (尝试 {retry + 1}/{maxRetries}): {ex.Message}");
                    if (retry == maxRetries - 1)
                    {
                        WriteLog($"下载最终失败，无法更新");
                        return false;
                    }
                    else
                    {
                        WriteLog($"等待 {retryDelayMs}ms 后重试...");
                        await Task.Delay(retryDelayMs);
                    }
                }
            }
            return false;
        }

       
        /// <summary>
        /// 异步下载更新文件并验证签名（内部不捕获异常，由重试方法处理）
        /// </summary>
        private static async Task DownloadUpdateAsync(string downloadUrl, string fileName, Version newVersion, string expectedHash = null)
        {
            string downloadedFile = Path.Combine(UPDATE_DIR, fileName);
            string signatureFile = SignatureVerifier.GetSignaturePath(downloadedFile);
            string currentExe = Application.ExecutablePath;
            string updaterScript = Path.Combine(UPDATE_DIR, "update.bat");

            WriteLog("开始下载更新文件...");

            using (WebClient client = new WebClient())
            {
                // 使用 CancellationTokenSource 实现超时
                using (var cts = new CancellationTokenSource())
                {
                    cts.CancelAfter(TimeSpan.FromSeconds(60)); // 60秒超时

                    client.DownloadProgressChanged += (s, e) =>
                    {
                        try
                        {
                            if (e.ProgressPercentage % 10 == 0)
                                WriteLog($"下载进度: {e.ProgressPercentage}%");
                        }
                        catch { }
                    };

                    var downloadTask = client.DownloadFileTaskAsync(new Uri(downloadUrl), downloadedFile);
                    var completedTask = await Task.WhenAny(downloadTask, Task.Delay(-1, cts.Token));

                    if (completedTask != downloadTask)
                    {
                        // 超时
                        client.CancelAsync();
                        throw new TimeoutException("下载超时（60秒）");
                    }

                    await downloadTask; // 重新抛出可能的异常
                }
            }

            WriteLog("更新文件下载完成，下载配套签名...");

            // 签名文件与 exe 同名同目录，仅后缀不同
            string signatureUrl = downloadUrl + SignatureVerifier.SIGNATURE_EXTENSION;
            try
            {
                using (WebClient client = new WebClient())
                {
                    await client.DownloadFileTaskAsync(new Uri(signatureUrl), signatureFile);
                }
                WriteLog("签名文件下载完成");
            }
            catch (Exception ex)
            {
                TryDelete(downloadedFile);
                throw new Exception(
                    "未找到更新文件的签名（应为 " + fileName + SignatureVerifier.SIGNATURE_EXTENSION +
                    "）。\n请到发布页下载完整更新包。\n原因: " + ex.Message);
            }

            WriteLog("开始验证签名...");

            // 验证文件是否为有效的 PE 可执行文件
            if (!IsValidPE(downloadedFile))
            {
                TryDelete(downloadedFile);
                throw new Exception("下载的文件不是有效的 Windows 可执行文件（PE 头验证失败），更新已取消");
            }

            // 第一道也是唯一能抵御"仓库被入侵"的关卡：RSA 发布签名。
            // 私钥只存在于发布者本机，攻击者即使拿到 GitHub 账号完全控制权，
            // 没有私钥也造不出能通过验证的更新。
            if (!SignatureVerifier.IsAcceptable(downloadedFile, out string signatureInfo))
            {
                TryDelete(downloadedFile);
                TryDelete(signatureFile);
                throw new Exception("下载的文件未通过签名校验，更新已取消。\n" + signatureInfo);
            }
            WriteLog($"签名校验: {signatureInfo}");

            // SHA256 仅作记录，便于排查下载损坏问题。
            // 注意：它不能作为安全依据——期望哈希与文件来自同一 HTTP 响应，
            // 攻击者能同时替换两者。真正的信任锚是上面的签名验证。
            string actualHash = ComputeSHA256(downloadedFile);
            WriteLog($"下载文件 SHA256: {actualHash}");
            if (!string.IsNullOrEmpty(expectedHash))
            {
                bool matched = string.Equals(actualHash, expectedHash, StringComparison.OrdinalIgnoreCase);
                WriteLog(matched
                    ? "哈希与发布说明一致"
                    : $"哈希与发布说明不一致（说明: {expectedHash}）——签名有效，不影响更新");
            }

            CreateUpdateScript(updaterScript, currentExe, downloadedFile, newVersion);
            WriteLog("更新脚本创建成功，准备退出主程序，由更新脚本完成后台替换");

            // Environment.Exit 不会执行 finally，必须在此显式释放互斥体，
            // 否则批处理脚本里的 taskkill 可能杀不掉本进程，导致文件被占用而替换失败
            ReleaseMutex();
            Thread.Sleep(200);

            ProcessStartInfo updateProcess = CreateUpdateScript(updaterScript, currentExe, downloadedFile, newVersion);
            Process.Start(updateProcess);

            // 退出当前程序，让更新脚本替换文件
            Shutdown(0, $"准备自动更新到 {newVersion}，交由更新脚本替换文件");
        }

        /// <summary>
        /// 统一的退出入口。
        ///
        /// 直接调用 Environment.Exit 会跳过 Main 的 finally，导致单实例互斥体
        /// 不被释放——在自更新场景下，批处理脚本 taskkill 杀不掉残留进程，
        /// exe 文件被占用而替换失败。这里统一先释放资源再退出。
        ///
        /// 同理，日志的结束标记也必须在此显式写入，不能只依赖 Main 的 finally。
        /// </summary>
        public static void Shutdown(int exitCode)
        {
            Shutdown(exitCode, null);
        }

        /// <summary>带结束原因的退出入口</summary>
        public static void Shutdown(int exitCode, string reason)
        {
            if (!string.IsNullOrEmpty(reason))
                _exitReason = reason;

            // 【顺序】先释放互斥体，再写结束标记——因为 ReleaseMutex 自己会写
            // 一条"互斥体已释放"日志，若放在后面就会跑到分隔之外。
            //
            // 这里能安全地把 ReleaseMutex 提前，是因为它内部会先
            // singleInstanceMutex.Close() 把句柄交还系统；Windows 在进程终止时
            // 也会自动放弃未释放的互斥体。也就是说，即便紧随其后的
            // Environment.Exit 把后续语句截断，单实例保护仍然成立。
            ReleaseMutex();

            // 结束标记最后写：它是"本次运行到此为止"的分界线，必须压在
            // 所有本次运行的日志之下，所以由它收尾。
            // 幂等由 LogService 内部保证，即使接着走到 Main 的 finally 也不会重复。
            LogService.EndSession(_exitReason);

            Application.Exit();
            Environment.Exit(exitCode);
        }

        /// <summary>尽力删除文件，忽略所有异常</summary>
        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }

        /// <summary>
        /// 创建更新脚本（批处理文件）- 静默更新，不提示不重启
        ///
        /// 路径传递方式：不用字符串插值把路径写进脚本，而是通过环境变量传递。
        /// 因为用户可能把程序放在任意目录，路径中若含 & % ! " 等字符，
        /// 直接插入批处理会造成语法破坏甚至命令注入。
        ///
        /// 失败保护：替换前先备份旧 exe，复制失败时回滚并重试（有次数上限），
        /// 避免"旧文件已删、新文件没拷上"导致程序彻底丢失。
        /// </summary>
        private static ProcessStartInfo CreateUpdateScript(string scriptPath, string currentExe, string newExe, Version newVersion)
        {
            string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
            string backupExe = currentExe + ".bak";
            File.WriteAllText(pendingFile, newExe);

            // 全部走 ASCII：批处理默认编码为 GBK，脚本内出现中文会导致解析错乱
            string batchContent = @"@echo off
setlocal enabledelayedexpansion
title Updating

set ""TARGET=%SEEWO_TARGET%""
set ""SOURCE=%SEEWO_SOURCE%""
set ""BACKUP=%SEEWO_BACKUP%""
set ""PENDING=%SEEWO_PENDING%""

:: Give the main process time to fully exit
ping -n 3 127.0.0.1 > nul

:: Kill any lingering instance
taskkill /f /im ""%~nx1"" > nul 2>&1
ping -n 2 127.0.0.1 > nul

:: Back up current executable
if exist ""%TARGET%"" copy /y ""%TARGET%"" ""%BACKUP%"" > nul

:: Replace with retries (bounded, never loop forever)
set /a ATTEMPT=0
:copyloop
set /a ATTEMPT+=1
if !ATTEMPT! GTR 5 goto :copyfailed

del /f /q ""%TARGET%"" > nul 2>&1
copy /y ""%SOURCE%"" ""%TARGET%"" > nul 2>&1

if exist ""%TARGET%"" goto :success
ping -n 2 127.0.0.1 > nul
goto :copyloop

:copyfailed
:: Restore backup so the program is not left missing
if exist ""%BACKUP%"" (
    copy /y ""%BACKUP%"" ""%TARGET%"" > nul 2>&1
    del /f /q ""%BACKUP%"" > nul 2>&1
)
echo UPDATE_FAILED > ""%PENDING%""
del /f /q ""%SOURCE%"" > nul 2>&1
exit /b 1

:success
:: Verify the new file is a real PE binary
set /a SIZE=0
for %%F in (""%TARGET%"") do set /a SIZE=%%~zF
if !SIZE! LSS 1024 goto :copyfailed

del /f /q ""%SOURCE%"" > nul 2>&1
del /f /q ""%SOURCE%.sig"" > nul 2>&1
del /f /q ""%BACKUP%"" > nul 2>&1
del /f /q ""%PENDING%"" > nul 2>&1
del /f /q ""%~f0"" > nul 2>&1
exit /b 0
";
            File.WriteAllText(scriptPath, batchContent);
            WriteLog($"更新脚本已创建: {scriptPath}");

            // 启动脚本时通过环境变量传递路径，避免路径中的特殊字符破坏批处理语法
            var psi = new ProcessStartInfo
            {
                FileName = scriptPath,
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            psi.EnvironmentVariables["SEEWO_TARGET"] = currentExe;
            psi.EnvironmentVariables["SEEWO_SOURCE"] = newExe;
            psi.EnvironmentVariables["SEEWO_BACKUP"] = backupExe;
            psi.EnvironmentVariables["SEEWO_PENDING"] = pendingFile;

            return psi;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                if (args.Name.StartsWith("Microsoft.Win32.TaskScheduler"))
                {
                    WriteLog($"AssemblyResolve 被触发，尝试加载：{args.Name}");
                    string appDir = Path.GetDirectoryName(Application.ExecutablePath);
                    string dllPath = Path.Combine(appDir, "Microsoft.Win32.TaskScheduler.dll");

                    if (File.Exists(dllPath))
                    {
                        Assembly assembly = Assembly.LoadFrom(dllPath);
                        WriteLog($"成功从 {dllPath} 加载程序集");
                        return assembly;
                    }
                    else
                    {
                        WriteLog($"文件不存在：{dllPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"程序集解析异常：{ex.Message}");
            }
            return null;
        }

        private static void EnsureRequiredDllsExist()
        {
            string[] dllsToExtract = { "Microsoft.Win32.TaskScheduler.dll" };
            string appDir = Path.GetDirectoryName(Application.ExecutablePath);

            foreach (string dllName in dllsToExtract)
            {
                string dllPath = Path.Combine(appDir, dllName);
                if (File.Exists(dllPath))
                {
                    WriteLog($"DLL已存在：{dllName}");
                    try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); }
                    catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
                    continue;
                }

                try
                {
                    WriteLog($"正在提取DLL：{dllName}");
                    // 尝试多种资源名
                    string[] possibleResourceNames = {
                $"WindowsFormsApp1.{dllName}",
                $"WindowsFormsApp1.Resources.{dllName}",
                $"SeewoOpt.{dllName}",
                $"SeewoOpt.Resources.{dllName}"
            };
                    Stream resourceStream = null;
                    foreach (var resName in possibleResourceNames)
                    {
                        resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resName);
                        if (resourceStream != null) break;
                    }

                    if (resourceStream == null)
                    {
                        WriteLog($"错误：找不到嵌入式资源 {dllName}");
                        WriteLog("可用的资源列表：");
                        foreach (string res in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                            WriteLog($"  - {res}");
                        continue;
                    }

                    using (resourceStream)
                    using (FileStream fileStream = new FileStream(dllPath, FileMode.Create, FileAccess.Write))
                    {
                        resourceStream.CopyTo(fileStream);
                    }
                    WriteLog($"√ DLL提取成功：{dllPath}");
                    try { Assembly.LoadFrom(dllPath); WriteLog($"预加载程序集成功：{dllName}"); }
                    catch (Exception ex) { WriteLog($"预加载程序集失败：{ex.Message}"); }
                }
                catch (Exception ex)
                {
                    WriteLog($"× 提取DLL {dllName} 时出错：{ex.Message}");
                }
            }
        }

        /// <summary>
        /// 验证文件是否为有效的 Windows PE 可执行文件（检查 MZ 和 PE 头）
        /// </summary>
        private static bool IsValidPE(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length < 64) return false;

                    byte[] mzHeader = new byte[2];
                    fs.Read(mzHeader, 0, 2);
                    if (mzHeader[0] != 'M' || mzHeader[1] != 'Z') return false;

                    fs.Seek(0x3C, SeekOrigin.Begin);
                    byte[] peOffsetBytes = new byte[4];
                    fs.Read(peOffsetBytes, 0, 4);
                    int peOffset = BitConverter.ToInt32(peOffsetBytes, 0);

                    if (peOffset < 64 || peOffset > fs.Length - 4) return false;

                    fs.Seek(peOffset, SeekOrigin.Begin);
                    byte[] peHeader = new byte[4];
                    fs.Read(peHeader, 0, 4);
                    return peHeader[0] == 'P' && peHeader[1] == 'E' && peHeader[2] == 0 && peHeader[3] == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 计算文件的 SHA256 哈希值
        /// </summary>
        private static string ComputeSHA256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();
            }
        }

        // 日志统一委托给 LogService，本类保留同名包装以免改动所有调用点。
        private static void WriteLog(string message)
        {
            LogService.Write(message);
        }
    }
}