using System;
using System.Collections.Generic;
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

        // ---- 更新源 ----
        //
        // 【为什么有两个】教室网络里 GitHub 经常整段不可达：
        //   · api.github.com —— DNS 被污染，或返回一个连不上的地址；
        //   · objects.githubusercontent.com —— Release 资产实际存放的地方，
        //     常在 IP 层被拦。
        // 用户看到的"无法连接到远程服务器"就是这两处，换个域名修不了，
        // 只能换一个国内可达的托管。
        //
        // Gitee 的发行版附件实测可匿名直链下载（302 到 foruda CDN 后返回 200，
        // 不需要 Cookie、不需要验证码、不需要代理），所以作为并列源：
        // 两边都查，取版本更高的；版本相同时优先用 Gitee 下载。
        // 任一源查询失败不影响另一个，也不会让程序其它功能受影响。
        private const string GITHUB_API = "https://api.github.com/repos/Foxelf-Studio/SeewoOptimizer/releases/latest";
        private const string GITEE_API = "https://gitee.com/api/v5/repos/Foxelf-Studio/SeewoOptimizer/releases/latest";
        private const string GITHUB_USER_AGENT = "SeewoOptimizer-Updater";

        // 查版本这一步的预算（秒）。
        // WebClient 没有 Timeout 属性，底层默认约 100 秒；原来这里没包超时，
        // 网络不通时表现为"检查更新像卡死"，而不是快速失败并换源。
        private const int VERSION_QUERY_TIMEOUT_SECONDS = 15;

        /// <summary>
        /// 更新缓存目录。
        ///
        /// 【为什么引用 LogService 而不是自己拼路径】
        /// 早先这里自己写了 Path.Combine(LocalApplicationData, "TimeSyncTool", "updates")，
        /// 与 LogService 里的日志目录各拼各的——同一串目录名散在两处，
        /// 改名时必然漏一处（这次改 SeewoOpt 时就差点只改到一边）。
        /// 现在统一由 LogService 定义，二者是同一目录下的兄弟关系，
        /// 目录名只此一处。
        /// </summary>
        private static readonly string UPDATE_DIR = LogService.UpdateDirectory;

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
        /// 退出时等待后台同步线程收尾的最长时间（毫秒）。
        ///
        /// 取 3 秒：这里已先发了取消请求，正常情况几十毫秒内就会退出。
        /// 留这个上限只是为了防止线程卡在某个不可中断的调用上时拖死退出——
        /// 同步线程是 IsBackground，超时不会阻塞进程结束，只是日志缺个尾巴。
        /// </summary>
        private const int BACKGROUND_DRAIN_TIMEOUT_MS = 3000;

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
                // 【单实例检查必须在开日志分段之前】
                // 日志文件是**多进程共用**的，而分段模型默认"同一时刻只有一次运行在写"。
                // 若第二个实例先开一段、又匆匆写下自己的结束标记，而第一个实例仍在运行，
                // 那么第一个实例后续的日志就会插到第二个实例的结束标记之后——
                // 实测就是这样：第二个实例（只活了 1 秒）的分段把第一个实例
                // 剩余的十几行收尾日志整段吞了进去，读日志的人会彻底误判。
                //
                // 所以顺序改为：先抢单实例，抢不到就**不碰日志文件**，
                // 只弹一个提示框然后退出。抢到了才开分段——此时可以确保
                // 整个文件只有本进程在写。
                if (!IsSingleInstance())
                {
                    // 抢不到实例说明已有程序在跑，此刻绝不能动日志文件。
                    // 提示框在下面（需要在显示前确保日志服务不被误用）。
                    MessageBox.Show("程序已在运行中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 抢到实例后才开分段：裁剪旧运行 + 写分隔头
                LogService.BeginSession();

                // 检查是否有待应用的更新（更新后重启）- 只处理不弹窗
                HandlePendingUpdate();

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
                // 【顺序】必须先让后台工作线程收尾，再写结束标记。
                //
                // 同步线程是独立线程，主线程走完 Application.Run 返回时它可能还在跑。
                // 若此刻就写结束标记，它的最后两条日志（"同步线程正常结束"、
                // "同步线程 finally 块"）会落到标记**之下**，看上去像属于下一次运行。
                //
                // 【曾经失败的原因】早先 WaitForBackgroundWork 拿不到窗体实例——
                // 查 _activeInstance 的清理被挂在窗体的 FormClosed 上，而
                // FormClosed 早于 Application.Run 返回、也就早于这里，
                // 于是引用已是 null，Join 被静默跳过，问题照旧。
                // 现在清理改由下面的 ClearActiveInstance 显式负责。
                TimeSyncForm.WaitForBackgroundWork(BACKGROUND_DRAIN_TIMEOUT_MS);

                ReleaseMutex();

                // 收尾日志分段。放在 finally 里是因为它是唯一保证会执行的路径：
                // 正常退出、Main 抛异常、Run 抛异常都会走到这里。
                // 被强杀（任务管理器结束进程、断电）时不会执行——那种情况
                // 日志里就没有结束标记，下次启动一眼能看出上次是异常终止的。
                // 重复调用是安全的：LogService.EndSession 自身幂等。
                LogService.EndSession(_exitReason);

                // 收尾彻底结束后才松手，保证引用覆盖整个等待窗口
                TimeSyncForm.ClearActiveInstance();
            }
        }

        /// <summary>
        /// 检查是否只有一个实例在运行。
        ///
        /// 【失败时不能留下 mutex 对象】`new Mutex(true, name, out createdNew)` 在
        /// createdNew=false 时，本进程**并未获得**这个互斥体——它只是打开了一个
        /// 指向别人已持有互斥体的句柄。此时若去 ReleaseMutex()，会抛
        /// "Object synchronization method was called from an unsynchronized
        /// block of code"（实测日志里确实出现过这条）。
        /// 所以未拿到所有权时立刻把句柄关掉并置空，让 ReleaseMutex 无事可做。
        /// </summary>
        private static bool IsSingleInstance()
        {
            try
            {
                // 尝试创建互斥体
                bool createdNew;
                singleInstanceMutex = new Mutex(true, "TimeSyncTool_UniqueMutex", out createdNew);

                if (!createdNew)
                {
                    // 没拿到所有权：立即释放句柄，且不让 ReleaseMutex 再碰它
                    try { singleInstanceMutex.Close(); } catch { }
                    singleInstanceMutex = null;
                    return false;
                }
                return true;
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
                // 【这不是故障，别报成故障】
                // 更新流程跑在后台线程上，而 Mutex 只能由**获取它的那个线程**释放，
                // 跨线程调用会抛"Object synchronization method was called from an
                // unsynchronized block of code"。实机日志里那行看着像出错，
                // 其实无害：进程退出时 Windows 会自动放弃未释放的互斥体，
                // 单实例保护不会因此失效（紧随其后的就是 Environment.Exit）。
                //
                // 措辞如实说明"由谁释放"，避免下次排查又被这行吓一跳。
                WriteLog($"互斥体未在本线程释放（{ex.Message}）；"
                       + "进程退出时由系统自动释放，不影响单实例保护");
            }
        }

        /// <summary>
        /// 检查是否有待应用的更新（更新后重启）。
        ///
        /// pending.txt 有两种合法内容：
        ///   · 新 exe 的完整路径  → 更新脚本已完成替换，清掉标记即可；
        ///   · "UPDATE_FAILED"    → 脚本在复制/替换阶段失败并已回滚，
        ///                          记一条明确的失败日志后清掉标记。
        ///
        /// 【旧实现的缺陷】它只判断“内容是不是一个存在的文件”。
        /// 失败标记不是路径，File.Exists("UPDATE_FAILED") 永远为 false，
        /// 于是标记永远清不掉，每次启动都误报“检测到待应用的更新”，
        /// 真正的失败原因反而被这句噪音盖住。
        /// </summary>
        private static void HandlePendingUpdate()
        {
            try
            {
                string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
                if (!File.Exists(pendingFile)) return;

                string content = File.ReadAllText(pendingFile).Trim();

                if (string.Equals(content, "UPDATE_FAILED", StringComparison.OrdinalIgnoreCase))
                {
                    WriteLog("上次自动更新在替换阶段失败（UPDATE_FAILED），"
                           + "程序保持原版本运行；已清除失败标记");
                    TryDelete(pendingFile);
                    return;
                }

                if (File.Exists(content))
                {
                    // 注意：这里不替换文件，因为已经在更新脚本中完成了
                    // 只清理标志文件
                    File.Delete(pendingFile);
                    WriteLog("更新标志已清除");
                    return;
                }

                // 路径已失效（缓存被清理、手工删除了下载文件等）：
                // 标记没有意义，清掉并留下一条可查的记录。
                WriteLog($"待更新标记内容不是有效路径，已清除：{content}");
                TryDelete(pendingFile);
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

                // 两个源并行查：任何一边通了都能出结果，互不阻塞。
                // 各自 15 秒超时，所以最坏 15 秒就有结论——原来这里没有超时，
                // 网络不通时会白等 WebClient 默认的约 100 秒，
                // 表现成"检查更新卡死"，用户以为程序坏了。
                UpdateCandidate[] found = await Task.WhenAll(
                    FetchLatestReleaseAsync(GITEE_API, "Gitee", true),
                    FetchLatestReleaseAsync(GITHUB_API, "GitHub", false));

                UpdateCandidate best = UpdateSourcePolicy.PickNewer(found[0], found[1]);

                if (best == null)
                {
                    // 两个源都不可用：这是网络环境问题，不是程序故障。
                    // 仍然置位失败标记——时间同步成功后会再试一次，
                    // 而"时钟错误导致所有 HTTPS 失败"正是要靠那次重试救回来的场景。
                    _updateCheckFailed = true;
                    WriteLog("两个更新源都不可用，本次跳过更新检查（不影响关机、同步等其它功能）");
                    return;
                }

                WriteLog($"选定更新源: {best.SourceName}，最新版本 {best.Version}");
                WriteLog($"选定资产: {best.AssetName}");
                WriteLog($"下载地址: {best.DownloadUrl}");

                _expectedUpdateHash = best.ExpectedSha256;
                if (_expectedUpdateHash != null)
                    WriteLog($"从发布说明中提取到 SHA256: {_expectedUpdateHash}");
                else
                    WriteLog("发布说明中未找到 SHA256 哈希，将仅依赖签名校验");

                if (best.Version.CompareTo(currentVersion) > 0)
                {
                    WriteLog("发现新版本，准备自动更新");

                    // 缓存更新信息
                    PendingUpdateInfo = Tuple.Create(best.Version.ToString(), "正在自动更新...");
                    UpdateDetected?.Invoke(best.Version.ToString(), "正在自动更新...");

                    // 备用下载地址：另一个源如果报的是同一个版本，
                    // 它就是这个版本的天然备份（Gitee 不通换 GitHub，反之亦然）。
                    // 版本不同就不加——那说明另一个源还没同步，拿它的地址只会下到旧版。
                    List<string> downloadUrls = UpdateSourcePolicy.BuildDownloadUrls(best, found);

                    // 带重试的下载
                    await DownloadUpdateWithRetryAsync(
                        downloadUrls, best.AssetName, best.Version, best.ExpectedSha256);
                }
                else
                {
                    WriteLog("已是最新版本");
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
        /// 查询一个更新源的"最新发行版"。失败返回 null，不抛异常。
        ///
        /// 两个源（GitHub / Gitee）的 JSON 结构足够接近
        /// （tag_name / body / assets[].name / assets[].browser_download_url），
        /// 所以复用同一个解析器。
        ///
        /// 已实测的差异：Gitee 的资产对象**没有 size 字段**，
        /// 反序列化出来是 0，所以日志里不能用它判断文件大小。
        /// </summary>
        private static async Task<UpdateCandidate> FetchLatestReleaseAsync(
            string apiUrl, string sourceName, bool isDomestic)
        {
            try
            {
                string json;
                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("User-Agent", GITHUB_USER_AGENT);
                    // Accept 是 GitHub 的约定；Gitee 不认这个头。
                    // 只给 GitHub 加，避免以后看代码时误以为 Gitee 也需要。
                    if (!isDomestic)
                        client.Headers.Add("Accept", "application/vnd.github.v3+json");
                    client.Encoding = Encoding.UTF8;

                    json = await DownloadStringWithTimeoutAsync(
                        client, apiUrl, VERSION_QUERY_TIMEOUT_SECONDS, sourceName + " 版本查询");
                }

                // 用框架自带的 JSON 序列化器解析，正确处理转义与多资产场景
                if (!ReleaseInfoParser.TryParse(json, out ReleaseInfo release, out string parseError))
                {
                    WriteLog($"{sourceName}: 解析发行版信息失败（{parseError}）");
                    return null;
                }

                // 从资产列表中挑选可执行文件——不再简单取第一个
                ReleaseAsset asset = ReleaseInfoParser.SelectExecutableAsset(release.Assets);
                if (asset == null)
                {
                    WriteLog($"{sourceName}: 发行版 {release.TagName} 中没有 .exe 资产");
                    return null;
                }

                Version version = ReleaseInfoParser.ParseVersion(release.TagName);
                if (version == null)
                {
                    WriteLog($"{sourceName}: 无法解析版本号 {release.TagName}");
                    return null;
                }

                // 从发布说明（body）中提取预期的 SHA256 哈希
                return new UpdateCandidate
                {
                    SourceName = sourceName,
                    IsDomestic = isDomestic,
                    Version = version,
                    AssetName = asset.Name,
                    DownloadUrl = asset.BrowserDownloadUrl,
                    ExpectedSha256 = ReleaseInfoParser.TryExtractExpectedSha256(release.Body)
                };
            }
            catch (Exception ex)
            {
                // 单源失败只记一行：常见原因是 DNS 污染、IP 被拦、
                // 或者 Gitee 仓库还没创建（404）。这些都属于"这个源用不了"，
                // 不是"程序坏了"，不该让整次检查以异常收场。
                //
                // 【403 要单独说清楚】它不是网络不通，是 GitHub 的限流：
                // 未认证请求每个出口 IP 每小时 60 次。学校几十台机器共用一个
                // NAT 出口，早上开机高峰能在几分钟内撞满——这时日志里若只写
                // "远程服务器返回错误: (403)"，看起来像被墙，排查方向就错了。
                // 这也正是需要第二个源的原因之一。
                HttpStatusCode? status = GetHttpStatus(ex);

                if (status == HttpStatusCode.Forbidden && !isDomestic)
                {
                    WriteLog($"{sourceName} 版本查询失败：API 限流（403），"
                           + "同一出口 IP 每小时只有 60 次未认证额度，本次跳过");
                }
                else if (status == HttpStatusCode.NotFound && isDomestic)
                {
                    WriteLog($"{sourceName} 版本查询失败：仓库或发行版不存在（404），"
                           + "国内源尚未发布过这一版");
                }
                else
                {
                    WriteLog($"{sourceName} 版本查询失败：{ex.Message}");
                }

                return null;
            }
        }

        /// <summary>从 WebException 里取出 HTTP 状态码；不是 HTTP 错误则返回 null。</summary>
        private static HttpStatusCode? GetHttpStatus(Exception ex)
        {
            WebException webException = ex as WebException;
            if (webException == null) return null;

            HttpWebResponse response = webException.Response as HttpWebResponse;
            if (response == null) return null;

            HttpStatusCode status = response.StatusCode;
            // 取完就释放：这个响应体我们不读，留着只会占连接。
            response.Close();
            return status;
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
        /// 带重试机制的下载更新文件。
        ///
        /// 【为什么参数是一串地址】有了两个源之后，同一个版本通常有两个
        /// 可下载地址（Gitee 与 GitHub）。原来一次只试一个 URL，
        /// 主源不通就整轮失败，白白浪费另一个本来可用的地址。
        ///
        /// 【为什么要有总预算】每个地址单次预算 60 + 30 秒，若不封顶，
        /// 两轮 × 两个地址就是 6 分钟——网络不通时用户只会以为程序卡死。
        /// 这里整体 180 秒封顶，到点就如实失败，保持旧版本继续运行。
        /// </summary>
        private static async Task<bool> DownloadUpdateWithRetryAsync(
            IList<string> downloadUrls, string fileName, Version newVersion, string expectedHash = null)
        {
            const int totalBudgetMs = 180000;
            const int retryDelayMs = 2000;

            Stopwatch budget = Stopwatch.StartNew();

            for (int round = 1; round <= 2; round++)
            {
                for (int i = 0; i < downloadUrls.Count; i++)
                {
                    if (totalBudgetMs - budget.ElapsedMilliseconds < 15000)
                    {
                        WriteLog($"更新下载已用尽 {totalBudgetMs / 1000} 秒总预算，停止尝试");
                        WriteLog("更新最终失败，程序保持原版本继续运行（未做任何改动）");
                        return false;
                    }

                    try
                    {
                        await DownloadUpdateAsync(downloadUrls[i], fileName, newVersion, expectedHash);
                        return true; // 成功
                    }
                    catch (Exception ex)
                    {
                        // 【措辞修正，来自实机日志】
                        // 这个方法包住的是"下载 → 验签 → 校验版本 → 生成并启动更新脚本"
                        // 整条链路，任何一步失败都会走到这里。原来一律写"下载失败"，
                        // 而实机日志里明明已经出现"下载进度 100% / 签名校验有效 /
                        // 哈希与发布说明一致"，紧接着却是一行"下载失败"——
                        // 排查方向被完全带偏（那次真正失败的是启动脚本）。
                        WriteLog($"更新失败（第 {round} 轮，源 {i + 1}/{downloadUrls.Count}）: {ex.Message}");
                    }
                }

                if (round == 1)
                {
                    WriteLog($"等待 {retryDelayMs}ms 后换源重试...");
                    await Task.Delay(retryDelayMs);
                }
            }

            WriteLog("更新最终失败，程序保持原版本继续运行（未做任何改动）");
            return false;
        }

       
        /// <summary>
        /// 带显式超时的文件下载。
        ///
        /// 【为什么不能只给 exe 加超时】WebClient 没有 Timeout 属性，
        /// 底层 HttpWebRequest 默认约 100 秒。exe 那条链路自己包了 60 秒，
        /// 但 384 字节的 .sig 走的是裸 DownloadFileTaskAsync——
        /// 遇到断网、认证门户或代理卡住时，会白等约 100 秒，
        /// 期间 _updateCheckRunning 一直被占，后续重试直接跳过。
        ///
        /// 这里用 CTS + CancelAsync 给每次下载一个明确预算，
        /// 超时就把被取消的任务观察掉，避免 UnobservedTaskException。
        /// </summary>
        private static async Task DownloadFileWithTimeoutAsync(
            WebClient client, string url, string path, int timeoutSeconds, string what)
        {
            using (var cts = new CancellationTokenSource())
            {
                cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

                Task downloadTask = client.DownloadFileTaskAsync(new Uri(url), path);
                Task completed = await Task.WhenAny(downloadTask, Task.Delay(-1, cts.Token));

                if (completed != downloadTask)
                {
                    client.CancelAsync();
                    // 观察被取消的任务，别让它的异常留到终结器线程。
                    _ = downloadTask.ContinueWith(t => t.Exception,
                        TaskContinuationOptions.OnlyOnFaulted);

                    throw new TimeoutException(what + "下载超时（" + timeoutSeconds + " 秒）");
                }

                await downloadTask;
            }
        }

        /// <summary>
        /// 带显式超时的字符串下载（用于查询版本 JSON）。
        ///
        /// 与 DownloadFileWithTimeoutAsync 同一套办法，只是返回文本：
        /// WebClient 没有 Timeout 属性，裸调 DownloadStringTaskAsync 会等约 100 秒。
        /// 两个源并行查询时若不设上限，一个被墙的源就能把整次检查拖住。
        /// </summary>
        private static async Task<string> DownloadStringWithTimeoutAsync(
            WebClient client, string url, int timeoutSeconds, string what)
        {
            using (var cts = new CancellationTokenSource())
            {
                cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

                Task<string> queryTask = client.DownloadStringTaskAsync(new Uri(url));
                Task completed = await Task.WhenAny(queryTask, Task.Delay(-1, cts.Token));

                if (completed != queryTask)
                {
                    client.CancelAsync();
                    // 观察被取消的任务，别让它的异常留到终结器线程。
                    _ = queryTask.ContinueWith(t => t.Exception,
                        TaskContinuationOptions.OnlyOnFaulted);

                    throw new TimeoutException(what + "超时（" + timeoutSeconds + " 秒）");
                }

                return await queryTask;
            }
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
                client.DownloadProgressChanged += (s, e) =>
                {
                    try
                    {
                        if (e.ProgressPercentage % 10 == 0)
                            WriteLog($"下载进度: {e.ProgressPercentage}%");
                    }
                    catch { }
                };

                await DownloadFileWithTimeoutAsync(
                    client, downloadUrl, downloadedFile, 60, "更新文件");
            }

            WriteLog("更新文件下载完成，下载配套签名...");

            // 签名文件与 exe 同名同目录，仅后缀不同
            string signatureUrl = downloadUrl + SignatureVerifier.SIGNATURE_EXTENSION;
            try
            {
                using (WebClient client = new WebClient())
                {
                    // 给 .sig 一个比 exe 更短的预算：它只有几百字节，
                    // 100 秒的默认等待没有任何意义，只会拖住更新检查。
                    await DownloadFileWithTimeoutAsync(
                        client, signatureUrl, signatureFile, 30, "签名文件");
                }
                WriteLog("签名文件下载完成");
            }
            catch (Exception ex)
            {
                TryDelete(downloadedFile);
                TryDelete(signatureFile);
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

            // 【版本校验：引入第二个源之后，这一条从"锦上添花"变成"必须"】
            // 签名只能证明"这个文件确实是作者发的"，不能证明"是我要的那一版"。
            // 任何处在中间的源（镜像、被劫持的域名、配置错的 CDN）都可以
            // 返回一个真实、签名有效的**旧版** exe，再配一个写着高版本号的 JSON。
            // 不校验版本，就等于把"降级到有已知缺陷的旧版"做成了一条正常路径。
            if (!MatchesTargetVersion(downloadedFile, newVersion, out string versionError))
            {
                TryDelete(downloadedFile);
                TryDelete(signatureFile);
                throw new Exception(versionError);
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

            // 【只创建一次】原先这里调用了两次 CreateUpdateScript：
            // 第一次只为写脚本、记日志，第二次重新写一遍并把 psi 返回去启动。
            // 后果是每次尝试的日志里都出现两遍"更新脚本已创建"，
            // 排查时看起来像"脚本被创建了两份/重试了两次"，极具误导性
            // （2026-10-11 的实机日志正是这样）。
            ProcessStartInfo updateProcess = CreateUpdateScript(
                updaterScript, currentExe, downloadedFile, newVersion);

            // 先启动脚本，再退出本进程：脚本开头有 3 次 ping 的等待，
            // 足够主程序退出并把 exe 的文件占用让出来。
            Process.Start(updateProcess);
            WriteLog("更新脚本已启动，主程序即将退出，由脚本完成后台替换");

            // Environment.Exit 不会执行 finally，这里显式释放一次互斥体。
            //
            // 【更正一条原来写错的理由】原注释说"不释放互斥体，
            // 批处理里的 taskkill 可能杀不掉本进程"——那是错的：
            // 持有互斥体与能否被 taskkill 无关。真正兜底的是进程退出时
            // Windows 自动放弃未释放的互斥体（下面紧随的 Environment.Exit
            // 就会触发）。这里保留调用只是"提前交还句柄"的顺手动作，
            // 失败也无所谓，因此 ReleaseMutex 内部只记日志、不抛。
            ReleaseMutex();
            Thread.Sleep(200);

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

            // 【顺序】让后台同步线程先收尾，再写结束标记——否则它的日志会落到
            // 分隔之下。理由与 Main 的 finally 完全相同。
            //
            // 自动更新路径也要经过这里：那条路径会把同步线程一并取消，
            // 但取消是异步的，不等一下照样会漏日志。因为下面紧接着就是
            // Environment.Exit，Join 超时的兜底尤其重要——绝不能无限等。
            TimeSyncForm.WaitForBackgroundWork(BACKGROUND_DRAIN_TIMEOUT_MS);

            // 再释放互斥体。ReleaseMutex 自身会写一条"互斥体已释放"日志，
            // 放在结束标记之后就会跑到分隔之外；放在这里则恰好压在标记之前。
            //
            // 能安全把 ReleaseMutex 提前，是因为它内部会先
            // singleInstanceMutex.Close() 把句柄交还系统；Windows 在进程终止时
            // 也会自动放弃未释放的互斥体。即便紧随其后的 Environment.Exit
            // 把后续语句截断，单实例保护仍然成立。
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
        /// 路径传递方式：**把路径写进脚本本身**，并把 % 转义成 %%
        /// （见 UpdateScriptLauncher.EscapeForBatch）。早先用环境变量传，
        /// 结果连踩两个坑——详见方法尾部那段注释，简言之
        /// `ProcessStartInfo.EnvironmentVariables` 既要求
        /// `UseShellExecute = false`，又会被用户机器上"大小写重复的环境变量"
        /// 直接搞崩，等于把成败押在别人的环境上。
        ///
        /// 失败保护：替换前先备份旧 exe，复制失败时回滚并重试（有次数上限），
        /// 避免"旧文件已删、新文件没拷上"导致程序彻底丢失。
        /// </summary>
        private static ProcessStartInfo CreateUpdateScript(string scriptPath, string currentExe, string newExe, Version newVersion)
        {
            string pendingFile = Path.Combine(UPDATE_DIR, "pending.txt");
            string backupExe = currentExe + ".bak";
            File.WriteAllText(pendingFile, newExe);

            // 生成脚本内容。模板本身住在 UpdateScriptLauncher 里——
            // 放在那里才能被探针链接到**同一份源码**去真的跑一遍
            // （把模板留在本类里时，探针只能自己另写一份批处理来测，
            //   结果我漏掉四个闭合引号它也没发现）。
            string batchContent = UpdateScriptLauncher.BuildScriptContent(
                currentExe, newExe, backupExe, pendingFile);

            // 【为什么不再用环境变量传路径——两个独立的坑，都踩过】
            //   1. `EnvironmentVariables` 只在 `UseShellExecute = false` 时可用，
            //      设 true 会直接抛异常。2026-10-11 实机的自动更新失败正是
            //      死在这一条：下载、验签、哈希比对全部通过，卡在启动脚本上。
            //   2. 更要命的是，`EnvironmentVariables` 的 getter 会把当前进程的
            //      环境变量整块复制进一个大小写不敏感的字典；只要环境里存在
            //      "只有大小写不同"的两个键（如 HTTP_PROXY 与 http_proxy，
            //      MSYS/Git 环境里很常见），它就会抛
            //      "已添加项。字典中的关键字…" —— 与 UseShellExecute 完全无关。
            //      也就是说：**只要用了这个机制，就永远受制于用户机器的环境变量**。
            //
            // 所以改为把值直接写进脚本本身，一个机制解决两个问题。
            // 代价是必须自己处理批处理转义（见 UpdateScriptLauncher.EscapeForBatch）。
            File.WriteAllText(scriptPath, batchContent);
            WriteLog($"更新脚本已创建: {scriptPath}");

            // 构造启动参数。为什么必须是 "UseShellExecute = false + 经 cmd.exe 启动"、
            // 为什么 cmd.exe 要用绝对路径，全部写在 UpdateScriptLauncher 的注释里。
            //
            // 抽成服务类还有个好处：探针（tools/UpdateLaunchProbe）能链接同一份源码
            // 真的启动一次进程，而不是在测试里复刻一份构造代码。
            return UpdateScriptLauncher.BuildStartInfo(scriptPath);
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
        /// 校验下载到的 exe 的版本号是否等于目标版本。
        ///
        /// 用 FileVersion 而不是 AssemblyVersion：前者是 FileVersionInfo 能直接读到的，
        /// 也是发布流程里和 tag 对齐的那个（CI 校验的就是它）。
        ///
        /// 比较前统一成 Major.Minor.Build 三段：AssemblyVersion 可能是
        /// "26.10.17"，而某些构建会写成 "26.10.17.0"。Version 的 CompareTo
        /// 会认为缺省段是 -1，两者不相等——直接比会误判成"版本不符"，
        /// 把一次正常的更新拦下来。
        /// </summary>
        private static bool MatchesTargetVersion(string exePath, Version target, out string error)
        {
            error = null;

            try
            {
                string fileVersion = FileVersionInfo.GetVersionInfo(exePath).FileVersion;

                if (string.IsNullOrWhiteSpace(fileVersion))
                {
                    error = "下载的文件没有版本信息，更新已取消（可能是被替换过的文件）";
                    return false;
                }

                if (!Version.TryParse(fileVersion.Trim(), out Version parsed))
                {
                    error = $"无法解析下载文件的版本号“{fileVersion}”，更新已取消";
                    return false;
                }

                if (!UpdateSourcePolicy.VersionMatches(parsed, target))
                {
                    error = $"下载的文件版本为 {fileVersion}，与目标版本 {target} 不一致，"
                          + "更新已取消（可能是更新源返回了旧版本）";
                    return false;
                }

                WriteLog($"版本校验通过：{fileVersion}");
                return true;
            }
            catch (Exception ex)
            {
                error = "读取下载文件版本失败：" + ex.Message;
                return false;
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