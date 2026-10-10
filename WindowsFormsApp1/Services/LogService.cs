using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 统一日志服务。
    ///
    /// 历史问题：Program.cs 与 TimeSyncForm.cs 各自实现了一份 WriteLog，
    /// 时间格式还不一致（一个带毫秒一个不带），导致日志交错且难以排查。
    /// 现统一到此处，格式固定为 "yyyy-MM-dd HH:mm:ss.fff"。
    ///
    /// 线程安全：多个线程（UI 线程、同步线程、更新线程）会并发写同一个文件，
    /// File.AppendAllText 在并发下可能抛 IOException（文件被占用），
    /// 因此用锁串行化写入。
    ///
    /// 【运行分段，2026-10-04 新增】
    /// 原先所有运行都追加到同一个 startup.log，仅靠时间戳区分，
    /// 翻日志时看不出"上一次运行到哪里结束、这一次从哪里开始"——
    /// 尤其是异常终止的运行没有结尾标记，边界完全靠猜。
    ///
    /// 现每次启动写一条醒目分隔头，结束时写结尾标记，并只保留最近
    /// <see cref="KeepSessions"/> 次运行，避免文件无限增长。
    /// </summary>
    public static class LogService
    {
        /// <summary>
        /// 本程序在 %LOCALAPPDATA% 下的数据根目录名。
        ///
        /// 【为什么要有这个常量，而不是各处自己拼字符串】
        /// 这个目录下同时放着日志（startup.log）与更新缓存（updates\），
        /// 分属 LogService 与 Program 两个类使用，卸载流程还要整体删除。
        /// 早先三处各自硬编码了同一串字面量，改名时漏掉任何一处，
        /// 就会变成"日志写在新目录、卸载删的是旧目录"这类错位。
        /// 统一到这一处后，改名只需改这一行。
        ///
        /// 【2026-10-11 从 "TimeSyncTool" 改为 "SeewoOpt"】
        /// 旧名是项目早期的代号，界面上早已统一为"陈叔叔希沃优化助手"，
        /// 唯独这个用户能在资源管理器里看到的目录名还留着旧名，观感不一致。
        /// 改名会留一份旧数据，见 LogDirectory 的注释。
        /// </summary>
        public const string DataRootName = "SeewoOpt";

        /// <summary>
        /// 日志根目录：%LOCALAPPDATA%\SeewoOpt
        ///
        /// 【改名后旧数据怎么办——这是有意的取舍】
        /// 目录名从 TimeSyncTool 改为 SeewoOpt 后，老用户机器上
        /// %LOCALAPPDATA%\TimeSyncTool\ 里的旧日志与旧更新缓存不会被自动带走。
        /// 不做自动迁移，理由：
        ///   · 这里只有**日志与可再生的更新缓存**，没有任何用户配置
        ///     （设置存在注册表里，不在这个目录），丢掉不损失用户数据；
        ///   · 自动迁移要处理"新旧目录同时存在""迁移到一半失败"等分支，
        ///     为一个纯日志目录引入这些复杂度不划算；
        ///   · 旧的更新缓存若被迁移过来，反而可能让程序误判
        ///     "有一个待安装的旧版本"。
        /// 因此让新目录干净地重新开始，旧目录由用户自行删除即可。
        /// </summary>
        public static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            DataRootName);

        /// <summary>更新缓存目录：%LOCALAPPDATA%\SeewoOpt\updates</summary>
        public static readonly string UpdateDirectory = Path.Combine(LogDirectory, "updates");

        /// <summary>日志文件完整路径</summary>
        public static readonly string LogFilePath = Path.Combine(LogDirectory, "startup.log");

        private static readonly object _writeLock = new object();

        /// <summary>
        /// 写日志用的编码：**UTF-8 且不写 BOM**。
        ///
        /// 【为什么不能用 Encoding.UTF8】
        /// Encoding.UTF8 是带 BOM，写新文件（或整篇重写）时会在最前面
        /// 插三个字节 EF BB BF。裁剪整篇重写之后，BOM 就落到了第一个
        /// 分隔头之前，导致：
        ///   1. 文件首行变成 "\uFEFF======== RUN 开始 ========"，
        ///      StartsWith 判断、肉眼对首行都会被这个不可见字符骗到；
        ///   2. 手工用编辑器打开会显示成 ""；
        ///   3. 任何按"首行就是开始标记"做的解析都会失配。
        /// 追加写入时 BOM 不会重复出现，所以这个坑只在"裁剪重写"那次显形——
        /// 单测构造的是理想字符串，抓不到，是端到端实跑才暴露的。
        /// </summary>
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        /// <summary>
        /// 运行序号的下限基准。
        ///
        /// 【为什么需要一个基准值】序号 = 基准 + 当前文件里的段数。
        /// 之所以不能直接用段数当序号，是因为裁剪到上限后段数恒为
        /// KeepSessions，序号会卡在同一个值反复出现（实测 5 段全是
        /// "第 6 次启动"）。有了文件里记录的基准，即使旧段被裁掉，
        /// 序号仍能持续递增。
        ///
        /// 【为什么是负数】为了让全新安装的第一段显示"第 1 次"：
        /// 首次运行时文件为空、段数为 0，基准 = 1 - 0 - 1 = 0。
        /// 取一个明显不合理的负数，还能防止"基准来自文件但被读成 0"
        /// 这类错误悄悄地把序号重置。
        /// </summary>
        private const int NoBaseSentinel = -1000000;

        /// <summary>
        /// 日志写入开关。卸载流程会删除整个日志目录，必须先置为 false 停写，
        /// 否则后续写日志会因目录不存在而失败。
        /// </summary>
        public static volatile bool Enabled = true;

        /// <summary>单条日志最大长度，防止异常信息过长撑爆日志文件</summary>
        private const int MaxMessageLength = 4000;

        /// <summary>
        /// 保留最近多少次运行的日志，**含当前这次**。
        ///
        /// 取值偏小是因为这个程序开机自启、一天可能启动多次，文件会快速变大。
        /// </summary>
        public const int KeepSessions = 5;

        /// <summary>
        /// 分隔头里的固定标记。用不可见字符做锚点不合适（日志要能直接肉眼读），
        /// 因此选一条固定文案，且足够醒目、不会被普通日志误匹配。
        ///
        /// 【必须每次运行只出现一次】统计运行次数就是数这个标记的出现次数。
        /// 曾把同一串标记同时用作首尾边框，结果每次运行出现两次、
        /// 分段计数翻倍，裁剪边界随之全错。收尾的边框另用 <see cref="Divider"/>。
        /// </summary>
        private const string SessionBeginMarker = "======== RUN 开始 ========";

        /// <summary>结尾标记，用来判断上一次运行是否正常结束</summary>
        private const string SessionEndMarker = "======== RUN 结束 ========";

        /// <summary>
        /// 纯视觉分隔线，不参与任何计数。
        /// 与两个标记都不同文，避免被当成标记匹配到。
        ///
        /// 【为什么是 internal 而不是 private】
        /// 测试 fixture 要构造"一段日志"就必须用这一串，
        /// 若测试里把字面量抄一遍，产品改了分隔线测试不会跟着变，
        /// 这种漂移测试全绿却完全没有鉴别力。暴露常量是为了让
        /// "分隔线只有一处定义"这条约束能被编译器强制住。
        /// </summary>
        internal const string Divider = "----------------------------------------";

        /// <summary>本次运行的启动时刻，供分隔头与文件名使用</summary>
        private static DateTime _sessionStart;

        /// <summary>
        /// 开始一次新的运行：裁剪旧运行、写入分隔头。
        ///
        /// 必须在写任何业务日志之前调用一次，否则本次运行的日志会
        /// 被归到上一次的分段里。
        /// </summary>
        public static void BeginSession()
        {
            if (!Enabled) return;

            _sessionStart = DateTime.Now;
            _sessionEnded = false;      // 新一次运行开始，重新允许写结束标记
            _sessionStarted = true;     // 标记本进程确实开过分段

            try
            {
                lock (_writeLock)
                {
                    EnsureDirectory();
                    // 先算出本次是第几次运行——必须在裁剪**之前**读，
                    // 因为裁剪会把旧段删掉，段数随之变小。
                    // 本次序号同时就是写给下次的基准（下次读到它再 +1）。
                    string existing = ReadAllText();
                    int sessionNumber = NextSessionNumber(existing);

                    // 再裁剪，给即将写入的新段腾出位置。
                    TrimOldSessions();

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine();
                    sb.AppendLine(SessionBeginMarker);
                    sb.AppendLine(string.Format("启动时间: {0:yyyy-MM-dd HH:mm:ss.fff}", _sessionStart));
                    sb.AppendLine(string.Format("运行序号: 历史上第 {0} 次启动（本文件保留最近 {1} 次）",
                        sessionNumber, KeepSessions));
                    // 基准 = 本次序号本身。下次启动读到它，序号即为本次 +1。
                    // 这个值不含"当前有几段"的信息，所以裁剪不会让它失真。
                    sb.AppendLine(string.Format("{0}{1}", BaseLinePrefix, sessionNumber));
                    sb.Append(Divider);
                    sb.Append(Environment.NewLine);

                    File.AppendAllText(LogFilePath, sb.ToString(), Utf8NoBom);
                }
            }
            catch
            {
                // 日志失败必须静默
            }
        }

        /// <summary>
        /// 本次运行的结束标记是否已写过，保证每次运行只写一条。
        ///
        /// 【为什么必须幂等】退出可能经由两条路径先后到达：
        /// Shutdown() 会写一次；它调用 Application.Exit() 的过程中，
        /// 主窗体会关闭并让 Main 走到 finally，若 finally 里再写一次，
        /// 日志里就会出现两条结束标记——实测确实如此，同一段里出现了
        /// 两个不同时刻的"结束时间"。
        /// 幂等判断必须放在这里而不是调用方：只有 LogService 自己
        /// 清楚它有没有写过。
        /// </summary>
        private static bool _sessionEnded;

        /// <summary>
        /// 本进程是否调用过 <see cref="BeginSession"/>。
        ///
        /// 【为什么需要它】"已有实例在运行"那条退出路径会走到 Main 的 finally，
        /// 而 finally 里会调 EndSession。那个进程从未开过分段，若允许它写结束标记，
        /// 就会把一条收尾插进**正在运行的另一个实例**的分段中间——
        /// 实测就是这个效果：第二实例只活了 1 秒，却用一个结束标记
        /// 把第一实例剩余的十几行收尾日志整段截断在了外面。
        /// </summary>
        private static bool _sessionStarted;

        /// <summary>
        /// 结束本次运行，写入结尾标记。重复调用只生效一次。
        ///
        /// 有这条标记时，下次启动就能一眼看出上一段是正常退出还是中途死掉
        /// （崩溃/被强杀时不会有结束标记）——这正是原先最难判断的信息。
        /// </summary>
        public static void EndSession(string reason)
        {
            if (!Enabled) return;

            try
            {
                lock (_writeLock)
                {
                    // 本进程压根没开过分段（例如"已有实例在运行"那条路径），
                    // 就不该写结束标记——否则会在别人正在写的分段里
                    // 插一条不属于它的收尾，把整段日志的归属搞乱。
                    if (!_sessionStarted) return;
                    if (_sessionEnded) return;      // 幂等：只写一次
                    _sessionEnded = true;

                    EnsureDirectory();

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(string.Format("结束时间: {0:yyyy-MM-dd HH:mm:ss.fff}",
                        DateTime.Now));
                    if (_sessionStart != default(DateTime))
                    {
                        sb.AppendLine(string.Format("本次运行时长: {0}",
                            FormatDuration(DateTime.Now - _sessionStart)));
                    }
                    if (!string.IsNullOrEmpty(reason))
                        sb.AppendLine(string.Format("结束原因: {0}", reason));
                    sb.Append(SessionEndMarker);
                    sb.Append(Environment.NewLine);

                    File.AppendAllText(LogFilePath, sb.ToString(), Utf8NoBom);
                }
            }
            catch
            {
                // 日志失败必须静默
            }
        }

        /// <summary>
        /// 写入一条日志。永不抛异常——日志失败不应影响主流程。
        /// </summary>
        public static void Write(string message)
        {
            if (!Enabled) return;

            try
            {
                string line = string.Format(
                    "{0:yyyy-MM-dd HH:mm:ss.fff} - {1}{2}",
                    DateTime.Now,
                    Truncate(message),
                    Environment.NewLine);

                lock (_writeLock)
                {
                    if (!Directory.Exists(LogDirectory))
                        Directory.CreateDirectory(LogDirectory);

                    File.AppendAllText(LogFilePath, line, Utf8NoBom);
                }
            }
            catch
            {
                // 日志失败必须静默——不能因为写不了日志就把程序搞崩
            }
        }

        /// <summary>截断超长消息，避免单条异常堆栈占满磁盘</summary>
        private static string Truncate(string message)
        {
            if (string.IsNullOrEmpty(message)) return string.Empty;
            if (message.Length <= MaxMessageLength) return message;
            return message.Substring(0, MaxMessageLength) + " ...(已截断)";
        }

        // ---------------------------------------------------------------
        // 运行分段支持
        // ---------------------------------------------------------------

        private static void EnsureDirectory()
        {
            if (!Directory.Exists(LogDirectory))
                Directory.CreateDirectory(LogDirectory);
        }

        private static string ReadAllText()
        {
            if (!File.Exists(LogFilePath)) return string.Empty;
            try
            {
                // 这里刻意用带 BOM 的 Encoding.UTF8 来**读**：
                // 读取时它会把开头的 BOM 当作文档签名吃掉，保证内容里
                // 不含 \uFEFF；写回的编码则统一用无 BOM 的 Utf8NoBom。
                // 一读一写的不对称是有意为之，不是漏改——旧版本曾写出
                // 带 BOM 的文件，这条读取路径要能兼容它们。
                return File.ReadAllText(LogFilePath, Encoding.UTF8);
            }
            catch
            {
                // 文件被占用等情况：当作空处理，不阻断启动
                return string.Empty;
            }
        }

        /// <summary>
        /// 统计内容里出现了多少次分隔头，即已有多少次运行记录。
        /// </summary>
        internal static int CountSessions(string content)
        {
            if (string.IsNullOrEmpty(content)) return 0;

            int count = 0;
            int index = 0;
            while ((index = content.IndexOf(SessionBeginMarker, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += SessionBeginMarker.Length;
            }
            return count;
        }

        /// <summary>基准行的固定前缀，用于从日志里把基准读回来</summary>
        private const string BaseLinePrefix = "序号基准: ";

        /// <summary>
        /// 从日志内容里读回"序号基准"。
        ///
        /// 采用"最后一条基准行生效"的语义：文件里可能同时存在多条
        /// （每段都写一条），最后那条对应最新状态。找不到则返回
        /// <see cref="NoBaseSentinel"/>，调用方据此走首次运行的初始化分支。
        ///
        /// 抽成 internal 纯函数是为了可测：序号一旦算错会出现
        /// "每次都显示同一序号"这种持续性错误，肉眼极难发现。
        /// </summary>
        internal static int ReadBase(string content)
        {
            if (string.IsNullOrEmpty(content)) return NoBaseSentinel;

            int last = -1;
            int index = 0;
            while ((index = content.IndexOf(BaseLinePrefix, index, StringComparison.Ordinal)) >= 0)
            {
                last = index;
                index += BaseLinePrefix.Length;
            }
            if (last < 0) return NoBaseSentinel;

            int start = last + BaseLinePrefix.Length;
            int end = start;
            while (end < content.Length && content[end] != '\r' && content[end] != '\n')
                end++;

            int value;
            string raw = content.Substring(start, end - start).Trim();
            if (int.TryParse(raw, System.Globalization.NumberStyles.Integer,
                              System.Globalization.CultureInfo.InvariantCulture, out value))
                return value;

            return NoBaseSentinel;
        }

        /// <summary>
        /// 按"历史累计"语义算出本次运行的序号：给定文件内容，返回本次应显示的序号。
        ///
        /// 语义：序号 = 上一次运行写下的序号 + 1；文件里没有基准时（首次运行，
        /// 或升级前的旧日志）退化为"当前段数 + 1"，即之前已有几段就从几接着数。
        ///
        /// 【为什么不用 CountSessions + 1】裁剪到上限后段数恒为 KeepSessions，
        /// 那样序号会永远卡在同一个数上（实测 5 段全是"第 6 次"）。
        ///
        /// 【为什么基准直接存"上一次的序号"】曾设计成"基准 = 序号 - 段数"，
        /// 由两者共同还原。但段数会因裁剪而减少，于是这个基准必须在
        /// 裁剪前算、却要在裁剪后仍成立，两边一旦不一致序号就成对跳号。
        /// 直接存上次序号则不依赖段数，裁剪再狠也不影响。
        /// </summary>
        internal static int NextSessionNumber(string content)
        {
            int b = ReadBase(content);
            if (b == NoBaseSentinel)
                return CountSessions(content) + 1;      // 无基准：退化为段数 + 1
            return b + 1;                                // 有基准：接着上次往下数
        }

        /// <summary>
        /// 裁剪逻辑的纯函数形态：给定文件内容，返回应当保留的部分。
        ///
        /// 抽成 internal 是为了可测——裁剪边界出过两次 off-by-one，
        /// 都是"看起来对、跑起来多一段"的类型，必须能被自动验证。
        /// </summary>
        internal static string TrimContent(string content, int keepSessions)
        {
            if (string.IsNullOrEmpty(content)) return content;

            int total = CountSessions(content);
            int keepBeforeAppend = keepSessions - 1;      // 给即将写入的新段留位
            if (total <= keepBeforeAppend) return content;

            int toDrop = total - keepBeforeAppend;

            // 从第 1 个标记开始，跳过 toDrop 个，落在第 (toDrop+1) 个标记上
            // ——那正是要保留的第一段的开头。
            int cutAt = content.IndexOf(SessionBeginMarker, StringComparison.Ordinal);
            if (cutAt < 0) return content;

            for (int i = 0; i < toDrop; i++)
            {
                int afterThis = cutAt + SessionBeginMarker.Length;
                int next = content.IndexOf(SessionBeginMarker, afterThis, StringComparison.Ordinal);
                if (next < 0) break;     // 理论上不会发生（total 已数过）
                cutAt = next;
            }

            // 直接从 cutAt 起切，再把开头残留的空白全部去掉——前一段以
            // "\r\n" 结尾，若不清理，文件会以空行开头。
            int start = cutAt;
            while (start < content.Length &&
                   (content[start] == '\r' || content[start] == '\n' ||
                    content[start] == ' ' || content[start] == '\t'))
                start++;

            return content.Substring(start);
        }

        /// <summary>
        /// 只保留最近若干次运行的内容。
        ///
        /// 【保留数要减 1】<see cref="KeepSessions"/> 是"含本次运行在内"的上限。
        /// 本方法在写入新分隔头**之前**调用，所以这里只能保留
        /// <c>KeepSessions - 1</c> 段，把位置让给即将追加的新段。
        /// 曾按 KeepSessions 裁剪，结果每次固定多出一段：
        /// 实测跑 7 次后留下 6 段。单元测试构造的是理想字节序列，
        /// 没能暴露这个"裁剪 + 追加"的联合效应，是端到端实跑才发现的。
        /// </summary>
        private static void TrimOldSessions()
        {
            string content = ReadAllText();
            if (string.IsNullOrEmpty(content)) return;

            string kept = TrimContent(content, KeepSessions);
            if (kept == content) return;        // 无需裁剪

            try
            {
                File.WriteAllText(LogFilePath, kept, Utf8NoBom);
            }
            catch
            {
                // 裁剪失败不影响本次运行，日志继续追加即可
            }
        }

        private static string FormatDuration(TimeSpan span)
        {
            if (span.TotalHours >= 1)
                return string.Format("{0:0} 小时 {1:0} 分 {2:0} 秒",
                    Math.Floor(span.TotalHours), span.Minutes, span.Seconds);
            if (span.TotalMinutes >= 1)
                return string.Format("{0:0} 分 {1:0} 秒", Math.Floor(span.TotalMinutes), span.Seconds);
            return string.Format("{0:0.0} 秒", span.TotalSeconds);
        }
    }
}
