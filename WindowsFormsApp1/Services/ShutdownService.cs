using System;
using System.Collections.Generic;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 关机调度服务：常驻托盘期间盯着时间点，到点弹提醒、再执行关机。
    ///
    /// 【为什么程序要常驻而不是交给 Windows 定时任务】
    /// 这是用户的明确选择。技术上系统定时任务更省资源，但本工具的核心场景
    /// 恰恰是"本机时钟错误"：开机时钟停在 2000 年，而 Windows 定时任务
    /// 是按系统时间触发的——一个靠不住的时钟排不出准点任务，还会把
    /// 触发时刻算到 2000 年去，任务永不触发并残留下来。
    /// 由程序自己盯着，时钟一被校正过来就立刻生效，不依赖任何外部调度。
    ///
    /// 【这个类只做"判断"和"执行"，不碰 UI】
    /// 弹窗交给窗体（它才知道该用什么 owner、什么措辞），
    /// 这里只回答两个问题：此刻该提醒吗？此刻该关机吗？
    /// 好处是判断逻辑能被单元测试直接覆盖（见 ShutdownScheduleLogic）。
    /// </summary>
    public static class ShutdownService
    {
        /// <summary>
        /// 关机前的系统级倒计时秒数。
        ///
        /// 【为什么不直接 shutdown /t 0 立即关】
        /// 留这段时间有两个用处：
        ///   1. 用户误点"确认"后还能用 `shutdown /a` 撤销；
        ///   2. Windows 会走正常的"请保存你的工作"流程，
        ///      不会硬切掉未保存的文档——教室一体机上很可能正开着课件。
        /// 与提醒弹窗的提前量（5 分钟）是两回事，不要混淆。
        /// </summary>
        public const int SystemCountdownSeconds = 300;

        /// <summary>
        /// 巡检间隔（毫秒）。
        ///
        /// 取 20 秒：关机时间点按"分钟"精度生效，20 秒的轮询足以保证
        /// 命中那一分钟内被检测到，同时单次检查只是读一次系统时钟
        /// 加几次整数比较，开销在微秒级，可以忽略。
        /// 更密没有意义（多消耗 CPU 却不提高精度），更疏则可能整分钟地错过。
        /// </summary>
        public const int PollIntervalMs = 20000;

        /// <summary>
        /// 已经为"哪一次关机"弹过提醒了。防止同一轮关机重复弹窗。
        ///
        /// 【为什么记的是"关机时刻"而不是"提醒的那一分钟"——这是个真实缺陷】
        /// 早先这里记的是"已提醒过的那个绝对分钟"（_lastWarnedMinute），
        /// 判据是 `_lastWarnedMinute != 当前分钟`。它只能挡住**同一分钟内**
        /// 20 秒 / 40 秒 / 60 秒的重复弹窗，挡不住跨分钟。
        ///
        /// 而 FindRuleToWarn 的窗口是一整段 [触发时刻-提前量, 触发时刻)：
        /// 23:40 关机 → 提醒窗口是 23:35~23:39，整整五分钟。
        /// 用户在 23:35 看到框后**没点任何按钮**（框就这么挂着），
        /// 23:36 定时器再 tick：当前分钟 23:36 ≠ 已记录 23:35 → 条件成立 → 又弹一个。
        /// 于是 23:36、23:37、23:38、23:39 各弹一个，实机表现为"每分钟弹一次窗"。
        ///
        /// 正确的判据不是"这一分钟提醒过没有"，而是"**这一次关机**提醒过没有"。
        /// 一轮关机（同一个 DueAt）只应弹一次，无论横跨多少个自然分钟。
        /// 重启后内存态归零，自然会重新弹——那正是需要的行为。
        /// </summary>
        private static DateTime _warnedForDue = DateTime.MinValue;

        /// <summary>已执行过关机的时刻，防止重复下发 shutdown 命令</summary>
        private static DateTime _lastShutdownMinute = DateTime.MinValue;

        /// <summary>上一次提醒用户选择"本次不关机"所对应的关机时刻</summary>
        private static DateTime _skipUntil = DateTime.MinValue;

        /// <summary>
        /// 本轮已经"承诺"给用户的那次关机时刻——提醒框一弹出来就定下，
        /// 到点据此执行。DateTime.MinValue 表示当前没有承诺。
        ///
        /// 【为什么需要它，而不是每轮从规则重新推导】
        /// 正常路径下，"承诺时刻"与规则时刻一致（规则 23:00，22:55 弹提醒时
        /// 承诺的也是 23:00），这时确实不必额外记。但重启后补弹提醒会顺延：
        /// 22:56 重启、在 22:56~22:59 之间补弹，承诺的是"现在 + 5 分钟"
        /// （如 23:01），而规则时刻仍是 23:00。两条路径必须只关一次，
        /// 就得有个明确的"本轮到底承诺了哪一刻"作为唯一判据——
        /// 否则 23:00（规则）与 23:01（顺延）会各关一次。
        ///
        /// 由调用方在弹提醒时通过 <see cref="PromiseShutdown"/> 写入。
        /// </summary>
        private static DateTime _promisedAt = DateTime.MinValue;

        /// <summary>
        /// 结果：本轮巡检该做什么。
        /// </summary>
        public class PollResult
        {
            /// <summary>应当关机（调用方需先向用户确认）</summary>
            public bool ShouldShutdown { get; set; }

            /// <summary>应当弹出"5 分钟后关机"提醒</summary>
            public bool ShouldWarn { get; set; }

            /// <summary>命中的规则，用于在提示里写出具体时间</summary>
            public ShutdownTimeRule Rule { get; set; }

            /// <summary>本次关机的目标时刻（用于"本次不关机"的判定）</summary>
            public DateTime DueAt { get; set; }
        }

        /// <summary>
        /// 执行一次巡检。读取全部内存记账状态，判断本轮该做什么。
        ///
        /// 【关于副作用——这一条与早先的设计不同，务必看清】
        /// 早先这里只读不写，因为"弹框要等用户点击，可能等很久"，
        /// 怕把状态消耗掉导致提醒丢失。那个担心在**用户会点**的前提下成立。
        ///
        /// 但实机暴露了反面：教室一体机上用户根本不点，模态框就一直挂着，
        /// UI 线程被阻塞——这时"等弹框返回再记账"根本不会发生。
        /// 结果是提醒窗口的每一分钟都重新弹一个框（实机表现为 23:40/41/42 三个框）。
        ///
        /// 因此现在在**产出 ShouldWarn 的同时**就把两笔账记好：
        ///   · _warnedForDue —— 这一轮关机已提醒，窗口内不再重复
        ///   · _promisedAt   —— 承诺时刻已定，到点据此执行
        /// 记在产出点而不是调用点，是"用户不点也不会重弹"的唯一保证。
        /// </summary>
        public static PollResult Poll(IEnumerable<ShutdownTimeRule> rules, DateTime now)
        {
            var result = new PollResult();

            DateTime thisMinute = new DateTime(now.Year, now.Month, now.Day,
                                               now.Hour, now.Minute, 0, now.Kind);

            // ---- 第一优先：已承诺的那次关机到点了 ----
            //
            // 这是本服务的主路径：提醒框弹出时就把"承诺时刻"定下来了
            // （由窗体调 PromiseShutdown 写入），此后每轮巡检只做一件事——
            // 看当前分钟是否等于承诺时刻。相等就关机。
            //
            // 【为什么不再从规则实时推导】
            // 因为重启后补弹提醒会顺延（22:56 补弹 → 承诺 23:01），
            // 而规则时刻还是 23:00。若仍按规则推导，23:00 会先关一次，
            // 23:01 又因承诺再关一次。用"承诺时刻"作唯一判据，
            // 23:00 那一刻 _promisedAt 指向的是 23:01，自然不命中，只关一次。
            if (_promisedAt != DateTime.MinValue
                && _promisedAt == thisMinute
                && _lastShutdownMinute != thisMinute
                && _skipUntil != thisMinute)
            {
                result.ShouldShutdown = true;
                result.Rule = FindRuleForMoment(rules, thisMinute);
                result.DueAt = thisMinute;
                return result;
            }

            // ---- 第二优先：该弹"5 分钟后关机"提醒了吗 ----
            //
            // 两种情况会走到这里：
            //   a) 正常路径：规则 23:00，22:55 命中提醒窗口 → 承诺 23:00
            //   b) 重启补弹：22:56 重启后 _warnedForDue 归零，
            //      此刻仍在 22:55~22:59 的提醒窗口内 → 重新弹，
            //      并把承诺时刻顺延为"现在 + 5 分钟"
            //
            // 注意 b) 的判据：只要"某条规则的下一次触发时刻"距现在 ≤ 提前量，
            // 就算还在窗口内。这样 22:56 重启也能补弹，而不局限于 22:55 整。
            //
            // 【去重判据：以"这一次关机"为单位，不是以"这一分钟"为单位】
            // 见 _warnedForDue 的注释——按分钟去重会让用户在提醒窗口内
            // （最长 5 分钟）看到 5 个弹窗。
            ShutdownTimeRule warnRule = ShutdownScheduleLogic.FindRuleToWarn(rules, now);
            if (warnRule != null)
            {
                DateTime? next = ShutdownScheduleLogic.NextOccurrence(
                    warnRule, thisMinute.AddMinutes(-1));

                // 【两道条件各司其职，不要合并】
                //   _warnedForDue != next.Value  —— 这一轮关机已提醒过（窗口内去重的**主防线**）
                //   _skipUntil    != next.Value  —— 用户已选"本次不关机"，别再问
                //
                // 曾经这里还有第三道 `_promisedAt != next.Value`，已删除：
                // 它的本意是"已承诺过就别再弹"，但承诺时刻在**顺延路径**上
                // （22:56 补弹 → 承诺 23:01）并不等于 next（23:00），
                // 于是它在最需要它挡的重启场景里恰好失效——真正挡住重复弹窗的
                // 始终是 _warnedForDue。留着它只会让人误以为它在起作用，
                // 而删掉它测试纹丝不动，正是"冗余防线掩盖真防线"的典型。
                if (next.HasValue && _warnedForDue != next.Value
                    && _skipUntil != next.Value)
                {
                    // 承诺时刻 = 从现在起算满提前量。
                    // 正常情况（22:55 弹）它恰好等于 next（23:00）；
                    // 重启补弹（22:56 弹）则顺延到 23:01。
                    DateTime promised = thisMinute
                        .AddMinutes(ShutdownScheduleLogic.WarnMinutesAhead);
                    if (promised < next.Value) promised = next.Value;

                    result.ShouldWarn = true;
                    result.Rule = warnRule;
                    result.DueAt = promised;

                    // 把"这一次关机"标成已提醒，让本轮窗口内后续分钟不再重弹。
                    // 这里直接记账（而不是等调用方 MarkWarned）是必要的：
                    // 调用方弹的是模态框，用户不点时它会一直挂着并阻塞 UI 线程，
                    // 而"不点"恰恰是最常见的场景——必须在这里就把去重定死。
                    _warnedForDue = next.Value;
                    PromiseShutdown(promised);
                    return result;
                }
            }

            return result;
        }

        /// <summary>
        /// 从规则里找出"目标时刻对应哪条规则"，仅用于把规则写进提示文案。
        /// 找不到时返回 null（例如重启补弹后承诺时刻已顺延、不再等于规则时刻），
        /// 调用方需容忍 null。
        /// </summary>
        private static ShutdownTimeRule FindRuleForMoment(
            IEnumerable<ShutdownTimeRule> rules, DateTime moment)
        {
            if (rules == null) return null;
            foreach (ShutdownTimeRule rule in rules)
            {
                if (rule != null && rule.Matches(moment)) return rule;
            }
            return null;
        }

        /// <summary>
        /// 记账：把"承诺时刻"定下来。窗体的提醒框一弹出就调用，
        /// 此后 Poll 只认这个时刻。
        /// </summary>
        public static void PromiseShutdown(DateTime at)
        {
            _promisedAt = new DateTime(at.Year, at.Month, at.Day,
                                       at.Hour, at.Minute, 0, at.Kind);
            LogService.Write($"已承诺关机时刻：{_promisedAt:yyyy-MM-dd HH:mm}");
        }

        /// <summary>
        /// 记账：本分钟已提醒过。
        ///
        /// 【保留原因】Poll 内部已自行完成"这一轮关机"的去重（见 _warnedForDue），
        /// 不再依赖调用方记账。这个方法留给调用方在弹框返回后补记一笔日志语义，
        /// 同时保持对外 API 兼容（单元测试仍会调它模拟窗体行为）。
        /// </summary>
        public static void MarkWarned(DateTime now)
        {
            // 兜底：即使调用方没经过 Poll 直接调它，也让同一分钟的重复弹窗被挡住。
            DateTime thisMinute = new DateTime(now.Year, now.Month, now.Day,
                                               now.Hour, now.Minute, 0, now.Kind);
            if (_warnedForDue == DateTime.MinValue) _warnedForDue = thisMinute;
        }

        /// <summary>记账：本分钟已执行关机</summary>
        public static void MarkShutdown(DateTime now)
        {
            _lastShutdownMinute = new DateTime(now.Year, now.Month, now.Day,
                                               now.Hour, now.Minute, 0, now.Kind);
        }

        /// <summary>
        /// 用户选择了"本次不关机"：把这一次的关机时刻记下来，本次不再触发。
        /// 下一天（或下一个命中日）的同一时间仍会正常提醒与执行。
        /// </summary>
        public static void SkipOnce(DateTime dueAt)
        {
            _skipUntil = new DateTime(dueAt.Year, dueAt.Month, dueAt.Day,
                                      dueAt.Hour, dueAt.Minute, 0, dueAt.Kind);
            LogService.Write($"用户选择本次不关机，跳过 {_skipUntil:yyyy-MM-dd HH:mm} 这一次");
        }

        /// <summary>
        /// 下发关机命令。
        ///
        /// 用 shutdown.exe 而不是 P/Invoke ExitWindowsEx：
        /// 前者带系统级的倒计时与"请保存工作"流程，用户还能用 shutdown /a 撤销；
        /// 后者是硬关机，会直接切掉未保存的文档。
        /// </summary>
        public static bool ExecuteShutdown(out string error)
        {
            error = null;
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "shutdown.exe",
                    Arguments = string.Format(
                        "/s /t {0} /c \"陈叔叔希沃优化助手：到预定时间，电脑将在 {1} 分钟内关机。若要取消，请在命令行执行 shutdown /a\"",
                        SystemCountdownSeconds, SystemCountdownSeconds / 60),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };

                System.Diagnostics.Process.Start(psi);
                LogService.Write($"已下发关机命令，系统倒计时 {SystemCountdownSeconds} 秒");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                LogService.Write($"下达关机命令失败：{ex.Message}");
                return false;
            }
        }

        /// <summary>撤销已下发的关机（用于"确认"之后又反悔的场景）</summary>
        public static void AbortShutdown()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "shutdown.exe",
                    Arguments = "/a",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                });
                LogService.Write("已撤销系统关机倒计时");
            }
            catch (Exception ex)
            {
                LogService.Write($"撤销关机失败：{ex.Message}");
            }
        }

        /// <summary>清空内存中的记账状态（用于"重新同步"或设置变更后重排）</summary>
        public static void ResetRuntimeState()
        {
            _warnedForDue = DateTime.MinValue;
            _lastShutdownMinute = DateTime.MinValue;
            _skipUntil = DateTime.MinValue;
            _promisedAt = DateTime.MinValue;
        }

        /// <summary>
        /// 只读快照，供单元测试断言内部状态。
        /// 生产代码不应依赖它——它是为测试留的观察窗。
        /// </summary>
        internal static DateTime PromisedAtSnapshot { get { return _promisedAt; } }

        /// <summary>只读快照：已经为哪一次关机弹过提醒</summary>
        internal static DateTime WarnedForDueSnapshot { get { return _warnedForDue; } }
    }
}
