using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

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
        /// 【为什么是 0——这是用户明确要求的行为变更，与早先设计相反】
        /// 早先取 300（5 分钟），理由是"留一点撤销窗口，且 Windows 会走
        /// 正常的'请保存工作'流程，不会硬切掉课件"。
        ///
        /// 实机截图暴露了那个设计的真实代价：300 秒倒计时**结束后**
        /// Windows 会自己弹出"您将要被注销"的系统框，与我们自己的提醒框
        /// 叠在一起（截图里三个框同屏），用户点哪个都不对；
        /// 而且那 5 分钟里用户的撤销**只能靠命令行 shutdown /a**——
        /// 教室一体机上没人会去敲命令行。
        ///
        /// 用户的最终判断是：**提前 5 分钟的提醒框就是唯一的确认窗口**。
        /// 框里问过了，用户要么点"本次不关机"（那时命令还没下发，
        /// 直接跳过即可），要么不点/点"确认"（到点就该干脆地关掉）。
        /// 到点之后再留一段系统倒计时，只是把决定权从我们的框
        /// 转移到一个用户看不懂的系统框里，纯属多余且制造混乱。
        ///
        /// 因此改为 0：到点立即执行，不再有二次系统倒计时。
        /// 与提醒弹窗的提前量（5 分钟）仍然是两回事，不要混淆。
        /// </summary>
        public const int SystemCountdownSeconds = 0;

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
        /// 关机点过后，允许补执行的宽限分钟数。
        ///
        /// 【为什么需要宽限，而不是只在"那一分钟"执行】
        /// 20 秒巡检只能保证"进程醒着"时命中。以下情况会让程序错过承诺时刻：
        ///   · 系统休眠/睡眠，唤醒时已经是几分钟后；
        ///   · 机器在关机点前后重启，进程重新起来时那一分钟已过；
        ///   · UI 线程被长时间阻塞（例如系统对话框、网络盘卡死）。
        /// 旧实现只认 `当前分钟 == 承诺时刻`，一旦跨过就整天不再关机，
        /// 而且没有任何提示——对一个"教室定点关机"工具来说这是功能失效。
        ///
        /// 【为什么是 10 分钟而不是一直补】
        /// 宽限的意义是"把刚才那一次补上"，不是"今天什么时候开机都关"。
        /// 取 10 分钟：足够覆盖休眠唤醒与重启，又不会让上午开机的机器
        /// 去补执行前一晚的关机。超过宽限只写日志，不再执行。
        /// </summary>
        public const int CatchUpGraceMinutes = 10;

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
        /// 本进程是否**真的下发过**关机命令（<see cref="ExecuteShutdown"/> 成功后置位）。
        ///
        /// 【它回答的问题】用户点"本次不关机"时，系统侧到底有没有东西需要撤销。
        ///
        /// 【为什么必须由本进程记账，而不是"反正调一下 shutdown /a 也无害"】
        /// 2026-10-11 实机日志暴露了"无害"的代价：跳过时无条件调一次
        /// shutdown /a，在没有东西可撤时返回 1116，日志写下
        /// 「撤销关机未生效」——用户看到"未生效"便合理怀疑没取消成功，
        /// 专门来问是不是有缺陷。
        ///
        /// 日志是这套系统唯一的排查凭据（教室机器上没人会去调试它），
        /// 所以"如实反映发生了什么"比"多调一次保险"更重要。
        ///
        /// 【为什么不在 ResetRuntimeState 里清零】
        /// 那个方法用于"设置变更后重排"，而本标志表达的是"本进程曾经下发过
        /// 一次关机"这一历史事实——设置变更不会让已发出的命令消失。
        /// 保持为 true 是更安全的方向：之后的跳过仍会尝试撤销。
        /// </summary>
        private static bool _shutdownIssued = false;

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
            // 看当前时间是否已经到达承诺时刻。到达就关机。
            //
            // 【为什么不再要求"当前分钟恰好等于承诺时刻"】
            // 旧实现写成 `_promisedAt == thisMinute`，等于把执行窗口压缩成
            // 60 秒。休眠唤醒、重启、UI 线程卡顿只要跨过那一分钟，
            // 当天就再也不会关机，且日志里什么都不会留下。
            // 现在改为**区间判定**：承诺时刻起、宽限期内，任意一次巡检都补执行。
            //
            // 【为什么仍然不允许重复执行】
            // _lastShutdownMinute 记录的是"哪一次承诺已执行"（由 MarkShutdown
            // 写入承诺时刻，而不是执行时的当前分钟）。补执行发生在 23:04 时，
            // 记下的仍是 23:00，因此后续巡检不会把同一次关机再发一遍。
            //
            // 【为什么不再从规则实时推导】
            // 因为重启后补弹提醒会顺延（22:56 补弹 → 承诺 23:01），
            // 而规则时刻还是 23:00。若仍按规则推导，23:00 会先关一次，
            // 23:01 又因承诺再关一次。用"承诺时刻"作唯一判据，
            // 23:00 那一刻 _promisedAt 指向的是 23:01，自然不命中，只关一次。
            if (_promisedAt != DateTime.MinValue
                && now >= _promisedAt
                && now < _promisedAt.AddMinutes(CatchUpGraceMinutes)
                && _lastShutdownMinute != _promisedAt
                && _skipUntil != _promisedAt)
            {
                result.ShouldShutdown = true;
                result.Rule = FindRuleForMoment(rules, _promisedAt);
                result.DueAt = _promisedAt;
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

                // 【三道条件各司其职，不要合并】
                //   _warnedForDue != next.Value  —— 这一轮关机已提醒过（窗口内去重的**主防线**）
                //   _skipUntil    != next.Value  —— 用户已选"本次不关机"，别再问
                //
                // 曾经这里还有第三道 `_promisedAt != next.Value`，已删除：
                // 它的本意是"已承诺过就别再弹"，但承诺时刻在**顺延路径**上
                // （22:56 补弹 → 承诺 23:01）并不等于 next（23:00），
                // 于是它在最需要它挡的重启场景里恰好失效——真正挡住重复弹窗的
                // 始终是 _warnedForDue。留着它只会让人误以为它在起作用，
                // 而删掉它测试纹丝不动，正是"冗余防线掩盖真防线"的典型。
                // 【第三道 _skipUntil != _promisedAt 与上面那道不是重复】
                // 上面的比较对象是“规则时刻”（next）；而用户在顺延路径上
                // 跳过的是“承诺时刻”（22:56 弹 → 承诺 23:01），两者不相等。
                // 进程内还有 _warnedForDue 兜住，所以平时看不出差别；
                // 但重启后 _warnedForDue 不落盘，只有 skip 被恢复——
                // 没有这一道，用户刚说过的“本次不关机”会再被问一遍。
                //
                // 必须排除两种“不能采信 skip”的情况：
                //   1. 尚未产生任何承诺：两者都是 MinValue，会被误判成已跳过；
                //   2. 承诺已经过期（超过补执行宽限）：那是昨天那一轮的记账，
                //      不能拿它挡住今天的提醒。
                // 只要承诺还在有效窗口内、且 skip 指向它，就说明这一轮
                // 用户已经明确拒绝，不再询问。
                bool promiseStillRelevant = _promisedAt != DateTime.MinValue
                    && _promisedAt >= now.AddMinutes(-CatchUpGraceMinutes);

                if (next.HasValue && _warnedForDue != next.Value
                    && _skipUntil != next.Value
                    && (!promiseStillRelevant || _skipUntil != _promisedAt))
                {
                    // ===== 承诺时刻 = max(发现时刻所在分钟 + 提前量, 规则时刻) =====
                    //
                    // 正常情况（22:55 弹）恰好等于 next（23:00）；
                    // 重启补弹（22:56 弹）顺延到 23:01。
                    //
                    // 【这个"顺延"是**有意为之**，不是缺陷——改它之前先读完这段】
                    // 公式的取向是"**保证给满提前量**"，规则时刻只作为下限。
                    // 后果：若在提醒窗口内才第一次被发现（程序在窗口内才启动，
                    // 或改设置重启了巡检——WinForms Timer 首次触发要等满一个
                    // 轮询间隔），关机时刻就会晚于用户配置的时间。
                    //
                    // 实测（2026-10-11 实机日志）：规则 08:50，实际 08:51:15 关机；
                    // 成因是 08:45:41 改设置重启巡检，首次巡检落在 20 秒后的
                    // 08:46:01，于是"发现+5 分钟"越过了规则时刻。
                    // 最坏可顺延到 规则时刻 + (提前量 - 1) 分钟。
                    //
                    // 【用户决定：维持现状（2026-10-11）】
                    // 理由是提醒框是**唯一**能点"本次不关机"的机会：
                    // 若不顺延，窗口内才启动的机器会弹出一个一闪而过的框，
                    // 等于没有机会拦截；而对话框上写的就是顺延后的真实时刻
                    // （"电脑将在 5 分钟后关机（08:51）"），用户被如实告知，
                    // 不存在"悄悄晚了"。
                    //
                    // 因此**不要把这里的 max 改成 min 或直接取 next**：
                    // 那会把上述取舍无声地推翻。真要改，先让用户重新决定，
                    // 并同步修改测试 TestShutdownLateWarnSlipsDeliberately
                    // （它把这个顺延钉成了预期行为）。
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

        /// <summary>
        /// 记账：哪一次承诺已经执行过。
        ///
        /// 参数必须是**承诺时刻**（PollResult.DueAt），不是执行时的当前时间。
        /// 补执行可能发生在承诺时刻之后几分钟；若按当前分钟记账，
        /// `_lastShutdownMinute != _promisedAt` 会一直成立，同一轮关机会
        /// 每 20 秒重复下发一次。
        /// </summary>
        public static void MarkShutdown(DateTime dueAt)
        {
            _lastShutdownMinute = new DateTime(dueAt.Year, dueAt.Month, dueAt.Day,
                                               dueAt.Hour, dueAt.Minute, 0, dueAt.Kind);
        }

        /// <summary>「本次不关机」实际做了什么，供调用方如实写日志。</summary>
        public enum SkipOutcome
        {
            /// <summary>只记了跳过标记：本进程从未下发过关机命令，没有东西需要撤销。</summary>
            MarkedOnly,

            /// <summary>记了跳过标记，并尝试撤销系统侧进行中的关机。</summary>
            MarkedAndAbortAttempted
        }

        /// <summary>
        /// 用户选择了"本次不关机"：把这一次的关机时刻记下来，本次不再触发，
        /// 必要时撤销已下发的系统关机。
        /// 下一天（或下一个命中日）的同一时间仍会正常提醒与执行。
        ///
        /// 【为什么要记账而不是只改内存标记——这是个真实存在的洞】
        /// 早先 SkipOnce 只写一个内存标记，假设"取消发生在关机命令下发之前"。
        /// 但实机截图里出现过这一幕：日志区显示"已下达关机命令"，
        /// 同时系统的"您将要被注销"框已经弹出来——说明关机命令已经发出，
        /// 此时只改内存标记，**系统那边的倒计时照样会走完并关机**。
        /// 用户以为点了"本次不关机"就没事了，结果电脑还是关了。
        ///
        /// 【为什么现在大多不需要撤销——2026-10-11 实机日志带来的修正】
        /// /t 0 的设计下，提醒框弹出的那一刻**不会**下发任何关机命令
        /// （见 SystemCountdownSeconds 的注释），命令只在到点时才发。
        /// 因此"在提醒里点本次不关机"这个最常见的动作发生时，
        /// 系统侧本来就没有可撤之物。
        ///
        /// 早先这里无条件下调 AbortShutdown，于是每次跳过都会在日志里留下
        /// "撤销关机未生效（退出码 1116）"——用户看到「未生效」三个字，
        /// 合理地怀疑是不是没取消成功。实机日志确认了这个误解。
        ///
        /// 因此改为**只撤销本进程真的下发过的那一次**：
        ///   · 没下发过 → 不调 shutdown /a，日志直说"无需撤销"；
        ///   · 下发过   → 仍然尽力撤销，这才是那条防线真正存在的意义
        ///                （例如极端时序：到点那一刻与点"本次不关机"撞在一起）。
        ///
        /// 【为什么不撤销"别人发起的关机"】
        /// 老师从开始菜单关机的动作不该被我们取消——那是他的意图。
        /// 系统发起的关机由 ShutdownWarningDialog.ClosedBySessionEnd 处理，
        /// 调用方在那条路径上根本不会走到这里。
        /// </summary>
        public static SkipOutcome SkipOnce(DateTime dueAt)
        {
            _skipUntil = new DateTime(dueAt.Year, dueAt.Month, dueAt.Day,
                                      dueAt.Hour, dueAt.Minute, 0, dueAt.Kind);
            LogService.Write($"用户选择本次不关机，跳过 {_skipUntil:yyyy-MM-dd HH:mm} 这一次");

            if (!ShouldAbortOnSkip())
            {
                // 最常见的情形：在提前 5 分钟的提醒框里点的。
                // 那时命令还没下发，所以"没有需要撤销的系统关机"是**正常且正确**的。
                LogService.Write("尚未下发过关机命令，无需撤销（本次跳过已生效）");
                return SkipOutcome.MarkedOnly;
            }

            // 已经下发过（极端时序，或将来改回带倒计时的设计）：
            // 尽力让系统那边真的停下来。失败不阻断——标记已写入，
            // 至少保证本程序不会再补一刀。
            string abortError;
            if (!AbortShutdown(out abortError))
                LogService.Write($"撤销未确认成功：{abortError}（本次跳过标记已生效）");

            return SkipOutcome.MarkedAndAbortAttempted;
        }

        /// <summary>
        /// 跳过时是否需要去撤销系统侧的关机——即本进程是否真的下发过。
        ///
        /// 单独抽成方法而不是内联那个字段，是为了让单元测试能直接盯着
        /// **这条决策**：测试没法调 ExecuteShutdown（那会真的把关机发出去），
        /// 但可以调 NoteShutdownIssued 模拟"已下发"，再看这里的判断是否翻转。
        /// </summary>
        internal static bool ShouldAbortOnSkip()
        {
            return _shutdownIssued;
        }

        /// <summary>
        /// 记下"本进程已下发过一次关机命令"。
        ///
        /// 由 <see cref="ExecuteShutdown"/> 成功后调用；测试也用它模拟该状态
        /// （唯一区别是测试不会真的执行关机）。
        /// </summary>
        internal static void NoteShutdownIssued()
        {
            _shutdownIssued = true;
        }

        /// <summary>
        /// 解析 shutdown.exe 的**绝对路径**。
        ///
        /// 【为什么绝不能用裸文件名 "shutdown.exe"】
        /// 实机截图报过：执行关机时弹出
        /// `shutdown.exe - 应用程序错误 / 应用程序无法正常启动(0xc0000142)`。
        ///
        /// 0xc0000142 是 STATUS_DLL_INIT_FAILED——**进程自己没起来**，
        /// 不是"关机被某个程序挡住了"（那种是另一个框）。
        /// 根因就是裸文件名：
        ///   · 裸名走的是 CreateProcess 的搜索顺序——先当前目录，
        ///     再 PATH。教室机器上 PATH 常被各种工具塞进奇怪条目，
        ///     只要有一个目录里有同名/损坏的 shutdown.exe，就先被找到，
        ///     那个副本的 DLL 依赖解析不了，于是 0xc0000142。
        ///   · 本程序由计划任务自启，工作目录可能是网络盘或受限目录，
        ///     子进程继承后会进一步影响它的 DLL 搜索。
        ///
        /// 解法：直接用 %SystemRoot%\System32\shutdown.exe 的绝对路径，
        /// 绕开一切搜索顺序。
        ///
        /// 【为什么要看 Sysnative】
        /// 若本程序以 32 位进程运行在 64 位系统上，文件系统重定向会把
        /// System32 映射到 SysWOW64——那里**没有** shutdown.exe。
        /// 此时正解是 Sysnative（32 位视角通往真实 System32 的别名）。
        /// 本程序当前是 AnyCPU、在 64 位系统上跑 64 位进程，走不到这一步；
        /// 但保留这个分支，是为了将来若改成 x86 目标不会突然失效。
        /// </summary>
        internal static string ResolveShutdownExe()
        {
            string system32 = Environment.GetFolderPath(
                Environment.SpecialFolder.System);
            string primary = System.IO.Path.Combine(system32, "shutdown.exe");
            if (System.IO.File.Exists(primary)) return primary;

            // 32 位进程在 64 位系统上的回退：Sysnative 指向真实的 System32
            string sysnative = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Sysnative", "shutdown.exe");
            if (System.IO.File.Exists(sysnative)) return sysnative;

            // 都找不到时退回绝对路径的 primary——交给 Process.Start 报错，
            // 我们照常把异常记进日志，不静默吞掉。
            return primary;
        }

        /// <summary>
        /// 构造 shutdown.exe 的启动参数。
        ///
        /// 两处调用共用，保证"绝对路径 + 固定工作目录"这两条纪律
        /// 不会只落实在一处、另一处又退回裸名。
        /// </summary>
        internal static System.Diagnostics.ProcessStartInfo BuildShutdownStartInfo(string arguments)
        {
            string exe = ResolveShutdownExe();
            return new System.Diagnostics.ProcessStartInfo
            {
                FileName = exe,
                Arguments = arguments,
                // 显式把工作目录钉在 System32：
                //   1. 不继承本程序可能异常的工作目录（见 ResolveShutdownExe）；
                //   2. 让子进程的隐含 DLL 搜索起点落在一个确定且可访问的位置。
                WorkingDirectory = System.IO.Path.GetDirectoryName(exe),
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            };
        }

        /// <summary>
        /// 下发关机命令。
        ///
        /// 用 shutdown.exe 而不是 P/Invoke ExitWindowsEx：
        /// 后者是硬关机、不走系统的"请保存工作"流程，会直接切掉未保存的文档。
        /// shutdown.exe 即使 /t 0 也会走正常的会话结束流程与关机提示。
        ///
        /// 【/t 0 的含义与代价，务必看清】
        /// 倒计时为 0 表示**立即执行**，没有可撤销的窗口：
        /// 命令一下发，系统就开始关会话、结束进程，`shutdown /a` 此时
        /// 已经救不回来了（没有"进行中的关机"可中止）。
        ///
        /// 这正是用户要的行为——**提前 5 分钟的提醒框就是唯一的确认窗口**。
        /// 因此不要在调用本方法之后还指望能撤销，见 <see cref="AbortShutdown"/> 的注释。
        /// </summary>
        public static bool ExecuteShutdown(out string error)
        {
            error = null;
            try
            {
                // 不带 /c 注释：/t 0 时系统不会弹"将要关机"的提示框
                // （那条提示只在有倒计时时出现），带注释也无处显示。
                var psi = BuildShutdownStartInfo(
                    string.Format("/s /t {0}", SystemCountdownSeconds));

                System.Diagnostics.Process.Start(psi);

                // 记账：之后的"本次不关机"才知道系统侧确有东西可撤。
                NoteShutdownIssued();

                LogService.Write(SystemCountdownSeconds > 0
                    ? $"已下发关机命令（{psi.FileName}），系统倒计时 {SystemCountdownSeconds} 秒"
                    : $"已下发关机命令（{psi.FileName}），立即执行（系统倒计时 0 秒）");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                LogService.Write($"下达关机命令失败：{ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 撤销已下发的关机倒计时。
        ///
        /// 【什么时候它真的能撤销，什么时候不能——别误解】
        /// 只有"系统手头有一个正在倒计时的关机任务"时，`shutdown /a` 才有意义。
        /// 在当前设计下（<see cref="SystemCountdownSeconds"/> = 0），
        /// 关机命令一下发就立即执行，**几乎没有可撤销的窗口**。
        ///
        /// 那这个方法还有什么用？——两处真实用途：
        ///   1. 防御性撤销：极端时序下"本次不关机"与"到点下发命令"撞在一起时，
        ///      必须仍能尽力拦住，而不是只改一个内存标记；
        ///   2. 将来若把倒计时改回非 0（或系统策略强制加缓冲），
        ///      这一步是唯一能救命的地方。
        ///
        /// 由 <see cref="SkipOnce"/> 调用，且只在<see cref="ShouldAbortOnSkip"/>
        /// 为真（本进程确实下发过）时才调——没下发过就别去调，
        /// 那只会往日志里写一条让人误以为失败的记录。详见 SkipOnce 的注释。
        ///
        /// 调用它是安全的：没有进行中的关机时不产生任何副作用。
        /// </summary>
        public static bool AbortShutdown(out string error)
        {
            error = null;
            try
            {
                // 与 ExecuteShutdown 共用同一个构建方法：
                // 撤销同样必须走绝对路径，否则可能出现
                // "关机用的是真的 shutdown.exe、撤销却调到了别的副本"——
                // 那种半对半错最难排查。
                var psi = BuildShutdownStartInfo("/a");

                using (var p = System.Diagnostics.Process.Start(psi))
                {
                    // 【为什么要等它退出并读 ExitCode】
                    // 不等它结束就直接报"撤销成功"是在骗人——
                    // 必须拿到真实的退出码，才能判断到底撤掉没有。
                    if (p != null && !p.WaitForExit(5000))
                    {
                        error = "撤销关机命令超时未返回";
                        LogService.Write(error);
                        return false;
                    }

                    int code = (p != null && p.HasExited) ? p.ExitCode : -1;

                    if (code == 0)
                    {
                        LogService.Write("已撤销系统侧的关机倒计时");
                        return true;
                    }

                    if (code == ERROR_NO_SHUTDOWN_IN_PROGRESS)
                    {
                        // 【措辞修正，来自实机日志】
                        // 这里原先写"撤销关机未生效"，用户看到"未生效"就以为
                        // 取消失败了。它的真实含义是"本来就没有进行中的关机"——
                        // 对我们而言恰恰是好消息：机器不会关。
                        LogService.Write("系统侧没有进行中的关机任务，无需撤销");
                        return true;
                    }

                    // 其余退出码：可能是真失败，也可能只是没东西可撤。
                    // 如实记下原始码（负数按十六进制打印，便于和
                    // 0xC0000142 这类系统状态码对照），不替用户下结论。
                    LogService.Write($"撤销关机命令返回退出码 {DescribeExitCode(code)}，未确认成功");
                    return false;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                LogService.Write($"撤销关机失败：{ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 没有进行中的关机任务时 shutdown.exe 的退出码
        /// （ERROR_NO_SHUTDOWN_IN_PROGRESS）。实机日志里大量出现，属正常。
        /// </summary>
        private const int ERROR_NO_SHUTDOWN_IN_PROGRESS = 1116;

        /// <summary>
        /// 把退出码写成便于对照的形式：负数（常见于进程根本没起来）按
        /// 十六进制输出。例如 -1073741502 → 0xC0000142（DLL 初始化失败），
        /// 一眼就能和系统状态码对上；只印十进制则会让人一头雾水。
        /// </summary>
        internal static string DescribeExitCode(int code)
        {
            return code < 0
                ? string.Format("0x{0:X8}", unchecked((uint)code))
                : code.ToString();
        }

        /// <summary>清空内存中的记账状态（用于"重新同步"或设置变更后重排）</summary>
        public static void ResetRuntimeState()
        {
            _warnedForDue = DateTime.MinValue;
            _lastShutdownMinute = DateTime.MinValue;
            _skipUntil = DateTime.MinValue;
            _promisedAt = DateTime.MinValue;
        }

        // ============================================================
        // 状态持久化
        //
        // 【为什么需要落盘】
        // 承诺时刻、跳过标记、已执行标记原先只在内存里。进程一退出
        // （用户重启、崩溃、更新后重启），这些状态全部丢失：
        //   · 重启发生在关机点之后 → 当天不会再关；
        //   · 用户点过“本次不关机”后崩溃重启 → 会再问一次；
        //   · 补执行时无法判断“这一次是不是已经关过了”。
        // 落盘后，承诺与跳过都能跨重启保留，配合 CatchUpGraceMinutes
        // 才能真正覆盖“休眠唤醒/重启/卡顿”这三类漏关场景。
        //
        // 【为什么不持久化 _warnedForDue】
        // “重启后补弹提醒并把承诺顺延为现在+5分钟”是用户明确要的行为，
        // 测试 TestShutdownPromiseAfterRestart 把它钉住了。若把“已提醒”
        // 也持久化，重启后就不会补弹，等于悄悄推翻那条决定。
        // 因此这里只保存“承诺/跳过/已执行”三笔账。
        // ============================================================

        private const string StateVersion = "v1";
        private const string StateTimeFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// 把当前记账导出为文本，供窗体落盘。
        /// 只做纯序列化，不碰文件系统——测试工程链接本文件时不会写用户磁盘。
        /// </summary>
        public static string ExportStateText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(StateVersion);
            sb.AppendLine("promised=" + FormatStateTime(_promisedAt));
            sb.AppendLine("skip=" + FormatStateTime(_skipUntil));
            sb.AppendLine("last=" + FormatStateTime(_lastShutdownMinute));
            return sb.ToString();
        }

        /// <summary>
        /// 从窗体读到的文本恢复记账状态。
        ///
        /// 【过期处理】承诺时刻早于 now - 宽限 的，一律丢弃：
        /// 那是上一次运行遗留的，补执行已经没有意义，留着反而会在
        /// 下一次巡检时突然关机。跳过标记同理。
        ///
        /// 返回一段可读摘要，供调用方写日志；不抛异常。
        /// </summary>
        public static string ImportStateText(string text, DateTime now)
        {
            _promisedAt = DateTime.MinValue;
            _skipUntil = DateTime.MinValue;
            _lastShutdownMinute = DateTime.MinValue;
            _warnedForDue = DateTime.MinValue;

            DateTime promised, skip, last;
            if (!TryParseStateText(text, out promised, out skip, out last))
                return "状态文件格式不正确，已忽略";

            List<string> notes = new List<string>();

            if (promised != DateTime.MinValue)
            {
                if (promised >= now.AddMinutes(-CatchUpGraceMinutes))
                {
                    _promisedAt = promised;
                    notes.Add(string.Format("恢复关机承诺 {0:yyyy-MM-dd HH:mm}", promised));
                }
                else
                {
                    notes.Add(string.Format(
                        "上次关机承诺 {0:yyyy-MM-dd HH:mm} 已超过 {1} 分钟宽限，放弃补执行",
                        promised, CatchUpGraceMinutes));
                }
            }

            if (skip != DateTime.MinValue)
            {
                if (skip >= now.AddMinutes(-CatchUpGraceMinutes))
                {
                    _skipUntil = skip;
                    notes.Add(string.Format("恢复“本次不关机”标记 {0:yyyy-MM-dd HH:mm}", skip));
                }
                else
                {
                    notes.Add("上次“本次不关机”标记已过期，忽略");
                }
            }

            if (last != DateTime.MinValue)
                _lastShutdownMinute = last;

            if (notes.Count == 0) return "无有效记录";
            return string.Join("；", notes.ToArray());
        }

        private static string FormatStateTime(DateTime value)
        {
            return value == DateTime.MinValue
                ? string.Empty
                : value.ToString(StateTimeFormat, CultureInfo.InvariantCulture);
        }

        private static bool TryParseStateText(
            string text, out DateTime promised, out DateTime skip, out DateTime last)
        {
            promised = DateTime.MinValue;
            skip = DateTime.MinValue;
            last = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(text)) return true;

            string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return true;

            if (!string.Equals(lines[0].Trim(), StateVersion, StringComparison.Ordinal))
                return false;

            bool sawPromised = false, sawSkip = false, sawLast = false;

            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0) continue;

                int eq = line.IndexOf('=');
                if (eq <= 0) return false;

                string key = line.Substring(0, eq).Trim();
                string raw = line.Substring(eq + 1).Trim();

                DateTime parsed;
                if (!TryParseStateTime(raw, out parsed)) return false;

                switch (key)
                {
                    case "promised": promised = parsed; sawPromised = true; break;
                    case "skip": skip = parsed; sawSkip = true; break;
                    case "last": last = parsed; sawLast = true; break;
                    default: break;      // 未来新增字段：忽略，保持向前兼容
                }
            }

            // 三个字段必须齐全；缺一个就当作损坏，宁可清空也不要猜。
            return sawPromised && sawSkip && sawLast;
        }

        private static bool TryParseStateTime(string raw, out DateTime value)
        {
            value = DateTime.MinValue;
            if (string.IsNullOrEmpty(raw)) return true;   // 空 = 未记录
            return DateTime.TryParseExact(
                raw, StateTimeFormat, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out value);
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
