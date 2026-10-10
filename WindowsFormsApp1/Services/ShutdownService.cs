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
        /// 已提醒过的关机时刻，防止同一分钟反复弹窗。
        ///
        /// 轮询间隔 20 秒，而判断是按分钟粒度——若不去重，
        /// 在提醒那一分钟内 20 秒、40 秒、60 秒会各弹一次，共三次。
        /// 记录"已经提醒过的那个绝对分钟"即可。
        /// </summary>
        private static DateTime _lastWarnedMinute = DateTime.MinValue;

        /// <summary>已执行过关机的时刻，防止重复下发 shutdown 命令</summary>
        private static DateTime _lastShutdownMinute = DateTime.MinValue;

        /// <summary>上一次提醒用户选择"本次不关机"所对应的关机时刻</summary>
        private static DateTime _skipUntil = DateTime.MinValue;

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
        /// 执行一次巡检。纯判断，不产生任何副作用（不改状态、不碰系统）。
        ///
        /// 【为什么要分成"判断"与"记账"两步】
        /// 调用方在收到 ShouldWarn 后要弹一个模态框，用户可能看很久。
        /// 若在判断时就更新 _lastWarnedMinute，那么弹窗期间的下一次巡检
        /// 会因为"已提醒过"而跳过——看似合理；但如果弹窗被用户直接关掉，
        /// 状态已经消耗掉了，这一分钟的提醒就永远丢了。
        /// 因此这里只读不写，由调用方在**确认处理完之后**调 Mark* 记账。
        /// </summary>
        public static PollResult Poll(IEnumerable<ShutdownTimeRule> rules, DateTime now)
        {
            var result = new PollResult();

            DateTime thisMinute = new DateTime(now.Year, now.Month, now.Day,
                                               now.Hour, now.Minute, 0, now.Kind);

            // ---- 先看是否该关机 ----
            //
            // 用户在本轮的"5 分钟提醒"里选了"本次不关机"，则该次关机作废。
            // 判断放在这里而不是直接删规则：规则是长期设置，
            // "本次不关机"只对这一次生效，明天同一时间仍然要关。
            if (_lastShutdownMinute != thisMinute && _skipUntil != thisMinute)
            {
                ShutdownTimeRule due = ShutdownScheduleLogic.FindDueRule(rules, now);
                if (due != null)
                {
                    result.ShouldShutdown = true;
                    result.Rule = due;
                    result.DueAt = thisMinute;
                    return result;
                }
            }

            // ---- 再看是否该提醒 ----
            if (_lastWarnedMinute != thisMinute)
            {
                ShutdownTimeRule warnRule = ShutdownScheduleLogic.FindRuleToWarn(rules, now);
                if (warnRule != null)
                {
                    DateTime? next = ShutdownScheduleLogic.NextOccurrence(
                        warnRule, thisMinute.AddMinutes(-1));

                    // next 只可能为 null 当规则掩码为空，届时不必提醒
                    if (next.HasValue && _skipUntil != next.Value)
                    {
                        result.ShouldWarn = true;
                        result.Rule = warnRule;
                        result.DueAt = next.Value;
                    }
                }
            }

            return result;
        }

        /// <summary>记账：本分钟已提醒过</summary>
        public static void MarkWarned(DateTime now)
        {
            _lastWarnedMinute = new DateTime(now.Year, now.Month, now.Day,
                                             now.Hour, now.Minute, 0, now.Kind);
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
            _lastWarnedMinute = DateTime.MinValue;
            _lastShutdownMinute = DateTime.MinValue;
            _skipUntil = DateTime.MinValue;
        }
    }
}
