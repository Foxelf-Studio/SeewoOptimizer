using System;

namespace SeewoOpt.Services
{
    /// <summary>
    /// "守护天数"起始日的可信度判定。
    ///
    /// 【为什么需要它——实机缺陷】
    /// 教室机器的时钟常停在 2000-01-01（主板电池没电 / 镜像未校时）。
    /// 程序启动时在那一刻把"首次运行日期"记了下来（2000-01-01），
    /// 等 NTP 把时钟校准到 2026-10-11 之后再看这个起点，
    /// 算出来就是 **9781 天**——用户实机截图里正是这个数字。
    ///
    /// 根本问题是：**把不可信的时钟当成了事实来源**。
    /// 时钟错的时候，"今天是几号"是错的，"首次运行是哪天"自然也错；
    /// 而一旦错误的值被写进注册表，它就会一直错下去。
    ///
    /// 修法分两层：
    ///   1. 不再在校准之前记录起始日（见 TimeSyncForm 的调用点）；
    ///   2. 已经写在注册表里的旧值，用这份判定剔除**不可能成立**的：
    ///      · 晚于今天      → 时钟曾经走在未来
    ///      · 早于功能上线日 → 本程序那时还不存在这件事
    ///
    /// 【为什么要单独抽一个类型，而不是写在窗体里】
    /// 窗体链不进测试工程（依赖太多）。抽成纯函数后，
    /// 单元测试链接的是**同一份源码**，可以直接拿实机那个
    /// 2000-01-01 来断言——而不是复刻一段"看起来一样"的判断。
    /// 本项目已多次确认：复刻的副本跑绿说明不了任何事。
    /// </summary>
    public static class GuardDayPolicy
    {
        /// <summary>
        /// 守护天数功能上线的日期（提交 c76187a，2026-10-10）。
        ///
        /// 【为什么这个常量永远不会过期】
        /// 它表达的是一个**不可动摇的历史事实**：本程序在 2026-10-10
        /// 之前没有这个功能，因此"首次运行日期"早于这一天必然是错的。
        /// 无论过多少年，这条判断都依然成立。
        /// </summary>
        public static readonly DateTime FeatureEpoch = new DateTime(2026, 10, 10);

        /// <summary>起始日的可信度判定结果</summary>
        public enum FirstRunDateVerdict
        {
            /// <summary>尚未记录过（全新安装，或记录被清掉）</summary>
            Missing,

            /// <summary>看起来成立，可以采信</summary>
            Trustworthy,

            /// <summary>晚于今天——说明它是在时钟走在未来时记下的</summary>
            InFuture,

            /// <summary>早于功能上线日——那时本程序还不存在这个功能</summary>
            BeforeFeatureEpoch
        }

        /// <summary>
        /// 判定一个已记录的起始日是否可信。
        ///
        /// 只剔除**逻辑上不可能成立**的值，不做"看起来太久"这类主观判断——
        /// 一个装了两年的机器显示 700 多天是完全正常的，不该被动。
        /// </summary>
        public static FirstRunDateVerdict Judge(DateTime firstRunDate, DateTime now)
        {
            if (firstRunDate == DateTime.MinValue)
                return FirstRunDateVerdict.Missing;

            DateTime recorded = firstRunDate.Date;
            DateTime today = now.Date;

            if (recorded > today)
                return FirstRunDateVerdict.InFuture;

            if (recorded < FeatureEpoch.Date)
                return FirstRunDateVerdict.BeforeFeatureEpoch;

            return FirstRunDateVerdict.Trustworthy;
        }

        /// <summary>
        /// 给出应当采用的起始日：可信就沿用，否则一律改用**今天**。
        ///
        /// 【为什么重置为今天，而不是"猜一个更合理的过去日期"】
        /// 真正的首次运行日期已经丢失了——当时那个时钟本身就不可信，
        /// 没有任何依据能把它还原。改为今天意味着计数从第 0 天重新开始，
        /// 这是唯一诚实的选项：它至少不会显示一个编造的、很大的天数。
        /// </summary>
        public static DateTime ResolveTrustedFirstRunDate(
            DateTime firstRunDate, DateTime now, out FirstRunDateVerdict verdict)
        {
            verdict = Judge(firstRunDate, now);
            return verdict == FirstRunDateVerdict.Trustworthy
                ? firstRunDate.Date
                : now.Date;
        }
    }
}
