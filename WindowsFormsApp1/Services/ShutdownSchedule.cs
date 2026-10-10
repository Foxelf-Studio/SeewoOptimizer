using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 一个关机时间点：几点几分 + 在一周的哪几天生效。
    ///
    /// 【为什么用"星期掩码"而不是布尔数组】
    /// 这个对象要存进注册表。注册表里能直接存的只有 int / string 这类标量，
    /// 一个 bool[7] 要么拆成七个键、要么序列化成字符串再解析——两者都会引入
    /// 格式版本与容错问题。把七天压进一个 int 的七个位，读写都是一行，
    /// 也不存在"解析失败"的中间态。
    ///
    /// 【掩码的位序约定】
    /// 位 0 = 周一，位 6 = 周日。这个顺序与 <see cref="DayOfWeek"/> 枚举
    /// **不一致**（后者是 Sunday=0），因此读写都必须走
    /// <see cref="IsDayEnabled"/> / <see cref="SetDay"/> 这类方法，
    /// 绝不直接拿 (int)DateTime.DayOfWeek 去移位——那会整体错开一位，
    /// 表现为"设了周一却周二执行"这种极难察觉的偏差。
    /// </summary>
    public class ShutdownTimeRule
    {
        /// <summary>小时（0-23）</summary>
        public int Hour { get; set; }

        /// <summary>分钟（0-59）</summary>
        public int Minute { get; set; }

        /// <summary>
        /// 星期掩码：位 0=周一 … 位 6=周日，1 表示该天生效。
        /// 全 0 表示这条规则没有任何生效日，等同于一条废规则。
        /// </summary>
        public int DayMask { get; set; }

        /// <summary>全周都生效的掩码（0b1111111 = 127）</summary>
        public const int EveryDayMask = 0x7F;

        /// <summary>七个位对应的中文简称，下标即"位"（0=周一）</summary>
        private static readonly string[] DayNames = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };

        /// <summary>把 DayOfWeek 枚举换算成本类的位号（周一=0 … 周日=6）</summary>
        public static int BitForDayOfWeek(DayOfWeek day)
        {
            // DayOfWeek: Sunday=0, Monday=1 … Saturday=6
            // 目标位号:   Monday=0, Tuesday=1 … Sunday=6
            // 换算：把 Sunday(0) 映射到 6，其余整体减 1。
            return ((int)day + 6) % 7;
        }

        /// <summary>取反：把位号还原成 DayOfWeek（供需要枚举时使用）</summary>
        public static DayOfWeek DayOfWeekForBit(int bit)
        {
            return (DayOfWeek)((bit + 1) % 7);
        }

        /// <summary>某一天是否生效</summary>
        public bool IsDayEnabled(DayOfWeek day)
        {
            return (DayMask & (1 << BitForDayOfWeek(day))) != 0;
        }

        /// <summary>设置某一天是否生效</summary>
        public void SetDay(DayOfWeek day, bool enabled)
        {
            int bit = 1 << BitForDayOfWeek(day);
            if (enabled) DayMask |= bit;
            else DayMask &= ~bit;
        }

        /// <summary>是否每天都生效</summary>
        public bool IsEveryDay
        {
            get { return (DayMask & EveryDayMask) == EveryDayMask; }
        }

        /// <summary>
        /// 把规则拆成"星期几 + 几点几分"，用于判断某个具体时刻是否命中。
        ///
        /// 【为什么抽成纯函数】这是整条关机链路里唯一完全确定的部分：
        /// 给定规则和时刻，命中与否唯一。抽出来后可以构造任意时刻来测边界
        /// （跨日、每分钟只命中一次、掩码错位），不必真的等到那个点。
        /// </summary>
        public bool Matches(DateTime moment)
        {
            return Matches(moment.DayOfWeek, moment.Hour, moment.Minute);
        }

        /// <summary>纯参数版本，便于测试直接构造边界</summary>
        public bool Matches(DayOfWeek day, int hour, int minute)
        {
            return IsDayEnabled(day) && Hour == hour && Minute == minute;
        }

        /// <summary>把时间与星期压成一个整数存注册表（值域 0 .. 167*127+126）</summary>
        public int ToPacked()
        {
            // 编码：分钟数(0-1439) * 128 + 掩码(0-127)
            // 乘数取 128（而非 127）是为了让编码/解码互逆且不产生歧义——
            // 用 127 时掩码 127 会与下一档的掩码 0 读数重叠。
            int minutes = NormalizeHour(Hour) * 60 + NormalizeMinute(Minute);
            return minutes * 128 + (DayMask & EveryDayMask);
        }

        /// <summary>从整数值还原。任何越界都会收敛到合法范围，绝不抛出。</summary>
        public static ShutdownTimeRule FromPacked(int packed)
        {
            if (packed < 0) packed = 0;

            int mask = packed % 128;
            int minutes = packed / 128;

            // 防御：分钟数超过一天时取模，避免损坏的注册表值让
            // 规则落在非法时刻上（越界的时间永远不可能命中，规则会静默失效）
            if (minutes >= 24 * 60) minutes = minutes % (24 * 60);

            return new ShutdownTimeRule
            {
                Hour = minutes / 60,
                Minute = minutes % 60,
                DayMask = mask & EveryDayMask
            };
        }

        private static int NormalizeHour(int h)
        {
            if (h < 0) return 0;
            if (h > 23) return 23;
            return h;
        }

        private static int NormalizeMinute(int m)
        {
            if (m < 0) return 0;
            if (m > 59) return 59;
            return m;
        }

        /// <summary>用于界面显示与日志，例如 "23:30（每天）" 或 "07:00（周一、周三）"</summary>
        public string Describe()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(CultureInfo.InvariantCulture, "{0:D2}:{1:D2}", Hour, Minute);
            sb.Append("（");
            if (IsEveryDay)
            {
                sb.Append("每天");
            }
            else
            {
                bool first = true;
                for (int bit = 0; bit < 7; bit++)
                {
                    if ((DayMask & (1 << bit)) == 0) continue;
                    if (!first) sb.Append("、");
                    sb.Append(DayNames[bit]);
                    first = false;
                }
                if (first) sb.Append("未选择任何日期");
            }
            sb.Append("）");
            return sb.ToString();
        }

        /// <summary>复制一份，避免界面上的编辑直接改到正在生效的规则</summary>
        public ShutdownTimeRule Clone()
        {
            return new ShutdownTimeRule { Hour = Hour, Minute = Minute, DayMask = DayMask };
        }
    }

    /// <summary>
    /// 关机调度的纯逻辑部分：给定一组规则和一个时刻，判断该不该关机。
    ///
    /// 【为什么单独一个类】这段逻辑要能被单元测试直接覆盖。
    /// 若把它塞进窗体里（读控件、开定时器），就只能靠"等到那个点看会不会关机"
    /// 来验证——既慢又不可重复。这里全部是纯函数，不碰 UI、不碰系统时钟。
    /// </summary>
    public static class ShutdownScheduleLogic
    {
        /// <summary>提前多少分钟提醒用户</summary>
        public const int WarnMinutesAhead = 5;

        /// <summary>
        /// 找出在给定时刻应当"立即执行关机"的规则。
        /// 同一时刻最多命中一条：规则按 (时,分) 归一化，同分钟多条只算一次。
        /// </summary>
        public static ShutdownTimeRule FindDueRule(
            IEnumerable<ShutdownTimeRule> rules, DateTime moment)
        {
            if (rules == null) return null;

            foreach (ShutdownTimeRule rule in rules)
            {
                if (rule == null) continue;
                if (rule.Matches(moment)) return rule;
            }
            return null;
        }

        /// <summary>
        /// 计算某条规则"下一次触发"的绝对时刻（用于显示倒计时与排定提醒）。
        ///
        /// from 之后（严格大于 from 的那一分钟起）的第一个命中时刻。
        /// 找不到（例如掩码为空）时返回 null。
        ///
        /// 【为什么要从"下一分钟"起算】若从 from 当前这一分钟起算，
        /// 而 from 恰好就在目标分钟上，会返回 from 自身——调用方据此排提醒
        /// 就会排出"此刻提醒、5 分钟后关机"，但那条规则其实刚刚已经触发过。
        /// 从下一分钟起算可避免重复触发。
        /// </summary>
        public static DateTime? NextOccurrence(ShutdownTimeRule rule, DateTime from)
        {
            if (rule == null) return null;
            if ((rule.DayMask & ShutdownTimeRule.EveryDayMask) == 0) return null;   // 空规则

            // 起点：from 之后的下一分钟，秒与毫秒归零，便于精确比较
            DateTime cursor = new DateTime(from.Year, from.Month, from.Day,
                                           from.Hour, from.Minute, 0, from.Kind)
                                  .AddMinutes(1);

            // 一周 7 天 × 一天 1440 分钟，最多找 8 天必命中（若掩码非空）
            for (int i = 0; i < 8 * 24 * 60; i++)
            {
                if (cursor.Hour == rule.Hour && cursor.Minute == rule.Minute
                    && rule.IsDayEnabled(cursor.DayOfWeek))
                {
                    return cursor;
                }
                cursor = cursor.AddMinutes(1);
            }
            return null;
        }

        /// <summary>
        /// 该不该在 moment 这一刻弹出"5 分钟后关机"的提醒。
        ///
        /// 判定方式：某条规则的下一次触发时刻减去提前量，正好落在 moment。
        /// 这样"提醒"和"关机"共用同一套时刻计算，不会出现两处逻辑各算各的、
        /// 慢慢错开的问题。
        /// </summary>
        public static ShutdownTimeRule FindRuleToWarn(
            IEnumerable<ShutdownTimeRule> rules, DateTime moment)
        {
            if (rules == null) return null;

            // 把提醒时刻归到"分钟"这一档上比较：moment 落在提醒那一分钟内即算命中。
            DateTime thisMinute = new DateTime(moment.Year, moment.Month, moment.Day,
                                               moment.Hour, moment.Minute, 0, moment.Kind);

            foreach (ShutdownTimeRule rule in rules)
            {
                if (rule == null) continue;

                DateTime? next = NextOccurrence(rule, thisMinute.AddMinutes(-1));
                if (!next.HasValue) continue;

                DateTime warnAt = next.Value.AddMinutes(-WarnMinutesAhead);
                if (warnAt == thisMinute) return rule;
            }
            return null;
        }

        /// <summary>
        /// 整体判定：给定全部规则与时刻，返回本轮应当做的动作。
        /// 抽出来是为了让"提醒"与"关机"的优先级关系可以被测试固定住。
        /// </summary>
        public enum ActionKind
        {
            /// <summary>什么都不做</summary>
            None,

            /// <summary>弹出 5 分钟后关机的提醒</summary>
            Warn,

            /// <summary>立即执行关机</summary>
            Shutdown
        }

        public static ActionKind Decide(IEnumerable<ShutdownTimeRule> rules, DateTime moment,
                                        out ShutdownTimeRule hit)
        {
            hit = null;

            // 关机优先于提醒：若两条规则恰好排成"某个规则的提醒点"与
            // "另一条规则的执行点"落在同一分钟，必须先执行关机——
            // 反过来会让程序在该关机的时刻弹出一个提醒框，把关机顶掉。
            ShutdownTimeRule due = FindDueRule(rules, moment);
            if (due != null)
            {
                hit = due;
                return ActionKind.Shutdown;
            }

            ShutdownTimeRule warn = FindRuleToWarn(rules, moment);
            if (warn != null)
            {
                hit = warn;
                return ActionKind.Warn;
            }

            return ActionKind.None;
        }
    }
}
