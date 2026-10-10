using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Text;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 应用设置的内存模型。
    /// 独立于窗体存在，便于设置窗体与主窗体之间传递数据，也便于单元测试。
    /// </summary>
    public class AppSettings
    {
        /// <summary>是否创建开机自启计划任务</summary>
        public bool AutoStart { get; set; }

        /// <summary>启动时是否静默隐藏到托盘</summary>
        public bool SilentStart { get; set; }

        /// <summary>是否自动调节系统音量</summary>
        public bool AutoVolume { get; set; }

        /// <summary>目标音量百分比（0-100）</summary>
        public int VolumeLevel { get; set; }

        /// <summary>是否在启动时结束 WPS 相关进程</summary>
        public bool KillWps { get; set; }

        /// <summary>
        /// 程序启动时是否自动开始执行既定任务。
        ///
        /// 【默认 false】用户明确要求"启动时不再默认自动执行"——
        /// 空教室的一体机被远程唤醒或课间重启时，不该自己动起来。
        /// 这条默认值本身就是需求的一部分，不要"顺手"改成 true。
        /// </summary>
        public bool AutoStartTask { get; set; }

        /// <summary>
        /// 首次运行日期（用于"已守护本电脑 x 天"）。
        ///
        /// 【为什么记"首次运行"而不是"安装"】
        /// 本程序没有安装程序，是绿色单文件。用户在资源管理器里双击那天
        /// 就是第一次运行，也就是"守护"的起点。卸载后重装会从第 1 天重算，
        /// 这与"守护这台电脑多少天"的语义一致。
        ///
        /// 用 DateTime.MinValue 表示"尚未记录"——首次运行时补写当天。
        /// </summary>
        public DateTime FirstRunDate { get; set; }

        /// <summary>
        /// 关机时间点，最多 <see cref="MaxShutdownRules"/> 条。
        /// 空列表表示不启用关机功能。
        /// </summary>
        public List<ShutdownTimeRule> ShutdownRules { get; set; }

        /// <summary>关机时间点的数量上限</summary>
        public const int MaxShutdownRules = 5;

        /// <summary>创建一份默认设置</summary>
        public static AppSettings CreateDefault()
        {
            return new AppSettings
            {
                AutoStart = true,
                SilentStart = false,
                AutoVolume = true,
                VolumeLevel = 60,
                KillWps = true,
                AutoStartTask = false,          // 需求：启动时不默认自动执行
                FirstRunDate = DateTime.MinValue,
                ShutdownRules = new List<ShutdownTimeRule>()
            };
        }

        /// <summary>把音量限制在有效区间，并收紧关机规则</summary>
        public void Normalize()
        {
            if (VolumeLevel < 0) VolumeLevel = 0;
            if (VolumeLevel > 100) VolumeLevel = 100;

            if (ShutdownRules == null)
                ShutdownRules = new List<ShutdownTimeRule>();

            // 上限必须在**这里**收紧，而不是只靠界面限制。
            // 注册表是用户可改的，界面的数字上下限拦不住手工写入；
            // 数量失控会让每次启动都要遍历一长串规则。
            while (ShutdownRules.Count > MaxShutdownRules)
                ShutdownRules.RemoveAt(ShutdownRules.Count - 1);
        }

        /// <summary>
        /// 已守护的天数（含第 1 天）。
        ///
        /// 【为什么把起始日按"日期"而不是"时刻"相减】
        /// 若用 (今天-起始).TotalDays 直接取整，那么"昨天 23:00 首次运行、
        /// 今天 08:00 打开"会算出 0 天——明明是第二天了。改为把两端都归到
        /// 当日零点再相减，跨过午夜就算一天。
        ///
        /// 结果恒 >= 1：首次运行当天显示"第 1 天"，不显示"第 0 天"。
        /// 起始日缺失（旧版本升级上来）时按今天算，同样是第 1 天。
        /// </summary>
        public int GuardedDays(DateTime today)
        {
            if (FirstRunDate == DateTime.MinValue) return 1;

            DateTime startDay = FirstRunDate.Date;
            DateTime todayDay = today.Date;

            int days = (int)(todayDay - startDay).TotalDays + 1;
            return days < 1 ? 1 : days;
        }
    }

    /// <summary>
    /// 设置的注册表持久化。
    ///
    /// 位置：HKCU\Software\TimeSyncTool
    ///
    /// 【为什么不跟着命名空间改名】
    /// 代码命名空间已统一为 SeewoOpt，但这里**必须保留 TimeSyncTool**。
    /// 命名空间是编译期概念，改名只影响源码；注册表路径是运行时数据，
    /// 改了等于换了一个存储位置——已装用户的全部设置会被静默重置为默认值。
    /// 二者不可混淆。同理见 AutoStartService.TaskName 与更新缓存目录。
    /// </summary>
    public static class SettingsStore
    {
        private const string REGISTRY_PATH = @"Software\TimeSyncTool";

        /// <summary>
        /// 从注册表读取设置。
        ///
        /// 修正记录：原实现在 catch 中静默套用默认值且不记日志，
        /// 导致注册表权限异常或值损坏时，用户改过的配置无声消失。
        /// 现在会记录失败原因，并对单个字段做类型容错。
        /// </summary>
        public static AppSettings Load()
        {
            AppSettings settings = AppSettings.CreateDefault();

            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REGISTRY_PATH))
                {
                    if (key == null)
                    {
                        LogService.Write("注册表中无设置项，使用默认设置");
                        return settings;
                    }

                    settings.AutoStart = ReadBool(key, "AutoStart", settings.AutoStart);
                    settings.SilentStart = ReadBool(key, "SilentStart", settings.SilentStart);
                    settings.AutoVolume = ReadBool(key, "AutoVolume", settings.AutoVolume);
                    settings.VolumeLevel = ReadInt(key, "VolumeLevel", settings.VolumeLevel);
                    settings.KillWps = ReadBool(key, "KillWps", settings.KillWps);
                    settings.AutoStartTask = ReadBool(key, "AutoStartTask", settings.AutoStartTask);
                    settings.FirstRunDate = ReadDate(key, "FirstRunDate", settings.FirstRunDate);
                    settings.ShutdownRules = ReadShutdownRules(key);
                }
            }
            catch (Exception ex)
            {
                LogService.Write($"读取设置失败，改用默认设置: {ex.Message}");
            }

            settings.Normalize();
            return settings;
        }

        /// <summary>将设置写回注册表</summary>
        public static void Save(AppSettings settings)
        {
            if (settings == null) throw new ArgumentNullException("settings");

            settings.Normalize();

            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    key.SetValue("AutoStart", settings.AutoStart);
                    key.SetValue("SilentStart", settings.SilentStart);
                    key.SetValue("AutoVolume", settings.AutoVolume);
                    key.SetValue("VolumeLevel", settings.VolumeLevel);
                    key.SetValue("KillWps", settings.KillWps);
                    key.SetValue("AutoStartTask", settings.AutoStartTask);

                    // 首次运行日期只在"尚未记录"时写入一次。
                    // 若每次都覆盖，卸载重装之外的任何情况都会被重置成今天，
                    // "已守护 x 天"会永远显示 1 天——而它恰恰是每天都要看的数字。
                    if (settings.FirstRunDate != DateTime.MinValue)
                        key.SetValue("FirstRunDate", FormatDate(settings.FirstRunDate));

                    // 关机规则存成一个 REG_SZ，各条之间用 ';' 分隔，每条是一个整数编码。
                    // 用单值而非五个独立键：规则数量是可变的，用定长键名会留下
                    // "删掉第 3 条后第 4、5 条要不要前移"这类歧义。
                    key.SetValue("ShutdownRules", SerializeRules(settings.ShutdownRules));
                }
            }
            catch (Exception ex)
            {
                LogService.Write($"保存设置失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>删除注册表设置项（卸载时调用）</summary>
        public static void Delete()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software", true))
            {
                if (key != null)
                    key.DeleteSubKeyTree("TimeSyncTool", false);
            }
        }

        /// <summary>读取布尔值，容忍 REG_SZ 等类型不一致的情况</summary>
        private static bool ReadBool(RegistryKey key, string name, bool defaultValue)
        {
            object value = key.GetValue(name, defaultValue);
            if (value is bool) return (bool)value;

            try
            {
                return Convert.ToBoolean(value);
            }
            catch
            {
                LogService.Write($"设置项 {name} 类型异常({value?.GetType().Name ?? "null"})，使用默认值 {defaultValue}");
                return defaultValue;
            }
        }

        /// <summary>读取整数，容忍类型不一致的情况</summary>
        private static int ReadInt(RegistryKey key, string name, int defaultValue)
        {
            object value = key.GetValue(name, defaultValue);
            if (value is int) return (int)value;

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                LogService.Write($"设置项 {name} 类型异常({value?.GetType().Name ?? "null"})，使用默认值 {defaultValue}");
                return defaultValue;
            }
        }

        /// <summary>
        /// 读取日期。统一用 "yyyy-MM-dd" 存，避免受系统区域设置影响。
        ///
        /// 【为什么不用 DateTime 直接存注册表】
        /// 虽然 RegistryKey 支持直接写 DateTime，但写入格式取决于 .NET 内部
        /// 表示，跨版本、跨机器读回时不好核对；文本格式在注册表编辑器里
        /// 肉眼可读可改，排查"守护天数不对"时这一点很重要。
        /// </summary>
        private static DateTime ReadDate(RegistryKey key, string name, DateTime defaultValue)
        {
            object value = key.GetValue(name);
            if (value == null) return defaultValue;

            if (value is DateTime) return (DateTime)value;

            DateTime parsed;
            if (DateTime.TryParse(Convert.ToString(value),
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out parsed))
                return parsed;

            LogService.Write($"设置项 {name} 不是合法日期（值为 {value}），忽略");
            return defaultValue;
        }

        private static string FormatDate(DateTime date)
        {
            return date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 把关机规则序列化成一行文本：每条规则一个整数，用 ';' 分隔。
        ///
        /// 空列表序列化为空串（而不是 null）——注册表里 REG_SZ 存 null
        /// 与空串在读取时表现不同，统一成空串可避免"有没有这个键"
        /// 与"键里有没有值"两种空状态混在一起。
        /// </summary>
        internal static string SerializeRules(List<ShutdownTimeRule> rules)
        {
            if (rules == null || rules.Count == 0) return string.Empty;

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < rules.Count; i++)
            {
                if (rules[i] == null) continue;
                if (sb.Length > 0) sb.Append(';');
                sb.Append(rules[i].ToPacked().ToString(
                    System.Globalization.CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }

        /// <summary>
        /// 解析关机规则。任何一条损坏都**只跳过那一条**，不影响其余规则。
        ///
        /// 【为什么不是整体失败就回退到默认】
        /// 规则是用户手工配置的，一条写坏就让全部关机设置消失，代价太大。
        /// 逐条隔离后，用户至少还能保住其余几条，也能从日志里看出是哪一条出的问题。
        /// </summary>
        internal static List<ShutdownTimeRule> DeserializeRules(string raw)
        {
            var result = new List<ShutdownTimeRule>();
            if (string.IsNullOrEmpty(raw)) return result;

            string[] parts = raw.Split(';');
            foreach (string part in parts)
            {
                string token = part.Trim();
                if (token.Length == 0) continue;

                int packed;
                if (!int.TryParse(token, System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture, out packed))
                {
                    LogService.Write($"关机规则项无法解析，已跳过：\"{token}\"");
                    continue;
                }

                result.Add(ShutdownTimeRule.FromPacked(packed));
                if (result.Count >= AppSettings.MaxShutdownRules) break;
            }
            return result;
        }

        /// <summary>从注册表读取并解析关机规则</summary>
        private static List<ShutdownTimeRule> ReadShutdownRules(RegistryKey key)
        {
            object value = key.GetValue("ShutdownRules");
            if (value == null) return new List<ShutdownTimeRule>();

            // 兼容有人手工把值写成 REG_MULTI_SZ 的情况
            string[] multi = value as string[];
            if (multi != null)
                return DeserializeRules(string.Join(";", multi));

            return DeserializeRules(Convert.ToString(value));
        }
    }
}
