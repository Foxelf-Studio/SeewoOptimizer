using Microsoft.Win32;
using System;

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

        /// <summary>创建一份默认设置</summary>
        public static AppSettings CreateDefault()
        {
            return new AppSettings
            {
                AutoStart = true,
                SilentStart = false,
                AutoVolume = true,
                VolumeLevel = 60,
                KillWps = true
            };
        }

        /// <summary>把音量限制在有效区间</summary>
        public void Normalize()
        {
            if (VolumeLevel < 0) VolumeLevel = 0;
            if (VolumeLevel > 100) VolumeLevel = 100;
        }
    }

    /// <summary>
    /// 设置的注册表持久化。
    ///
    /// 位置：HKCU\Software\TimeSyncTool
    /// 注意：路径沿用历史值 TimeSyncTool，重命名会导致老用户设置丢失。
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
    }
}
