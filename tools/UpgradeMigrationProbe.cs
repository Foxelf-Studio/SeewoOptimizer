using System;
using System.Collections.Generic;
using Microsoft.Win32;
using SeewoOpt.Services;

namespace SeewoOpt.Tools
{
    /// <summary>
    /// 升级迁移探针 —— "老用户补勾启动自动执行"在**真实注册表**上的回归测试。
    ///
    /// 【要盯住的是什么】
    /// 判定本身（纯策略）有单元测试；这里盯的是注册表这一层：
    ///   1. 写进去的标记，读的时候必须认得出来（同一个值名，一个常量收口）；
    ///   2. **全新安装也要落标记**——这是本批最隐蔽的一条：第一次启动时
    ///      键还不存在，判定正确；但同一趟运行程序自己会把键建起来，
    ///      漏掉标记的话**第二次启动**就把新用户误判成老用户。
    ///      这个错误隔一次启动才暴露，必须端到端跑一遍才算验过。
    ///
    /// 【为什么用真实注册表而不是模拟】
    /// 产品源码里注册表路径是常量，无法注入。与其在探针里复刻一套"等价"读写
    /// （那就是假测试），不如直接在真实 HKCU 上跑——代价是必须做**可回滚**：
    /// 运行前备份全部值，结束时恢复；原本不存在则删键。
    ///
    /// 退出码：0 = 全部通过，1 = 有失败项，2 = 环境不适合运行（拒绝执行）。
    /// </summary>
    internal static class UpgradeMigrationProbe
    {
        private const string KeyPath = @"Software\TimeSyncTool";

        private static int _failed;
        private static int _checks;

        private static void L(string s)
        {
            Console.WriteLine(s);
        }

        private static void Check(string name, bool ok, string detail = null)
        {
            _checks++;
            L(ok ? $"  [通过] {name}" : $"  [失败] {name}");
            if (!ok && !string.IsNullOrEmpty(detail))
                L($"          → {detail}");
            if (!ok) _failed++;
        }

        [STAThread]
        private static int Main()
        {
            Console.WriteLine("=== 升级迁移探针（真实注册表，全程可回滚）===");
            Console.WriteLine();

            bool existedBefore;
            List<ValueBackup> backup;

            try
            {
                existedBefore = KeyExists();
                if (existedBefore && HasSubKeys())
                {
                    // 本探针只备份值、不备份子键树。出现子键说明是一台
                    // 真实使用中的机器（本程序的键从不建子键），拒绝运行更安全。
                    L("设置键下存在子键，为安全起见拒绝运行（本探针只准备处理无子键的键）");
                    return 2;
                }
                backup = existedBefore ? BackupValues() : new List<ValueBackup>();
                L($"运行前状态：设置键 {(existedBefore ? $"已存在，已备份 {backup.Count} 个值" : "不存在（全新机器）")}");
                L("");
            }
            catch (Exception ex)
            {
                L("读取运行前状态失败，拒绝运行：" + ex.Message);
                return 2;
            }

            try
            {
                Run();
            }
            catch (Exception ex)
            {
                L("探针异常：" + ex);
                _failed++;
            }
            finally
            {
                Restore(existedBefore, backup);
                bool okAfter = existedBefore ? KeyExists() : !KeyExists();
                L("");
                L(okAfter
                    ? $"已恢复运行前状态（{(existedBefore ? "还原备份" : "删除本次创建的键")}）"
                    : "警告：恢复运行前状态失败，请手工检查注册表");
                if (!okAfter) _failed++;
            }

            Console.WriteLine();
            Console.WriteLine(_failed == 0
                ? $"全部通过：{_checks} 项断言"
                : $"有 {_failed} 项失败（共 {_checks} 项）");
            return _failed == 0 ? 0 : 1;
        }

        private static void Run()
        {
            // ---------------------------------------------------------
            L("--- 场景 1：全新安装（注册表里什么都没有）---");
            DeleteKey();
            Check("起点：设置键不存在", !KeyExists());

            AutoStartTaskUpgradeVerdict v1 = SettingsStore.DetectAutoStartTaskUpgrade();
            Check("全新安装 → 判为 KeepDefaultForNewInstall",
                v1 == AutoStartTaskUpgradeVerdict.KeepDefaultForNewInstall,
                $"得到 {v1}");

            Check("全新安装落账成功", SettingsStore.TryApplyAutoStartTaskUpgrade(v1));
            Check("标记已写入注册表（下次启动据此认定\"已处理\"）", MarkerPresent());
            Check("全新安装不补勾 AutoStartTask（不替新用户做主）",
                ReadAutoStartTask() == null, "注册表里出现了 AutoStartTask 值");

            // ---------------------------------------------------------
            L("");
            L("--- 场景 2：同一台新机器的**第二次启动**（本批最隐蔽的陷阱）---");
            // 两次启动之间，程序自己会保存设置（Save 会写入 AutoStartTask 等值），
            // 于是"键存在"这个事实对新手安装也成立了。
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                k.SetValue("AutoStartTask", false);

            AutoStartTaskUpgradeVerdict v2 = SettingsStore.DetectAutoStartTaskUpgrade();
            Check("第二次启动 → 判为 AlreadyApplied（不会把新用户误判成老用户）",
                v2 == AutoStartTaskUpgradeVerdict.AlreadyApplied, $"得到 {v2}");

            // ---------------------------------------------------------
            L("");
            L("--- 场景 3：老用户升级（键在、但没有标记）---");
            RemoveMarker();

            AutoStartTaskUpgradeVerdict v3 = SettingsStore.DetectAutoStartTaskUpgrade();
            Check("键在而无标记 → 判为 EnableForUpgrade",
                v3 == AutoStartTaskUpgradeVerdict.EnableForUpgrade, $"得到 {v3}");

            Check("老用户落账成功", SettingsStore.TryApplyAutoStartTaskUpgrade(v3));
            bool? at = ReadAutoStartTask();
            Check("老用户被补勾：注册表 AutoStartTask = true",
                at == true, at == null ? "值为 null（没写进去）" : $"得到 {at}");
            Check("标记同时写入", MarkerPresent());

            // ---------------------------------------------------------
            L("");
            L("--- 场景 4：老用户升级之后的下一次启动 ---");
            AutoStartTaskUpgradeVerdict v4 = SettingsStore.DetectAutoStartTaskUpgrade();
            Check("再次启动 → AlreadyApplied（之后手动取消勾选不会被改回去）",
                v4 == AutoStartTaskUpgradeVerdict.AlreadyApplied, $"得到 {v4}");
        }

        // =================================================================
        // 注册表辅助
        // =================================================================

        private static bool KeyExists()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                return k != null;
        }

        private static bool HasSubKeys()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                return k != null && k.SubKeyCount > 0;
        }

        private static bool MarkerPresent()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                return k != null
                    && k.GetValue(AutoStartTaskUpgradePolicy.MarkerValueName) != null;
        }

        /// <summary>读 AutoStartTask；值不存在返回 null（"没写"与"写成 false"必须能分辨）</summary>
        private static bool? ReadAutoStartTask()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
            {
                if (k == null) return null;
                object v = k.GetValue("AutoStartTask");
                if (v == null) return null;
                return Convert.ToBoolean(v);
            }
        }

        private static void RemoveMarker()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath, true))
            {
                if (k != null)
                    k.DeleteValue(AutoStartTaskUpgradePolicy.MarkerValueName, false);
            }
        }

        private static void DeleteKey()
        {
            using (RegistryKey software = Registry.CurrentUser.OpenSubKey("Software", true))
            {
                if (software != null)
                    software.DeleteSubKeyTree("TimeSyncTool", false);
            }
        }

        // =================================================================
        // 备份 / 恢复
        // =================================================================

        private class ValueBackup
        {
            public string Name;
            public object Value;
            public RegistryValueKind Kind;
        }

        private static List<ValueBackup> BackupValues()
        {
            var list = new List<ValueBackup>();

            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
            {
                if (k == null) return list;

                foreach (string name in k.GetValueNames())
                {
                    list.Add(new ValueBackup
                    {
                        Name = name,
                        Value = k.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames),
                        Kind = k.GetValueKind(name)
                    });
                }
            }

            return list;
        }

        private static void Restore(bool existedBefore, List<ValueBackup> backup)
        {
            try
            {
                DeleteKey();
                if (!existedBefore) return;

                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    foreach (ValueBackup b in backup)
                    {
                        if (b.Value != null)
                            k.SetValue(b.Name, b.Value, b.Kind);
                    }
                }
            }
            catch (Exception ex)
            {
                L("恢复注册表失败：" + ex.Message);
                _failed++;
            }
        }
    }
}
