namespace SeewoOpt.Services
{
    /// <summary>
    /// "升级迁移"（把启动自动执行补勾给老用户）的判定结果。
    /// </summary>
    public enum AutoStartTaskUpgradeVerdict
    {
        /// <summary>迁移已经处理过（注册表里有标记），什么都不用做</summary>
        AlreadyApplied,

        /// <summary>老用户升级：注册表里已有设置 → 这台机器跑过旧版本 → 替他补勾</summary>
        EnableForUpgrade,

        /// <summary>全新安装：注册表里什么都没有 → 保持默认不勾</summary>
        KeepDefaultForNewInstall,
    }

    /// <summary>
    /// 升级迁移策略：从旧版本升级上来的用户要自动勾选
    /// "程序启动时自动开始执行既定任务"。
    ///
    /// 【为什么需要这次迁移】
    /// v26.6.9 及以前**没有这个开关**——程序启动 2 秒后无条件自动开始同步。
    /// 新版把它做成了一个默认**不勾**的选项（空教室的一体机被远程唤醒或
    /// 课间重启时，不该自己动起来）。于是老用户升级后会"感觉程序不干活了"：
    /// 以前开机就自己跑，现在必须有人去点一下"开始执行"。
    /// 他们要的是原来的体验，所以升级时要替他们把这一项补勾上；
    /// 但**全新安装**的用户不能被替做决定。
    ///
    /// 【怎么区分"老用户升级"与"全新安装"】
    /// 旧版没有记录任何可以直接读的版本号，所以用**注册表设置键是否存在**：
    ///   - 旧版（v26.6.9）的 LoadSettings 用的是 CreateSubKey——只要启动过，
    ///     哪怕用户从没改过设置，也必然创建 HKCU\Software\TimeSyncTool；
    ///   - 新版的 Load 用的是 OpenSubKey，全新安装时该键不存在。
    /// 再加上一个**一次性标记**（见 <see cref="MarkerValueName"/>）：
    /// 迁移处理完成后写入，此后无论键是否存在都不再重复处理。
    ///
    /// 【为什么"全新安装"也必须写标记——这是必须钉死的陷阱】
    /// 全新安装第一次启动：键不存在 → 判定为 KeepDefaultForNewInstall。
    /// 但同一趟运行里，程序自己会把键建起来（保存设置、记录守护起始日）。
    /// 如果这时**不写标记**，第二次启动时就是"键存在 → 判成老用户 → 自动勾选"，
    /// 新用户被悄悄改了设置——而且这个 Bug 只在"第二次启动"才出现，
    /// 第一次测试永远是好的。所以标记必须对"非 AlreadyApplied"的**所有**判定
    /// 都写，这就是 <see cref="ShouldWriteMarker"/> 存在的全部理由。
    ///
    /// 【已知的误判窗口，接受它】
    /// 本次修复上线前装过 26.10.11 / 26.11.0 的"新机器"，键已存在又没有标记，
    /// 会被判成老用户而补勾一次（用户可手动取消勾选，取消后不会再变——标记已落）。
    /// 这个窗口无法消除：那两版程序不写标记，历史无从追溯。
    /// </summary>
    public static class AutoStartTaskUpgradePolicy
    {
        /// <summary>一次性迁移标记的注册表值名（写成 1 表示已处理）</summary>
        public const string MarkerValueName = "AutoStartTaskMigrated";

        /// <summary>
        /// 纯判定：迁移标记是否已存在 + 设置键是否存在 → 这次该怎么办。
        ///
        /// 做成纯函数（不碰注册表）是为了能单元测试——注册表读写在本项目里
        /// 只由 SettingsStore 负责，测试宿主上不该动真实用户数据。
        /// </summary>
        public static AutoStartTaskUpgradeVerdict Judge(bool markerPresent, bool settingsKeyExists)
        {
            if (markerPresent) return AutoStartTaskUpgradeVerdict.AlreadyApplied;

            return settingsKeyExists
                ? AutoStartTaskUpgradeVerdict.EnableForUpgrade
                : AutoStartTaskUpgradeVerdict.KeepDefaultForNewInstall;
        }

        /// <summary>这次判定是否要把"启动时自动执行"写为 true（只有老用户升级需要）</summary>
        public static bool ShouldEnableAutoStartTask(AutoStartTaskUpgradeVerdict verdict)
        {
            return verdict == AutoStartTaskUpgradeVerdict.EnableForUpgrade;
        }

        /// <summary>
        /// 这次判定是否需要落"已处理"标记。
        /// **除"已处理过"之外一律要落**——见类注释里的"第二次启动"陷阱。
        /// </summary>
        public static bool ShouldWriteMarker(AutoStartTaskUpgradeVerdict verdict)
        {
            return verdict != AutoStartTaskUpgradeVerdict.AlreadyApplied;
        }
    }
}
