using System.Windows.Forms;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 窗体关闭策略：按"是谁要求关闭"决定放行还是拦截。
    ///
    /// 【为什么单独抽成一个类型，而不是直接写在 TimeSyncForm 里】
    /// 这段判断是实机缺陷「本程序正在阻止 Windows 关闭」的修复点，
    /// 必须能被回归测试盯住。而 TimeSyncForm 依赖太多（同步流程、托盘、
    /// 各类 Service），没法整个链进测试工程。
    ///
    /// 抽成独立类型后，探针（tools/ShutdownSessionProbe，net48）可以链接
    /// **同一份源码**，用真实的 WM_QUERYENDSESSION / SC_CLOSE 消息
    /// 去驱动真实决策逻辑——而不是在测试里复刻一份"看起来一样"的判断。
    ///
    /// 【复刻是这里最危险的坑】
    /// 一旦产品改了、测试里的副本没改，测试照样全绿，却什么都没验证到。
    /// 本项目已经吃过这类亏（见代码审查报告第六批「鉴别力验证」：
    /// 一条重放产品逻辑的测试，注入 bug 后依然全绿）。
    /// 所以这里的选择是：共享判断，而不是共享注释。
    /// </summary>
    public static class SessionClosePolicy
    {
        /// <summary>
        /// 该关闭原因是否属于"系统正在结束会话"（关机 / 注销）。
        ///
        /// 命中时必须**放行**关闭，并走正常的退出清理，绝不能拦。
        ///
        /// 【拦了会怎样——实机现象】
        /// 探针实测（tools/ShutdownSessionProbe 探针 1）：
        /// FormClosing 里设了 e.Cancel = true，.NET 对 WM_QUERYENDSESSION
        /// 就返回 0（FALSE）= 否决系统关机。Windows 随即弹出
        /// 「还需要关闭 1 个程序 / 此程序正在阻止 Windows 关闭」，
        /// 要求用户手动点"强制关机"。教室一体机上没人会去点那个按钮，
        /// 机器就一直卡在关机界面上。
        ///
        /// 【为什么只放行 WindowsShutDown，不放行 TaskManagerClosing】
        /// 见 TimeSyncForm.MainForm_FormClosing 的注释：任务管理器"结束任务"
        /// 更可能来自想绕过关机管理的学生，那种情况下保持常驻才是本意。
        /// 这条策略是有意的，探针 6 专门把它钉住。
        /// </summary>
        public static bool IsSystemSessionEnd(CloseReason reason)
        {
            return reason == CloseReason.WindowsShutDown;
        }
    }
}
