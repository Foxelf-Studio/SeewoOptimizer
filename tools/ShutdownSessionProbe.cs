using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SeewoOpt.Services;

namespace SeewoOpt.Tools
{
    /// <summary>
    /// 会话结束探针 —— 关机/注销链路的回归测试。
    ///
    /// 【为什么需要它，为什么不放在 NtpClientTests 里】
    /// 实机缺陷：弹着"即将关机"提醒框、用户没操作时到点执行关机，
    /// Windows 弹出「此程序正在阻止 Windows 关闭」，要用户手动点强制关机。
    ///
    /// 根因在**窗口消息层**：TimeSyncForm.MainForm_FormClosing 里一句
    /// 无条件的 `e.Cancel = true`（那是给"用户点 × 缩到托盘"用的），
    /// 让 .NET 对 WM_QUERYENDSESSION 返回 FALSE，等于否决系统关机。
    ///
    /// 这种缺陷 NtpClientTests（net10、无窗体、无消息循环）根本测不到：
    /// 它不存在于任何纯函数里，只存在于"真实消息被派发到真实窗口过程"
    /// 这条路径上。所以必须用 net48 + 真实窗体 + 真实 SendMessage 取证。
    ///
    /// 【探针怎么模拟系统关机】
    /// 直接 SendMessage 投递 WM_QUERYENDSESSION / WM_ENDSESSION / WM_CLOSE，
    /// 读回窗口过程的返回值。这是在不真的关机的前提下，能拿到的最接近
    /// 真实语义的手段（禁止、也不需要在沙箱里真的重启机器）。
    ///
    /// 【断言的含义】
    /// 探针 1、2 是对 .NET 既有语义的**基线取证**——它们记录了
    /// "为什么必须修"。若哪天变红，说明框架行为变了，需要重新审视整个修复。
    /// 探针 3、4、5 断言的是**本程序自己的契约**，是真正的回归防线。
    ///
    /// 结果同时写入 out/probe-result.txt（控制台编码在中文环境下不可靠）。
    /// 退出码：0 = 全部通过，1 = 有失败项。
    /// </summary>
    internal static class ShutdownSessionProbe
    {
        private const int WM_QUERYENDSESSION = 0x0011;
        private const int WM_ENDSESSION = 0x0016;
        private const int WM_CLOSE = 0x0010;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_CLOSE = 0xF060;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private static readonly StringBuilder Out = new StringBuilder();
        private static int _failed;

        private static void L(string s)
        {
            Out.AppendLine(s);
            Console.WriteLine(s);
        }

        private static void Check(string name, bool ok, string detail)
        {
            L(ok ? $"  [通过] {name}" : $"  [失败] {name}");
            if (!ok && !string.IsNullOrEmpty(detail))
                L($"          → {detail}");
            if (!ok) _failed++;
        }

        [STAThread]
        private static void Main()
        {
            L("=== 会话结束探针（关机/注销链路）===");
            L("");

            try
            {
                L("--- 探针 1：【基线】FormClosing 里 e.Cancel = true（复刻修复前的主窗体）---");
                ProbeBaselineCancel();
                L("");

                L("--- 探针 2：【基线】FormClosing 放行 ---");
                ProbeBaselineAllow();
                L("");

                L("--- 探针 3：真实 ShutdownWarningDialog 模态弹出时收到 WM_QUERYENDSESSION ---");
                ProbeDialog();
                L("");

                L("--- 探针 4：修复后的主窗体策略 · 系统关机应放行 ---");
                ProbeFixedOnShutdown();
                L("");

                L("--- 探针 5：修复后的主窗体策略 · 用户点 × 仍应缩到托盘 ---");
                ProbeFixedOnUserClose();
                L("");

                L("--- 探针 6：外部 WM_CLOSE（任务管理器「结束任务」）仍应被拦住 ---");
                ProbeFixedOnExternalClose();
            }
            catch (Exception ex)
            {
                L("探针异常：" + ex);
                _failed++;
            }

            L("");
            L(_failed == 0
                ? "=== 全部通过 ==="
                : $"=== 有 {_failed} 项失败 ===");

            try
            {
                string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "out");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "probe-result.txt"), Out.ToString(), Encoding.UTF8);
            }
            catch { }

            Environment.Exit(_failed == 0 ? 0 : 1);
        }

        /// <summary>
        /// 基线：FormClosing 中设 e.Cancel = true 时，.NET 对
        /// WM_QUERYENDSESSION 返回什么。
        ///
        /// 期望 0 —— 这一条就是"本程序阻止 Windows 关闭"的机制本身。
        /// 修复后主窗体不再走这条分支（见探针 4），但 .NET 的这条语义
        /// 仍然成立：谁在 FormClosing 里拦，谁就是在否决系统关机。
        /// </summary>
        private static void ProbeBaselineCancel()
        {
            var f = new Form { Text = "probe-veto", ShowInTaskbar = false };
            f.FormClosing += (s, e) =>
            {
                L($"         [FormClosing] 触发，CloseReason={e.CloseReason}");
                e.Cancel = true;   // 复刻修复前 MainForm_FormClosing 的行为
            };
            f.Show();
            Application.DoEvents();

            IntPtr r = SendMessage(f.Handle, WM_QUERYENDSESSION, IntPtr.Zero, IntPtr.Zero);
            L($"         -> 返回 {r}（{(r.ToInt64() == 1 ? "TRUE" : r.ToInt64() == 0 ? "FALSE" : "非标准")}）");
            Check("e.Cancel = true 会让 .NET 返回 FALSE（否决系统关机）",
                r.ToInt64() == 0,
                $"得到 {r.ToInt64()}——若这里不再是 0，说明 .NET 语义已变，整个修复的前提需要重新审视");

            try { f.Dispose(); } catch { }
        }

        /// <summary>基线：不取消时返回 TRUE（同意关机）</summary>
        private static void ProbeBaselineAllow()
        {
            var f = new Form { Text = "probe-allow", ShowInTaskbar = false };
            f.FormClosing += (s, e) =>
            {
                L($"         [FormClosing] 触发，CloseReason={e.CloseReason}");
                // 不设 e.Cancel
            };
            f.Show();
            Application.DoEvents();

            IntPtr r = SendMessage(f.Handle, WM_QUERYENDSESSION, IntPtr.Zero, IntPtr.Zero);
            L($"         -> 返回 {r}");
            Check("放行时 .NET 返回 TRUE（同意系统关机）",
                r.ToInt64() == 1,
                $"得到 {r.ToInt64()}");

            try { f.Dispose(); } catch { }
        }

        /// <summary>
        /// 真实对话框：系统关机时它应自己让开，且明确标记"不是用户选的"。
        ///
        /// 计时器设在父窗体上，靠 WM_TIMER 驱动——模态对话框的嵌套消息循环
        /// 同样会派发 WM_TIMER，所以 tick 能在 ShowDialog 阻塞期间执行。
        /// </summary>
        private static void ProbeDialog()
        {
            var parent = new Form { Text = "probe-parent", ShowInTaskbar = false };
            parent.Show();
            Application.DoEvents();

            ShutdownWarningDialog dlg = null;
            IntPtr queryResult = new IntPtr(-1);
            CloseReason closingReason = CloseReason.None;
            bool sawSessionEndFlag = false;
            bool seenInClosedEvent = false;
            bool returned = false;
            DialogResult seen = DialogResult.None;
            bool closedBySessionEndAfter = false;

            var t = new Timer { Interval = 300 };
            t.Tick += (s, e) =>
            {
                t.Stop();

                queryResult = SendMessage(dlg.Handle, WM_QUERYENDSESSION, IntPtr.Zero, IntPtr.Zero);
                L($"         [发送] WM_QUERYENDSESSION -> 返回 {queryResult}");
                L($"         [即时] IsDisposed={dlg.IsDisposed} Visible={dlg.Visible} DialogResult={dlg.DialogResult}");

                var t2 = new Timer { Interval = 400 };
                t2.Tick += (s3, e3) =>
                {
                    t2.Stop();
                    sawSessionEndFlag = SafeFlag(dlg);
                    L($"         [400ms 后] IsDisposed={dlg.IsDisposed} Visible={dlg.Visible} "
                    + $"DialogResult={dlg.DialogResult} ClosedBySessionEnd={sawSessionEndFlag}");

                    if (!dlg.IsDisposed)
                    {
                        L("         （对话框未自行关闭，探针兜底关闭以免挂死）");
                        try { dlg.DialogResult = DialogResult.Abort; dlg.Close(); } catch { }
                    }
                };
                t2.Start();
            };

            dlg = new ShutdownWarningDialog(5, DateTime.Now.AddMinutes(5));
            dlg.FormClosing += (s2, e2) =>
            {
                closingReason = e2.CloseReason;
            };
            dlg.FormClosed += (s2, e2) =>
            {
                // 在关闭那一刻读，避免之后被 Dispose 影响
                try { seenInClosedEvent = dlg.ClosedBySessionEnd; } catch { }
            };

            t.Start();
            seen = dlg.ShowDialog(parent);
            returned = true;
            closedBySessionEndAfter = SafeFlag(dlg);

            L($"         [ShowDialog 返回] DialogResult={seen}，"
            + $"关闭原因={closingReason}，ClosedBySessionEnd={closedBySessionEndAfter}");

            Check("对话框同意系统关机（返回 TRUE，不做阻止关机的那一个）",
                queryResult.ToInt64() == 1,
                $"得到 {queryResult.ToInt64()}");
            Check("对话框在系统关机时自行关闭（ShowDialog 返回，不留窗口让系统等）",
                returned && closingReason == CloseReason.WindowsShutDown,
                $"returned={returned}, CloseReason={closingReason}");
            Check("CloseReason 被识别为 WindowsShutDown（而非用户关窗）",
                closingReason == CloseReason.WindowsShutDown,
                $"得到 {closingReason}");
            Check("ClosedBySessionEnd = true（调用方据此避免误判为\"本次不关机\"）",
                closedBySessionEndAfter || seenInClosedEvent,
                $"ShowDialog 后={closedBySessionEndAfter}, FormClosed 时={seenInClosedEvent}");

            try { parent.Dispose(); } catch { }
        }

        /// <summary>
        /// 修复后的主窗体策略：系统关机必须放行（返回 TRUE，不缩托盘）。
        /// 逻辑与 TimeSyncForm.MainForm_FormClosing 的分流保持一致。
        /// </summary>
        private static void ProbeFixedOnShutdown()
        {
            var f = new Form { Text = "probe-fixed-shutdown", ShowInTaskbar = false };
            bool exiting = false;
            bool minimizedToTray = false;

            f.FormClosing += (s, e) =>
            {
                // 复刻修复后的分流：系统级关闭 → 放行
                // 关键：调用**产品同一份**策略源码（见 csproj 链接说明），
                // 而不是在这里复刻一份判断——复刻的副本跑绿说明不了任何事。
                if (SessionClosePolicy.IsSystemSessionEnd(e.CloseReason))
                    exiting = true;

                if (exiting) return;

                e.Cancel = true;
                minimizedToTray = true;
                f.Hide();
            };

            f.Show();
            Application.DoEvents();

            IntPtr r = SendMessage(f.Handle, WM_QUERYENDSESSION, IntPtr.Zero, IntPtr.Zero);
            L($"         -> 返回 {r}，是否误走缩托盘={minimizedToTray}");

            Check("修复后：系统关机请求被同意（返回 TRUE）",
                r.ToInt64() == 1,
                $"得到 {r.ToInt64()}");
            Check("修复后：系统关机不再被误判为\"点 × 缩托盘\"",
                !minimizedToTray,
                "被误判会让 Windows 认为本程序在阻止关机");

            try { f.Dispose(); } catch { }
        }

        /// <summary>
        /// 修复后的主窗体策略：用户点 × 仍必须缩到托盘（不能把常驻功能修没了）。
        ///
        /// 【为什么用 WM_SYSCOMMAND + SC_CLOSE，而不是直接发 WM_CLOSE】
        /// 这是探针开发中实测出来的差异，值得记一笔：
        /// 外部直接 SendMessage(WM_CLOSE) 时，.NET 判定的 CloseReason 是
        /// **TaskManagerClosing**；而用户点标题栏 × 时，窗口过程收到的是
        /// **WM_SYSCOMMAND(SC_CLOSE)**（非客户区点击由系统翻译成这条），
        /// 那条路径才对应 CloseReason.UserClosing。
        ///
        /// 最初版本用 WM_CLOSE 模拟点 ×，结果拿到 TaskManagerClosing，
        /// 断言失败——那不是产品的问题，是模拟方式不对。
        /// 换成 SC_CLOSE 才与真实点击一致。
        /// </summary>
        private static void ProbeFixedOnUserClose()
        {
            var f = new Form { Text = "probe-fixed-userclose", ShowInTaskbar = false };
            bool exiting = false;
            bool minimizedToTray = false;
            CloseReason seenReason = CloseReason.None;

            f.FormClosing += (s, e) =>
            {
                seenReason = e.CloseReason;
                // 关键：调用**产品同一份**策略源码（见 csproj 链接说明），
                // 而不是在这里复刻一份判断——复刻的副本跑绿说明不了任何事。
                if (SessionClosePolicy.IsSystemSessionEnd(e.CloseReason))
                    exiting = true;

                if (exiting) return;

                e.Cancel = true;      // 用户关窗 → 取消关闭、缩到托盘
                minimizedToTray = true;
                f.Hide();
            };

            f.Show();
            Application.DoEvents();

            SendMessage(f.Handle, WM_SYSCOMMAND, new IntPtr(SC_CLOSE), IntPtr.Zero);

            bool alive = !f.IsDisposed;
            bool hidden = !f.Visible;
            L($"         -> CloseReason={seenReason}，仍存活={alive}，已隐藏={hidden}");

            Check("修复后：用户点 × 仍走缩到托盘（CloseReason = UserClosing）",
                seenReason == CloseReason.UserClosing && minimizedToTray && alive && hidden,
                $"CloseReason={seenReason}, 缩托盘={minimizedToTray}, 存活={alive}, 隐藏={hidden}");
            Check("修复后：只有 WindowsShutDown 被放行，用户关窗不受影响",
                seenReason != CloseReason.WindowsShutDown,
                $"得到 {seenReason}");

            try { f.Dispose(); } catch { }
        }

        /// <summary>
        /// 修复后的主窗体策略：外部 WM_CLOSE 仍应被拦住。
        ///
        /// .NET 把外部 WM_CLOSE 判为 TaskManagerClosing，而本程序**有意**
        /// 仍然拦住它——教室里更常见的是学生想绕过关机管理去结束它，
        /// 那种情况下"缩到托盘、继续守着"才是该做的事。
        /// 本批只修实机报的关机缺陷，不顺手改动这条策略，这里把它钉住。
        /// </summary>
        private static void ProbeFixedOnExternalClose()
        {
            var f = new Form { Text = "probe-fixed-externalclose", ShowInTaskbar = false };
            bool exiting = false;
            bool minimizedToTray = false;
            CloseReason seenReason = CloseReason.None;

            f.FormClosing += (s, e) =>
            {
                seenReason = e.CloseReason;
                // 关键：调用**产品同一份**策略源码（见 csproj 链接说明），
                // 而不是在这里复刻一份判断——复刻的副本跑绿说明不了任何事。
                if (SessionClosePolicy.IsSystemSessionEnd(e.CloseReason))
                    exiting = true;

                if (exiting) return;

                e.Cancel = true;
                minimizedToTray = true;
                f.Hide();
            };

            f.Show();
            Application.DoEvents();

            SendMessage(f.Handle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);

            bool alive = !f.IsDisposed;
            bool hidden = !f.Visible;
            L($"         -> CloseReason={seenReason}，仍存活={alive}，已隐藏={hidden}");

            Check("外部结束任务（TaskManagerClosing）仍被拦住，程序继续常驻",
                seenReason == CloseReason.TaskManagerClosing && minimizedToTray && alive && hidden,
                $"CloseReason={seenReason}, 缩托盘={minimizedToTray}, 存活={alive}, 隐藏={hidden}");

            try { f.Dispose(); } catch { }
        }

        private static bool SafeFlag(ShutdownWarningDialog dlg)
        {
            try { return dlg.ClosedBySessionEnd; }
            catch { return false; }
        }
    }
}
