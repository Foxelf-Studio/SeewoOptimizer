using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using SeewoOpt.Services;

namespace SeewoOpt.Tools
{
    /// <summary>
    /// 更新脚本探针 —— 自动更新链路的回归测试。
    ///
    /// 【要盯住的是什么】
    /// 自动更新流程：下载 exe → 下载 .sig → 验签 → 比对哈希 →
    /// **生成并启动替换脚本** → 主程序退出、脚本替换文件。
    /// 2026-10-11 实机失败发生在倒数第二步：前四步全部通过，启动脚本时抛
    /// `The Process object must have the UseShellExecute property set to false
    /// in order to use environment variables.`
    ///
    /// 顺着这条线又挖出两个更深的问题：
    ///   · `EnvironmentVariables` 在环境变量含"仅大小写不同"的重复键时也会抛异常
    ///     （与 UseShellExecute 无关）——机制本身把成败押在用户机器的环境上；
    ///   · 原脚本里的 `taskkill /f /im "%~nx1"` 从来没拿到过参数，
    ///     兜底防线形同虚设。
    ///
    /// 探针做两件事：
    ///   1. 断言启动参数的决定性契约（确定性判断，不看环境）；
    ///   2. **用产品真实模板**在临时目录里跑一次完整替换演练——
    ///      源文件、目标文件、备份、pending 全是真文件，脚本也是真脚本，
    ///      最后逐项核对替换结果。
    ///
    /// 【为什么必须链接真实模板，不能自己写一份】
    /// 改模板时漏掉四个闭合引号、探针却全绿——因为它当时测的是自己另写的
    /// 批处理。这正是本项目反复撞到的"复刻副本 = 假测试"。
    ///
    /// 【安全约束】演练用的目标文件名故意不叫 SeewoOpt.exe，
    /// 免得脚本里的 taskkill 误杀真正在运行的程序。
    ///
    /// 退出码：0 = 全部通过，1 = 有失败项。
    /// </summary>
    internal static class UpdateLaunchProbe
    {
        private static readonly StringBuilder Out = new StringBuilder();
        private static int _failed;

        private static void L(string s)
        {
            Out.AppendLine(s);
            Console.WriteLine(s);
        }

        private static void Check(string name, bool ok, string detail = null)
        {
            L(ok ? $"  [通过] {name}" : $"  [失败] {name}");
            if (!ok && !string.IsNullOrEmpty(detail))
                L($"          → {detail}");
            if (!ok) _failed++;
        }

        [STAThread]
        private static void Main()
        {
            L("=== 更新脚本探针 ===");
            L("");

            try
            {
                L("--- 探针 1：启动参数的决定性契约（不依赖任何环境）---");
                ProbeStartInfoContract();
                L("");

                L("--- 探针 2：用真实模板跑一次完整替换演练 ---");
                ProbeRealReplacement();
                L("");

                L("--- 探针 3：【留证】环境变量传值的机制有多脆（不计成败）---");
                ObserveEnvironmentVariablesHazard();
            }
            catch (Exception ex)
            {
                L("探针异常：" + ex);
                _failed++;
            }

            L("");
            L(_failed == 0 ? "=== 全部通过 ===" : $"=== 有 {_failed} 项失败 ===");

            try
            {
                string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "out");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "probe-result.txt"), Out.ToString(), Encoding.UTF8);
            }
            catch { }

            Environment.Exit(_failed == 0 ? 0 : 1);
        }

        /// <summary>断言产品构造出的启动参数。全部是确定性判断。</summary>
        private static void ProbeStartInfoContract()
        {
            string plain = Path.Combine(Path.GetTempPath(), "update.bat");
            ProcessStartInfo psi = UpdateScriptLauncher.BuildStartInfo(plain);

            L($"         FileName  = {psi.FileName}");
            L($"         Arguments = {psi.Arguments}");
            L($"         UseShellExecute = {psi.UseShellExecute}");

            Check("cmd.exe 用绝对路径（不依赖 PATH / 当前目录搜索）",
                Path.IsPathRooted(psi.FileName)
                    && string.Equals(Path.GetFileName(psi.FileName), "cmd.exe",
                                     StringComparison.OrdinalIgnoreCase),
                $"得到 {psi.FileName}");

            Check("UseShellExecute = false（.bat 不是可执行映像，必须交给 cmd.exe）",
                !psi.UseShellExecute,
                "走 ShellExecute 等于把命令行交给外壳解释：引号规则不同、"
              + "可能闪现控制台窗口；而最初的实机故障正是"
              + "UseShellExecute = true 配上环境变量传值，直接抛异常");

            Check("参数把脚本路径包在引号里",
                psi.Arguments == "/c \"" + plain + "\"",
                $"得到 {psi.Arguments}");

            string spaced = Path.Combine(Path.GetTempPath(), "some folder", "update.bat");
            Check("脚本路径含空格时引号仍然正确",
                UpdateScriptLauncher.BuildStartInfo(spaced).Arguments == "/c \"" + spaced + "\"",
                "引号规则写错会导致脚本根本起不来");

            Check("空路径被拒绝",
                Throws(() => UpdateScriptLauncher.BuildStartInfo(null))
                    && Throws(() => UpdateScriptLauncher.BuildStartInfo("")));
        }

        /// <summary>
        /// 完整替换演练：造真文件 → 生成真实脚本 → 真的执行 → 逐项核对结果。
        ///
        /// 这一步验证的是"更新真的能把文件换掉"，而不只是"能启动进程"：
        ///   · 路径含 % 与空格时，转义与引号是否都正确；
        ///   · 替换结果、源文件清理、备份清理、pending 清理是否都发生。
        /// </summary>
        private static void ProbeRealReplacement()
        {
            string root = Path.Combine(Path.GetTempPath(),
                                       "seewo upd probe " + Guid.NewGuid().ToString("N"));
            // 目录名故意含 % 和空格——这两个字符最容易破坏批处理语法
            string dir = Path.Combine(root, "100% done");
            Directory.CreateDirectory(dir);

            // 【安全】目标文件名故意不叫 SeewoOpt.exe：
            // 脚本里有 taskkill /f /im %TARGET_NAME%，用真名会误杀正在运行的程序
            string target = Path.Combine(dir, "SeewoProbe-Target.exe");
            string source = Path.Combine(root, "SeewoProbe-Source.exe");
            string backup = target + ".bak";
            string pending = Path.Combine(root, "pending.txt");
            string script = Path.Combine(root, "update.bat");

            File.WriteAllText(target, "OLD");
            byte[] newBytes = new byte[4096];
            for (int i = 0; i < newBytes.Length; i++) newBytes[i] = (byte)(i % 251);
            File.WriteAllBytes(source, newBytes);
            File.WriteAllText(pending, source);

            string content = UpdateScriptLauncher.BuildScriptContent(
                target, source, backup, pending);
            File.WriteAllText(script, content, Encoding.ASCII);

            // 先静态核对生成出来的脚本：set 行的引号必须成对（我在这里栽过一次）
            string escTarget = UpdateScriptLauncher.EscapeForBatch(target);
            Check("生成的脚本里 set 行引号成对（含 % 的路径已转义为 %%）",
                content.Contains("set \"TARGET=" + escTarget + "\""),
                "set 行少了闭合引号的话，TARGET 会带上多余字符，替换必然失败");

            L($"         演练目录：{dir}");

            using (Process p = Process.Start(UpdateScriptLauncher.BuildStartInfo(script)))
            {
                if (!p.WaitForExit(30000))
                {
                    Check("替换脚本在 30 秒内结束", false, "超时，脚本可能卡住了");
                    return;
                }
                L($"         脚本退出码：{p.ExitCode}");
                Check("替换成功（脚本退出码 0）", p.ExitCode == 0,
                    $"退出码 {p.ExitCode}——非 0 表示走到了 :copyfailed 分支");
            }

            bool replaced = File.Exists(target)
                && new FileInfo(target).Length == newBytes.Length;
            Check("目标文件已被替换成新版本（字节数一致）", replaced,
                replaced ? "" : $"目标文件大小 {SafeLen(target)}，期望 {newBytes.Length}");

            Check("源文件已清理", !File.Exists(source));
            Check("备份文件已清理", !File.Exists(backup));
            Check("pending 标记已清理", !File.Exists(pending));
            // 脚本**不**自删：自删会让 cmd 读不到后续命令，
            // 连 exit /b 0 都执行不到，退出码会变成 1（探针实测）。
            // 因此这里反过来断言"脚本仍在"，把那个决定钉住。
            Check("脚本保留（不自删——自删会破坏退出码）", File.Exists(script));

            try { Directory.Delete(root, true); } catch { }
        }

        private static long SafeLen(string path)
        {
            try { return new FileInfo(path).Length; } catch { return -1; }
        }

        /// <summary>
        /// 【留证，不判成败】环境变量传值这条路有多脆。
        /// 结果取决于运行环境，因此只打印观察值。
        /// </summary>
        private static void ObserveEnvironmentVariablesHazard()
        {
            L($"         UseShellExecute = false 时访问 EnvironmentVariables：{TryEnv(false)}");
            L($"         UseShellExecute = true  时访问 EnvironmentVariables：{TryEnv(true)}");
            L("         （产品已不使用该机制：它既要求 UseShellExecute = false，");
            L("           又会被用户机器上重复的环境变量搞崩。）");
        }

        private static string TryEnv(bool shellExecute)
        {
            try
            {
                var psi = new ProcessStartInfo { FileName = "x.bat", UseShellExecute = shellExecute };
                psi.EnvironmentVariables["PROBE_KEY"] = "1";
                return "可访问（未抛异常）";
            }
            catch (Exception ex)
            {
                return "抛异常：" + ex.Message;
            }
        }

        private static bool Throws(Action action)
        {
            try { action(); return false; }
            catch (ArgumentNullException) { return true; }
        }
    }
}
