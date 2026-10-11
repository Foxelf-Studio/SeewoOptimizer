using System;
using System.Diagnostics;
using System.IO;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 构造"启动更新替换脚本"所需的进程参数。
    ///
    /// 【为什么单独抽出来——这个缺陷差点让自动更新永久失效】
    /// 2026-10-11 实机：自动更新在最后一步崩了，日志写着
    /// `The Process object must have the UseShellExecute property set to false
    /// in order to use environment variables.`
    /// 下载、验签、哈希比对**全都通过了**，只死在"启动脚本"这一句上。
    ///
    /// 两个约束必须同时满足：
    ///   1. `EnvironmentVariables` 只在 `UseShellExecute = false` 时可用；
    ///      设 true 会直接抛异常。而这里必须用环境变量传路径——
    ///      路径里可能有空格，拼进命令行会破坏批处理语法。
    ///   2. `.bat` 不是可执行映像，`UseShellExecute = false` 时**不能直接启动**，
    ///      必须交给 `cmd.exe /c`。
    ///
    /// 抽成独立类型是为了让回归测试链接**同一份源码**去真的启动一次进程。
    /// 这类缺陷纯逻辑测试永远测不到——它只在"真的 CreateProcess"时暴露，
    /// 而在测试里复刻一份构造代码等于什么都没验证（本项目已多次确认）。
    /// </summary>
    public static class UpdateScriptLauncher
    {
        /// <summary>
        /// 解析 cmd.exe 的绝对路径。
        ///
        /// 【为什么不用裸文件名 "cmd.exe"】与 shutdown.exe 同理（第十批）：
        /// 裸名走 CreateProcess 的搜索顺序（**先当前目录**、再 PATH）。
        /// 本程序常被放在桌面或任意文件夹里，同目录下一个同名文件就会被先命中，
        /// 结果是启动了别的东西（甚至是别人放的 cmd.exe）。
        /// </summary>
        public static string ResolveCmdExe()
        {
            string system32 = Environment.GetFolderPath(Environment.SpecialFolder.System);
            string primary = Path.Combine(system32, "cmd.exe");
            if (File.Exists(primary)) return primary;

            // 32 位进程跑在 64 位系统上时 System32 会被重定向到 SysWOW64（那里没有它），
            // 正解是 Sysnative。本程序目前是 AnyCPU 走不到这里，留作保险。
            string sysnative = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Sysnative", "cmd.exe");
            if (File.Exists(sysnative)) return sysnative;

            return primary;
        }

        /// <summary>
        /// 把值安全地写进批处理脚本。
        ///
        /// 【为什么不用环境变量传值——探针实测两个坑，都会让自动更新彻底失效】
        /// 1. `EnvironmentVariables` 只在 `UseShellExecute = false` 时可用；
        ///    设 true 会抛 "The Process object must have the UseShellExecute
        ///    property set to false in order to use environment variables."
        ///    ——2026-10-11 实机更新失败就是这一条。
        /// 2. 更隐蔽：该属性的 getter 会把当前进程的环境变量整块搬进一个
        ///    **大小写不敏感**的字典，只要环境里存在"仅大小写不同"的两个键
        ///    （`HTTP_PROXY` 与 `http_proxy`，MSYS/Git 环境里很常见），
        ///    就会抛 "已添加项。字典中的关键字…"。这一条与 UseShellExecute
        ///    毫无关系——**用这个机制就等于把成败押在用户机器的环境变量上**。
        ///
        /// 所以改为写进脚本本身。只需要处理 `%`：
        ///   · `%` 必须写成 `%%`（批处理里 `%%` 折叠回 `%`）；
        ///   · `"` 在 Windows 路径里非法，不可能出现；
        ///   · `&` `^` `(` `)` 在 `set "VAR=..."` 的引号内是安全的——
        ///     引号正是为此而加。
        /// </summary>
        public static string EscapeForBatch(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Replace("%", "%%");
        }

        /// <summary>
        /// 生成更新替换脚本的完整内容（批处理文本）。
        ///
        /// 【为什么模板要放在这里，而不是留在 Program.CreateUpdateScript 里】
        /// 之前探针只测了"启动参数"，批处理模板本身没被测到——
        /// 结果我在改写模板时漏掉了四个闭合引号（`set "TARGET=值` 少了结尾的 `"`），
        /// 探针照样全绿，因为它测的是**自己另写的一份**批处理。
        /// 这正是本项目反复撞到的"复刻副本 = 假测试"。
        /// 把模板搬到这里之后，探针链接同一份源码，写出的就是产品真正要执行的脚本。
        ///
        /// 全部走 ASCII：批处理默认按 GBK 解析，脚本里出现中文会解析错乱。
        /// </summary>
        public static string BuildScriptContent(string targetExe, string newExe,
                                                string backupExe, string pendingFile)
        {
            string template = @"@echo off
setlocal enabledelayedexpansion
title Updating

set ""TARGET=@@TARGET@@""
set ""SOURCE=@@SOURCE@@""
set ""BACKUP=@@BACKUP@@""
set ""PENDING=@@PENDING@@""
set ""TARGET_NAME=@@TARGET_NAME@@""

:: Give the main process time to fully exit
ping -n 3 127.0.0.1 > nul

:: Kill any lingering instance
::
:: [WHY TARGET_NAME AND NOT %~nx1]
:: The original script used %~nx1 (name+ext of the first command line
:: argument), but no argument was ever passed, so it was always empty --
:: the taskkill silently failed behind >nul 2>&1 and the 「process still
:: holds the file」 safety net did not actually exist.
taskkill /f /im ""%TARGET_NAME%"" > nul 2>&1
ping -n 2 127.0.0.1 > nul

:: Back up current executable
if exist ""%TARGET%"" copy /y ""%TARGET%"" ""%BACKUP%"" > nul

:: Replace with retries (bounded, never loop forever)
set /a ATTEMPT=0
:copyloop
set /a ATTEMPT+=1
if !ATTEMPT! GTR 5 goto :copyfailed

del /f /q ""%TARGET%"" > nul 2>&1
copy /y ""%SOURCE%"" ""%TARGET%"" > nul 2>&1

if exist ""%TARGET%"" goto :success
ping -n 2 127.0.0.1 > nul
goto :copyloop

:copyfailed
:: Restore backup so the program is not left missing
if exist ""%BACKUP%"" (
    copy /y ""%BACKUP%"" ""%TARGET%"" > nul 2>&1
    del /f /q ""%BACKUP%"" > nul 2>&1
)
echo UPDATE_FAILED > ""%PENDING%""
del /f /q ""%SOURCE%"" > nul 2>&1
exit /b 1

:success
:: Verify the new file is a real PE binary
set /a SIZE=0
for %%F in (""%TARGET%"") do set /a SIZE=%%~zF
if !SIZE! LSS 1024 goto :copyfailed

del /f /q ""%SOURCE%"" > nul 2>&1
del /f /q ""%SOURCE%.sig"" > nul 2>&1
del /f /q ""%BACKUP%"" > nul 2>&1
del /f /q ""%PENDING%"" > nul 2>&1
::
:: [WHY THIS SCRIPT DOES NOT DELETE ITSELF]
:: It used to end with  del /f /q ""%~f0""  followed by  exit /b 0.
:: That is a trap: once the running script file is gone, cmd can no longer
:: read the next command from it, so exit /b 0 never ran and the exit code
:: became 1 -- the script reported failure after a successful replacement.
:: (Confirmed by tools/UpdateLaunchProbe: target replaced, everything
::  cleaned up, yet exit code 1.)
:: Leaving the file costs nothing: it is overwritten on the next update
:: attempt, and the whole updates folder is removed on uninstall.
exit /b 0
";

            return template
                .Replace("@@TARGET@@", EscapeForBatch(targetExe))
                .Replace("@@SOURCE@@", EscapeForBatch(newExe))
                .Replace("@@BACKUP@@", EscapeForBatch(backupExe))
                .Replace("@@PENDING@@", EscapeForBatch(pendingFile))
                .Replace("@@TARGET_NAME@@", EscapeForBatch(Path.GetFileName(targetExe)));
        }

        /// <summary>
        /// 构造启动参数：用 cmd.exe 承载批处理。
        /// </summary>
        public static ProcessStartInfo BuildStartInfo(string scriptPath)
        {
            if (string.IsNullOrEmpty(scriptPath))
                throw new ArgumentNullException("scriptPath");

            return new ProcessStartInfo
            {
                FileName = ResolveCmdExe(),

                // 【引用规则】/c 后面用一对引号包住脚本路径：
                // cmd 的规则是"首尾都是引号时去掉最外层那对"，
                // 因此路径含空格时也能正确执行。
                Arguments = "/c \"" + scriptPath + "\"",

                UseShellExecute = false,     // 必须：.bat 不是可执行映像，只能交给 cmd
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
        }
    }
}
