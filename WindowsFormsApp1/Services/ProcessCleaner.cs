using System;
using System.Diagnostics;

namespace SeewoOpt.Services
{
    /// <summary>单个进程的结束结果</summary>
    public struct ProcessKillResult
    {
        /// <summary>进程名</summary>
        public string ProcessName;

        /// <summary>进程 ID</summary>
        public int ProcessId;

        /// <summary>是否成功结束</summary>
        public bool Success;

        /// <summary>失败或超时原因</summary>
        public string Message;
    }

    /// <summary>
    /// 结束指定名称的进程（用于清理 WPS 相关进程）。
    ///
    /// 原实现直接在窗体里遍历并 Kill，此处抽出以便单独测试与复用。
    /// 逐个进程单独捕获异常——某个进程因权限不足杀不掉时，
    /// 不应影响后续进程的处理。
    /// </summary>
    public static class ProcessCleaner
    {
        /// <summary>WPS 及其组件的进程名</summary>
        private static readonly string[] WpsProcessNames =
        {
            "wps",        // 文字
            "wpp",        // 演示
            "et",         // 表格
            "wpscloudsvr",// 云服务
            "ksolaunch"   // 启动器
        };

        /// <summary>结束 WPS 相关进程。</summary>
        /// <param name="onProgress">
        /// 单个进程的处理回调，参数为该进程的结果。用于向界面输出日志。
        /// 可为 null。
        /// </param>
        /// <returns>成功结束的数量</returns>
        public static int KillWpsProcesses(Action<ProcessKillResult> onProgress = null)
        {
            int killedCount = 0;

            foreach (string processName in WpsProcessNames)
            {
                Process[] processes;
                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch (Exception ex)
                {
                    // 枚举进程本身失败（如权限不足），记录后继续处理下一组
                    Report(onProgress, new ProcessKillResult
                    {
                        ProcessName = processName,
                        ProcessId = 0,
                        Success = false,
                        Message = "枚举进程失败: " + ex.Message
                    });
                    continue;
                }

                foreach (Process process in processes)
                {
                    ProcessKillResult result = KillOne(process, processName);

                    if (result.Success) killedCount++;

                    Report(onProgress, result);
                }
            }

            return killedCount;
        }

        /// <summary>结束单个进程并等待其真正退出</summary>
        private static ProcessKillResult KillOne(Process process, string fallbackName)
        {
            var result = new ProcessKillResult
            {
                ProcessName = fallbackName,
                ProcessId = 0,
                Success = false,
                Message = null
            };

            try
            {
                result.ProcessId = process.Id;
                result.ProcessName = process.ProcessName;

                process.Kill();
                process.WaitForExit(3000);

                if (process.HasExited)
                {
                    result.Success = true;
                }
                else
                {
                    result.Message = "超时";
                }
            }
            catch (Exception ex)
            {
                result.Message = ex.Message;
            }
            finally
            {
                // Process 对象持有系统句柄，必须释放
                process.Dispose();
            }

            return result;
        }

        private static void Report(Action<ProcessKillResult> onProgress, ProcessKillResult result)
        {
            if (onProgress == null) return;

            try
            {
                onProgress(result);
            }
            catch
            {
                // 日志回调自身的异常不能影响清理流程
            }
        }
    }
}
