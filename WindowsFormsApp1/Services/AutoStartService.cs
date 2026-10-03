using System;
using System.IO;
using Microsoft.Win32.TaskScheduler;

namespace SeewoOpt.Services
{
    /// <summary>自启同步结果</summary>
    public class AutoStartResult
    {
        /// <summary>操作是否成功</summary>
        public bool Success { get; set; }

        /// <summary>失败原因（Success 为 false 时有值）</summary>
        public string ErrorMessage { get; set; }

        /// <summary>需要向用户弹窗提示（权限问题等）</summary>
        public bool ShouldNotifyUser { get; set; }

        /// <summary>用于日志的过程描述</summary>
        public string Detail { get; set; }
    }

    /// <summary>
    /// 开机自启的计划任务管理。
    ///
    /// 任务名固定为 TimeSyncTool，触发方式为「用户登录后延迟 5 秒」，
    /// 动作为启动当前 exe。原先这些逻辑在 TimeSyncForm 里，
    /// 与窗体生命周期耦合，不利于单独验证。
    /// </summary>
    public static class AutoStartService
    {
        /// <summary>计划任务名称（沿用历史值，改动会导致老用户的自启失效）</summary>
        public const string TaskName = "TimeSyncTool";

        /// <summary>登录后延迟多少秒启动</summary>
        private static readonly TimeSpan TriggerDelay = TimeSpan.FromSeconds(5);

        private const string TaskDescription = "川中计算机协会 - 陈叔叔系统优化工具";
        private const string TaskAuthor = "TimeSyncTool";

        /// <summary>任务当前是否存在。查询失败时返回 false。</summary>
        public static bool IsTaskRegistered(out string error)
        {
            error = null;
            try
            {
                using (TaskService ts = new TaskService())
                {
                    return ts.GetTask(TaskName) != null;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 让计划任务的存在状态与期望的 enabled 保持一致。
        /// enabled 为 true 则创建（已存在则先删后建），false 则删除。
        /// </summary>
        public static AutoStartResult Sync(bool enabled, string executablePath)
        {
            var result = new AutoStartResult();
            var detail = new System.Text.StringBuilder();

            try
            {
                using (TaskService ts = new TaskService())
                {
                    if (enabled)
                    {
                        if (string.IsNullOrEmpty(executablePath) || !File.Exists(executablePath))
                        {
                            result.Success = false;
                            result.ErrorMessage = "程序文件不存在，无法创建自启任务";
                            result.Detail = detail.ToString();
                            return result;
                        }

                        if (ts.GetTask(TaskName) != null)
                        {
                            ts.RootFolder.DeleteTask(TaskName, false);
                            detail.AppendLine("已删除旧任务");
                        }

                        TaskDefinition td = ts.NewTask();
                        td.RegistrationInfo.Description = TaskDescription;
                        td.RegistrationInfo.Author = TaskAuthor;

                        td.Triggers.Add(new LogonTrigger
                        {
                            UserId = null,
                            Delay = TriggerDelay
                        });

                        td.Actions.Add(new ExecAction(executablePath, null, null));

                        ts.RootFolder.RegisterTaskDefinition(TaskName, td);
                        detail.AppendLine("自启任务创建成功");
                    }
                    else
                    {
                        if (ts.GetTask(TaskName) != null)
                        {
                            ts.RootFolder.DeleteTask(TaskName, false);
                            detail.AppendLine("自启任务已删除");
                        }
                        else
                        {
                            detail.AppendLine("自启任务本就不存在，无需删除");
                        }
                    }
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                // 权限类错误值得提示用户，普通失败只记日志
                result.ShouldNotifyUser = true;
                LogService.Write("自启任务操作异常: " + ex);
            }

            result.Detail = detail.ToString();
            return result;
        }

        /// <summary>
        /// 检查实际状态与设置是否一致，不一致时自动修正。
        /// 用户可能直接在任务计划程序里删过任务，这里做自愈。
        /// </summary>
        public static AutoStartResult EnsureConsistency(bool desiredEnabled, string executablePath)
        {
            string queryError;
            bool exists = IsTaskRegistered(out queryError);

            if (queryError != null)
            {
                LogService.Write("查询自启任务状态失败: " + queryError);
                return new AutoStartResult { Success = true, Detail = "查询失败，跳过一致性检查" };
            }

            if (exists == desiredEnabled)
            {
                LogService.Write($"自启状态一致（任务存在={exists}，设置={desiredEnabled}），无需调整");
                return new AutoStartResult { Success = true, Detail = "状态一致，无需调整" };
            }

            LogService.Write($"自启状态不一致（任务存在={exists}，设置={desiredEnabled}），正在修正");
            return Sync(desiredEnabled, executablePath);
        }

        /// <summary>删除自启任务（卸载流程使用）。失败不抛异常。</summary>
        public static bool Remove()
        {
            try
            {
                using (TaskService ts = new TaskService())
                {
                    if (ts.GetTask(TaskName) == null) return true;
                    ts.RootFolder.DeleteTask(TaskName, false);
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogService.Write("删除自启任务失败: " + ex.Message);
                return false;
            }
        }
    }
}
