using System;
using System.IO;
using System.Text;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 统一日志服务。
    ///
    /// 历史问题：Program.cs 与 TimeSyncForm.cs 各自实现了一份 WriteLog，
    /// 时间格式还不一致（一个带毫秒一个不带），导致日志交错且难以排查。
    /// 现统一到此处，格式固定为 "yyyy-MM-dd HH:mm:ss.fff"。
    ///
    /// 线程安全：多个线程（UI 线程、同步线程、更新线程）会并发写同一个文件，
    /// File.AppendAllText 在并发下可能抛 IOException（文件被占用），
    /// 因此用锁串行化写入。
    /// </summary>
    public static class LogService
    {
        /// <summary>日志根目录：%LOCALAPPDATA%\TimeSyncTool</summary>
        public static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TimeSyncTool");

        /// <summary>日志文件完整路径</summary>
        public static readonly string LogFilePath = Path.Combine(LogDirectory, "startup.log");

        private static readonly object _writeLock = new object();

        /// <summary>
        /// 日志写入开关。卸载流程会删除整个日志目录，必须先置为 false 停写，
        /// 否则后续写日志会因目录不存在而失败。
        /// </summary>
        public static volatile bool Enabled = true;

        /// <summary>单条日志最大长度，防止异常信息过长撑爆日志文件</summary>
        private const int MaxMessageLength = 4000;

        /// <summary>
        /// 写入一条日志。永不抛异常——日志失败不应影响主流程。
        /// </summary>
        public static void Write(string message)
        {
            if (!Enabled) return;

            try
            {
                string line = string.Format(
                    "{0:yyyy-MM-dd HH:mm:ss.fff} - {1}{2}",
                    DateTime.Now,
                    Truncate(message),
                    Environment.NewLine);

                lock (_writeLock)
                {
                    if (!Directory.Exists(LogDirectory))
                        Directory.CreateDirectory(LogDirectory);

                    File.AppendAllText(LogFilePath, line, Encoding.UTF8);
                }
            }
            catch
            {
                // 日志失败必须静默——不能因为写不了日志就把程序搞崩
            }
        }

        /// <summary>截断超长消息，避免单条异常堆栈占满磁盘</summary>
        private static string Truncate(string message)
        {
            if (string.IsNullOrEmpty(message)) return string.Empty;
            if (message.Length <= MaxMessageLength) return message;
            return message.Substring(0, MaxMessageLength) + " ...(已截断)";
        }
    }
}
