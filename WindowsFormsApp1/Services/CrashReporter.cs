using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 一次崩溃/失败上报的内容。
    ///
    /// 字段刻意保持"设备画像"级别：能判断问题与环境是否相关，
    /// 但不足以定位到具体的人。姓名、机器名、IP 在发送前已抹掉。
    /// </summary>
    [DataContract]
    public class CrashReportPayload
    {
        [DataMember(Name = "schema", Order = 0)] public int Schema { get; set; }
        [DataMember(Name = "event", Order = 1)] public string Event { get; set; }
        [DataMember(Name = "time", Order = 2)] public string Time { get; set; }
        [DataMember(Name = "version", Order = 3)] public string Version { get; set; }
        [DataMember(Name = "build", Order = 4)] public string Build { get; set; }
        [DataMember(Name = "os", Order = 5)] public string Os { get; set; }
        [DataMember(Name = "os64", Order = 6)] public bool Os64 { get; set; }
        [DataMember(Name = "proc64", Order = 7)] public bool Proc64 { get; set; }
        [DataMember(Name = "ram_mb", Order = 8)] public int RamMb { get; set; }
        [DataMember(Name = "cpu", Order = 9)] public string Cpu { get; set; }
        [DataMember(Name = "device", Order = 10)] public string Device { get; set; }
        [DataMember(Name = "detail", Order = 11)] public string Detail { get; set; }
        [DataMember(Name = "log", Order = 12)] public string Log { get; set; }
    }

    /// <summary>
    /// 失败上报。
    ///
    /// 【只做一件事】程序里出现"本该成功但失败了"的情况时，
    /// 把当时的日志尾部与设备基本信息发到一个 HTTP 端点。
    /// 不做行为统计、不记录用户操作、不走第三方分析平台。
    ///
    /// 【三条硬约束，改代码前先看】
    ///   1. 绝不影响主流程：所有异常在这里吞掉，超时 8 秒，
    ///      不重试、不排队、不阻塞任何线程。
    ///   2. 发出前必须脱敏：用户目录、计算机名、用户名、IPv4 全部替换。
    ///      日志里天然带路径，不脱敏等于把每个人的 C:\Users\张三 发出去。
    ///   3. 节流：同一类事件 10 分钟最多一次，单次运行最多 8 条。
    ///      崩溃往往成串出现，不限流会把端点打满。
    /// </summary>
    public static class CrashReporter
    {
        /// <summary>上报端点。留空表示不上报（开发版默认留空）。</summary>
        public static string Endpoint = "";

        /// <summary>与服务端约定的简单口令，用于过滤脚本乱发。它不是安全边界。</summary>
        public static string Token = "";

        /// <summary>由设置界面的开关控制</summary>
        public static bool Enabled = true;

        /// <summary>匿名设备标识（随机 GUID，首次运行生成，重装即变）</summary>
        public static string DeviceId = "";

        /// <summary>日志尾部最多带多少字符。够看清失败上下文，又不至于让请求过大。</summary>
        private const int LogTailChars = 12 * 1024;

        private const int SendTimeoutSeconds = 8;
        private const int ThrottleMinutes = 10;
        private const int MaxReportsPerSession = 8;

        private static readonly object Sync = new object();
        private static readonly Dictionary<string, DateTime> LastSent =
            new Dictionary<string, DateTime>(StringComparer.Ordinal);
        private static int _sentThisSession;

        /// <summary>
        /// 记录一次失败。立即返回，真正的发送在后台线程完成。
        /// </summary>
        public static void Report(string eventName, string detail)
        {
            try
            {
                if (!Enabled) return;
                if (string.IsNullOrEmpty(Endpoint)) return;
                if (string.IsNullOrEmpty(eventName)) return;
                if (!ShouldSend(eventName)) return;

                Task.Run(() => Send(eventName, detail));
            }
            catch
            {
                // 上报失败永远不能影响程序本身
            }
        }

        /// <summary>节流判断。抽成独立方法便于单测。</summary>
        public static bool ShouldSend(string eventName, DateTime now)
        {
            lock (Sync)
            {
                if (_sentThisSession >= MaxReportsPerSession) return false;

                DateTime last;
                if (LastSent.TryGetValue(eventName, out last)
                    && (now - last).TotalMinutes < ThrottleMinutes)
                {
                    return false;
                }

                LastSent[eventName] = now;
                _sentThisSession++;
                return true;
            }
        }

        private static bool ShouldSend(string eventName)
        {
            return ShouldSend(eventName, DateTime.UtcNow);
        }

        /// <summary>供测试重置节流状态</summary>
        public static void ResetThrottleForTests()
        {
            lock (Sync)
            {
                LastSent.Clear();
                _sentThisSession = 0;
            }
        }

        /// <summary>
        /// 组装上报内容。纯函数（除了读本机环境），便于测试脱敏与字段。
        /// </summary>
        public static CrashReportPayload BuildPayload(
            string eventName, string detail, DateTime nowUtc, string logText)
        {
            return new CrashReportPayload
            {
                Schema = 1,
                Event = eventName,
                Time = nowUtc.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                Version = Safe(() => Assembly.GetExecutingAssembly().GetName().Version.ToString()),
                Build = Safe(() => BuildInfo.BuildId),
                Os = Safe(() => Environment.OSVersion.VersionString),
                Os64 = Safe(() => Environment.Is64BitOperatingSystem),
                Proc64 = Safe(() => Environment.Is64BitProcess),
                RamMb = QueryTotalMemoryMb(),
                Cpu = QueryCpuName(),
                Device = DeviceId ?? "",
                Detail = Redact(detail),
                Log = Redact(logText)
            };
        }

        private static void Send(string eventName, string detail)
        {
            try
            {
                string log = ReadLogTail(LogService.LogFilePath, LogTailChars);
                CrashReportPayload payload = BuildPayload(eventName, detail, DateTime.UtcNow, log);
                string json = Serialize(payload);
                Post(json);
            }
            catch
            {
                // 静默失败：网络不通是常态，不该在日志里刷屏
            }
        }

        /// <summary>把上报内容序列化成 JSON。</summary>
        public static string Serialize(CrashReportPayload payload)
        {
            var serializer = new DataContractJsonSerializer(typeof(CrashReportPayload));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, payload);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static void Post(string json)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/json; charset=utf-8");
                client.Headers.Add("User-Agent", "SeewoOpt-Report/1");
                if (!string.IsNullOrEmpty(Token))
                    client.Headers.Add("X-Report-Token", Token);

                Task upload = client.UploadStringTaskAsync(new Uri(Endpoint), "POST", json);
                if (!upload.Wait(TimeSpan.FromSeconds(SendTimeoutSeconds)))
                {
                    client.CancelAsync();
                    return;
                }
            }
        }

        /// <summary>
        /// 从路径里抹掉能指认到人的信息。
        ///
        /// 【为什么必须做】日志里到处是完整路径（C:\Users\张三\桌面\...）与
        /// 同步服务器地址；原样上传等于把每台机器的主人姓名发出去。
        /// 只替换"能定位到人"的四类，不动其余内容，否则日志就失去排查价值。
        /// </summary>
        public static string Redact(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            string result = text;

            // 1) 用户目录整段替换（覆盖其中夹带的用户名）
            string profile = Safe(() => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            if (!string.IsNullOrEmpty(profile) && profile.Length > 3)
                result = ReplaceIgnoreCase(result, profile, "%USERPROFILE%");

            // 2) 计算机名。用 \b 边界，避免把长单词里的片段替换掉。
            string machine = Safe(() => Environment.MachineName);
            if (!string.IsNullOrEmpty(machine) && machine.Length >= 3)
                result = ReplaceWord(result, machine, "<pc>");

            // 3) 用户名。同样加边界——用户名可能恰好是某个单词的前缀。
            string user = Safe(() => Environment.UserName);
            if (!string.IsNullOrEmpty(user) && user.Length >= 3)
                result = ReplaceWord(result, user, "<user>");

            // 4) IPv4。时间服务器的地址属于环境信息，留着能帮排查，
            //    但网段/公网 IP 可能定位到学校，统一抹掉。
            result = Regex.Replace(result, @"\b(?:\d{1,3}\.){3}\d{1,3}\b", "<ip>");

            return result;
        }

        private static string ReplaceWord(string text, string word, string replacement)
        {
            try
            {
                return Regex.Replace(
                    text,
                    @"\b" + Regex.Escape(word) + @"\b",
                    replacement.Replace("$", "$$"),
                    RegexOptions.IgnoreCase);
            }
            catch
            {
                return text;
            }
        }

        private static string ReplaceIgnoreCase(string text, string from, string to)
        {
            int index = text.IndexOf(from, StringComparison.OrdinalIgnoreCase);
            if (index < 0) return text;

            var builder = new StringBuilder(text.Length);
            int cursor = 0;
            while (index >= 0)
            {
                builder.Append(text, cursor, index - cursor);
                builder.Append(to);
                cursor = index + from.Length;
                index = text.IndexOf(from, cursor, StringComparison.OrdinalIgnoreCase);
            }
            builder.Append(text, cursor, text.Length - cursor);
            return builder.ToString();
        }

        /// <summary>读取日志文件尾部。文件被占用或不存在都返回空串。</summary>
        public static string ReadLogTail(string path, int maxChars)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path)) return "";

                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    long length = stream.Length;
                    long start = Math.Max(0, length - maxChars);
                    stream.Seek(start, SeekOrigin.Begin);

                    var buffer = new byte[length - start];
                    int read = stream.Read(buffer, 0, buffer.Length);
                    string text = Encoding.UTF8.GetString(buffer, 0, read);

                    // 从中间截断可能把一行切成两半，丢掉第一行残片
                    if (start > 0)
                    {
                        int newline = text.IndexOf('\n');
                        if (newline >= 0 && newline + 1 < text.Length)
                            text = text.Substring(newline + 1);
                    }
                    return text;
                }
            }
            catch
            {
                return "";
            }
        }

        private static int QueryTotalMemoryMb()
        {
            try
            {
                var status = new MEMORYSTATUSEX();
                status.dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
                if (!GlobalMemoryStatusEx(status)) return 0;

                return (int)(status.ullTotalPhys / (1024 * 1024));
            }
            catch
            {
                return 0;
            }
        }

        private static string QueryCpuName()
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(
                    @"HARDWARE\DESCRIPTION\System\CentralProcessor\0"))
                {
                    if (key == null) return "";
                    string name = key.GetValue("ProcessorNameString") as string;
                    return string.IsNullOrEmpty(name) ? "" : name.Trim();
                }
            }
            catch
            {
                return "";
            }
        }

        private static T Safe<T>(Func<T> getter)
        {
            try { return getter(); }
            catch { return default(T); }
        }

        [StructLayout(LayoutKind.Sequential)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);
    }
}
