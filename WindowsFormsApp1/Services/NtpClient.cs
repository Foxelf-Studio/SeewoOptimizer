using System;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Threading;

namespace SeewoOpt.Services
{
    /// <summary>NTP 协议字���（NTPv4，48 字节）</summary>
    internal static class NtpProtocol
    {
        public const int PacketSize = 48;

        // 偏移 0：标志位。LI(2) VN(3) Mode(3)
        public const int OffsetLeapIndicator = 0;
        public const int OffsetVersion = 1;
        public const int OffsetMode = 2;

        // 偏移 1：Stratum（1-15，16 表示"未同步"）
        public const int OffsetStratum = 1;

        public const int ModeServer = 4;

        /// <summary>1900-01-01 到 1970-01-01 之间的秒数，NTP 纪元偏移</summary>
        public const long NtpEpochOffsetSeconds = 2208988800L;
    }

    /// <summary>从 NTP 服务器获取的时间（含协议校验结果）</summary>
    public class NtpResult
    {
        /// <summary>服务器返回的 UTC 时间</summary>
        public DateTime UtcTime { get; set; }

        /// <summary>Stratum 层数（1 最接近原子钟，16 表示不可用）</summary>
        public byte Stratum { get; set; }

        /// <summary>Mode 字段，服务器应答应为 4</summary>
        public byte Mode { get; set; }

        /// <summary>版本号，期望为 4</summary>
        public byte Version { get; set; }

        /// <summary>Leap Indicator：0 正常，3 表示时钟不同步</summary>
        public byte LeapIndicator { get; set; }

        /// <summary>是否通过了协议层校验</summary>
        public bool IsValid { get; set; }

        /// <summary>校验失败的原因（IsValid 为 true 时为空）</summary>
        public string ValidationError { get; set; }

        /// <summary>用于日志的可读摘要</summary>
        public string Describe()
        {
            if (!IsValid)
                return "NTP 响应无效：" + ValidationError;

            return string.Format("NTP 响应有效 stratum={0} mode={1} v{2} li={3} → UTC {4:yyyy-MM-dd HH:mm:ss}",
                Stratum, Mode, Version, LeapIndicator, UtcTime);
        }
    }

    /// <summary>
    /// NTP 客户端：向时间服务器查询标准时间。
    ///
    /// 【相比原实现补强了什么】
    /// 原 GetNetworkTime 只做了长度检查（bytesReceived < 48）就取时间，
    /// 未校验协议字段。实际会出问题的场景：
    /// - 收到 Stratum=16（服务器自己没同步），此时应拒绝该响应
    /// - 收到 Mode≠4（客户端/广播报文），不是服务器应答
    /// - Leap Indicator=3 表示时钟不同步，时间值不可信
    /// 这些情况下原实现会返回一个看似正常实则错误的时间。
    /// </summary>
    public static class NtpClient
    {
        private const int DefaultTimeoutMs = 5000;

        /// <summary>
        /// 查询 NTP 服务器，返回经协议校验的时间。
        /// 失败或校验不通过时返回 IsValid=false 的结果，不抛异常。
        /// </summary>
        public static NtpResult Query(string ntpServer, int timeoutMs = DefaultTimeoutMs)
        {
            var result = new NtpResult();

            try
            {
                byte[] packet = new byte[NtpProtocol.PacketSize];

                // 客户端请求：LI=0, VN=4, Mode=3
                packet[NtpProtocol.OffsetLeapIndicator] = 0x1B;

                IPAddress address = ResolveIPv4(ntpServer);
                if (address == null)
                {
                    result.ValidationError = "无法解析 NTP 服务器地址: " + ntpServer;
                    return result;
                }

                var endPoint = new IPEndPoint(address, 123);

                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                {
                    socket.ReceiveTimeout = timeoutMs;
                    socket.SendTimeout = timeoutMs;

                    int bytesReceived = Exchange(socket, endPoint, packet, timeoutMs);
                    if (bytesReceived < NtpProtocol.PacketSize)
                    {
                        result.ValidationError = string.Format(
                            "NTP 响应不完整（{0}/{1} 字节）", bytesReceived, NtpProtocol.PacketSize);
                        return result;
                    }
                }

                // 协议字段校验——这部分是原实现缺失的
                result.LeapIndicator = (byte)((packet[0] >> 6) & 0x03);
                result.Version = (byte)((packet[0] >> 3) & 0x07);
                result.Mode = (byte)(packet[0] & 0x07);
                result.Stratum = packet[NtpProtocol.OffsetStratum];

                if (result.Mode != NtpProtocol.ModeServer)
                {
                    result.ValidationError = string.Format(
                        "响应 Mode={0} 不是服务器应答（期望 {1}）", result.Mode, NtpProtocol.ModeServer);
                    return result;
                }

                if (result.Stratum == 0 || result.Stratum >= 16)
                {
                    result.ValidationError = string.Format(
                        "服务器未同步（Stratum={0}），时间不可用", result.Stratum);
                    return result;
                }

                if (result.LeapIndicator == 3)
                {
                    result.ValidationError = "服务器标记时钟不同步（LI=3），时间不可用";
                    return result;
                }

                DateTime utcTime = ParseTimestamp(packet, 40);

                // 合理性下限校验。
                // 引入本项目时曾因时间戳纪元处理错误，把 2026 年解析成 1956 年，
                // 而协议层校验（Mode/Stratum/LI）全部通过，故障静默且难以察觉。
                // 这里加一道与本机当前时间的偏差检查，使同类错误必然暴露。
                // 下限取 2000-01-01：本工具面向教室场景，真实时间不可能早于此；
                // 上限取本机时间 +1 天，容忍各服务器间的正常差异与轻微时钟漂移。
                DateTime earliest = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                DateTime latest = DateTime.UtcNow.AddDays(1);

                if (utcTime < earliest || utcTime > latest)
                {
                    result.ValidationError = string.Format(
                        "服务器时间 {0:yyyy-MM-dd HH:mm:ss} 超出合理范围（允许 {1:yyyy-MM-dd} ~ {2:yyyy-MM-dd}），解析可能有误",
                        utcTime, earliest, latest);
                    return result;
                }

                result.UtcTime = utcTime;
                result.IsValid = true;
            }
            catch (Exception ex)
            {
                result.ValidationError = ex.Message;
            }

            return result;
        }

        /// <summary>解析 NTP 64 位时间戳（秒 + 小数）</summary>
        private static DateTime ParseTimestamp(byte[] data, int offset)
        {
            ulong intPart = (ulong)data[offset] << 24
                         | (ulong)data[offset + 1] << 16
                         | (ulong)data[offset + 2] << 8
                         | data[offset + 3];

            ulong fractPart = (ulong)data[offset + 4] << 24
                            | (ulong)data[offset + 5] << 16
                            | (ulong)data[offset + 6] << 8
                            | data[offset + 7];

            // NTP 时间戳本身就是从 1900-01-01 起算的秒数，
            // 因此基准也用 1900-01-01，不能再减 Unix 纪元偏移。
            //
            // 修正记录：此处在重构时曾写成
            //   seconds = intPart - NtpEpochOffsetSeconds; 配 1900 基准
            // 等于把纪元偏移减了两次，解析结果比真实时间早 70 年
            // （日志表现为 "UTC 1956-10-03" 而时刻是对的）。
            // 两种正确写法二选一：
            //   A. intPart 原值 + 1900 基准
            //   B. intPart - 纪元偏移 + 1970 基准
            long milliseconds = (long)((fractPart * 1000) >> 32);

            return new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)
                .AddSeconds((long)intPart)
                .AddMilliseconds(milliseconds);
        }

        /// <summary>优先取 IPv4 地址，NTP 只在 UDP/IPv4 上工作</summary>
        private static IPAddress ResolveIPv4(string host)
        {
            IPAddress[] addresses = Dns.GetHostEntry(host).AddressList;

            foreach (IPAddress addr in addresses)
            {
                if (addr.AddressFamily == AddressFamily.InterNetwork)
                    return addr;
            }

            // 原实现在此处会退回到任意地址族，但 IPv6 下 NTP 不可用
            return addresses.Length > 0 ? addresses[0] : null;
        }

        /// <summary>UDP 的一次请求-应答，带超时控制</summary>
        private static int Exchange(Socket socket, EndPoint endPoint, byte[] buffer, int timeoutMs)
        {
            // UDP 无连接，"Connect" 只是绑定默认对端并设置超时
            IAsyncResult connectResult = socket.BeginConnect(endPoint, null, null);
            if (!connectResult.AsyncWaitHandle.WaitOne(timeoutMs, false))
                throw new TimeoutException("连接 NTP 服务器超时");
            socket.EndConnect(connectResult);

            IAsyncResult sendResult = socket.BeginSend(buffer, 0, buffer.Length, SocketFlags.None, null, null);
            if (!sendResult.AsyncWaitHandle.WaitOne(timeoutMs, false))
                throw new TimeoutException("发送 NTP 请求超时");
            socket.EndSend(sendResult);

            IAsyncResult receiveResult = socket.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, null, null);
            if (!receiveResult.AsyncWaitHandle.WaitOne(timeoutMs, false))
                throw new TimeoutException("接收 NTP 响应超时");

            return socket.EndReceive(receiveResult);
        }
    }

    /// <summary>系统时间设置结果</summary>
    public class SetTimeResult
    {
        /// <summary>是否设置成功</summary>
        public bool Success { get; set; }

        /// <summary>是否因缺少管理员权限而失败</summary>
        public bool PermissionDenied { get; set; }

        /// <summary>Windows 错误码（PermissionDenied 为 false 且失败时有值）</summary>
        public uint ErrorCode { get; set; }

        /// <summary>失败原因的可读描述</summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 系统时间设置。
    ///
    /// 【命名冲突修正】
    /// 原代码里包装方法叫 SetLocalTime(DateTime)，P/Invoke 也叫 SetLocalTime(ref SYSTEMTIME)，
    /// 靠重载解析区分。这在阅读时极易误认为递归调用，且一旦包装方法签名变化就会静默
    /// 变成调用自身。现分别命名为 TrySetLocalTime / SetSystemTimeNative。
    ///
    /// 【时区修正】
    /// 原实现写死 AddHours(8)，默认用户在中国时区。改用 TimeZoneInfo.ConvertTimeFromUtc，
    /// 非中国区用户也能得到正确的本地时间。
    /// </summary>
    public static class SystemTimeSetter
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct SYSTEMTIME
        {
            public short wYear;
            public short wMonth;
            public short wDayOfWeek;
            public short wDay;
            public short wHour;
            public short wMinute;
            public short wSecond;
            public short wMilliseconds;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetSystemTimeNative(ref SYSTEMTIME st);

        [DllImport("kernel32.dll")]
        private static extern uint GetLastErrorNative();

        /// <summary>Win32 错误码：ERROR_ACCESS_DENIED</summary>
        private const uint ErrorAccessDenied = 5;

        /// <summary>Win32 错误码：ERROR_PRIVILEGE_NOT_HELD</summary>
        private const uint ErrorPrivilegeNotHeld = 1314;

        /// <summary>
        /// 把 UTC 时间转换为当前系统的本地时间。
        ///
        /// 原实现是 AddHours(8) 写死东八区。此处按系统实际时区转换，
        /// 对 UTC+8 用户结果完全一致，对其他时区用户则是正确的修正。
        /// </summary>
        public static DateTime ToLocalTime(DateTime utcTime)
        {
            if (utcTime.Kind != DateTimeKind.Utc)
                utcTime = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc);

            return TimeZoneInfo.ConvertTimeFromUtc(utcTime, TimeZoneInfo.Local);
        }

        /// <summary>设置系统时间。结果通过返回值携带，不抛异常。</summary>
        public static SetTimeResult SetToLocalTime(DateTime utcTime)
        {
            DateTime local = ToLocalTime(utcTime);

            var result = new SetTimeResult();
            try
            {
                SYSTEMTIME st = new SYSTEMTIME
                {
                    wYear = (short)local.Year,
                    wMonth = (short)local.Month,
                    wDay = (short)local.Day,
                    wHour = (short)local.Hour,
                    wMinute = (short)local.Minute,
                    wSecond = (short)local.Second,
                    wMilliseconds = (short)local.Millisecond
                };

                if (SetSystemTimeNative(ref st))
                {
                    result.Success = true;
                    return result;
                }

                result.ErrorCode = GetLastErrorNative();
                result.PermissionDenied = IsPermissionError(result.ErrorCode);
                result.ErrorMessage = GetErrorMessage(result.ErrorCode);
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.PermissionDenied = IsPermissionException(ex);
                return result;
            }
        }

        /// <summary>是否权限不足的错误码</summary>
        public static bool IsPermissionError(uint errorCode)
        {
            return errorCode == ErrorAccessDenied || errorCode == ErrorPrivilegeNotHeld;
        }

        /// <summary>是否权限不足的异常</summary>
        public static bool IsPermissionException(Exception ex)
        {
            return ex is UnauthorizedAccessException
                || ex.Message.IndexOf("access", StringComparison.OrdinalIgnoreCase) >= 0
                || ex.Message.Contains("权限")
                || ex.Message.IndexOf("privilege", StringComparison.OrdinalIgnoreCase) >= 0
                || ex.Message.IndexOf("admin", StringComparison.OrdinalIgnoreCase) >= 0
                || ex.Message.Contains("管理员");
        }

        /// <summary>把 Win32 错误码转成人话</summary>
        public static string GetErrorMessage(uint errorCode)
        {
            switch (errorCode)
            {
                case ErrorAccessDenied:
                    return "访问被拒绝，需要管理员权限";
                case ErrorPrivilegeNotHeld:
                    return "缺少调整系统时间的特权";
                case 87:
                    return "参数无效";
                case 121:
                    return "通信超时";
                default:
                    return $"系统调用失败（错误码 {errorCode}）";
            }
        }
    }
}
