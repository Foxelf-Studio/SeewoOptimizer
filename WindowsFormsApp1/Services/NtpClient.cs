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
        /// 单次查询内部的发包次数。
        ///
        /// 【修正记录】原实现每次 Query 只发 1 个包。实测日志显示一个稳定规律：
        /// 每个服务器的"第 1 次尝试"几乎必然超时，而"第 2 次"立刻就有响应，
        /// 且这个规律对全部 10 个服务器一致。原因是 UDP 的 BeginConnect 只是本地
        /// 记录默认对端（不产生任何报文），随后的首个 Send 需要先完成 ARP 解析
        /// 才能落地，而 Windows 在 ARP 未就绪时会把这一包丢掉——于是每个新服务器
        /// 都白白搭上 5 秒超时。
        ///
        /// 现在改为在同一次 Query 内连续发 2 包、共用同一总预算，
        /// 第 1 包丢失时第 2 包立即补上，不必退回上层重试逻辑重来一遍。
        /// </summary>
        private const int SendAttemptsPerQuery = 2;

        /// <summary>
        /// 查询 NTP 服务器，返回经协议校验的时间。
        /// 失败或校验不通过时返回 IsValid=false 的结果，不抛异常。
        /// </summary>
        public static NtpResult Query(string ntpServer, int timeoutMs = DefaultTimeoutMs)
        {
            byte[] packet = new byte[NtpProtocol.PacketSize];

            try
            {
                // 客户端请求：LI=0, VN=4, Mode=3
                packet[NtpProtocol.OffsetLeapIndicator] = 0x1B;

                IPAddress address = ResolveIPv4(ntpServer);
                if (address == null)
                {
                    return new NtpResult { ValidationError = "无法解析 NTP 服务器地址: " + ntpServer };
                }

                var endPoint = new IPEndPoint(address, 123);

                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
                {
                    socket.ReceiveTimeout = timeoutMs;
                    socket.SendTimeout = timeoutMs;

                    int bytesReceived = Exchange(socket, endPoint, packet, timeoutMs);
                    if (bytesReceived < NtpProtocol.PacketSize)
                    {
                        return new NtpResult
                        {
                            ValidationError = string.Format(
                                "NTP 响应不完整（{0}/{1} 字节）", bytesReceived, NtpProtocol.PacketSize)
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                // 网络层的失败（超时、DNS、Socket 异常）在这里收敛成结果对象，
                // 不向调用方抛——上层要遍历 10 个服务器，抛异常会打断整个循环。
                return new NtpResult { ValidationError = ex.Message };
            }

            // 字节已拿到，交给纯函数解析。异常不会从这里逃逸。
            return ParseResponse(packet);
        }

        /// <summary>
        /// 解析并校验一个 NTP 响应报文。**纯函数**：只依赖传入的字节，
        /// 不碰网络、不读本机时钟、不抛异常。
        ///
        /// 【为什么要单独抽出来】
        /// 协议解析是整条校时链路里唯一完全确定的部分——给定字节，结果唯一。
        /// 但它原先内嵌在 Query 里，与 Socket 操作混在一起，
        /// 导致只能靠连真实服务器来"验证"，而网络又不稳定，
        /// 失败时根本无法区分是解析错、还是包丢了、还是服务器拒答。
        /// 抽成纯函数后可以构造任意畸形报文来测边界。
        ///
        /// 注意刻意不访问 DateTime.UtcNow：本工具的核心场景就是本机时钟错误，
        /// 用本机时间做合理性判断是循环论证（详见下方 earliest/latest 注释）。
        /// </summary>
        internal static NtpResult ParseResponse(byte[] packet)
        {
            var result = new NtpResult();

            if (packet == null || packet.Length < NtpProtocol.PacketSize)
            {
                result.ValidationError = string.Format(
                    "NTP 响应不完整（{0}/{1} 字节）",
                    packet == null ? 0 : packet.Length, NtpProtocol.PacketSize);
                return result;
            }

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

            // 合理性校验：只用绝对边界，绝不与本机时钟比较。
            //
            // 设计教训：此校验最初写成"上限 = DateTime.UtcNow.AddDays(1)"，
            // 结果本机时间被设为 2000-01-01 时（这正是本工具要修正的典型场景），
            // 服务器返回的真实时间被判为"超出范围"而全部拒绝，
            // 等于把工具的核心功能堵死。用本机时钟判断 NTP 时间是循环论证——
            // 正因为本机时钟不准才需要校时。
            //
            // 边界取自协议本身而非本机状态：
            //   下限 2000-01-01：能拦住纪元处理错误造成的 1956/1970 等畸形值，
            //     同时不会误伤任何真实服务器。
            //   上限 2036-02-07 06:28:16：NTP era 0 的理论终点（2^32 秒）。
            //     32 位秒字段无法表达更晚的时间——真到那时需启用 era 1，
            //     当前实现未支持，因此明确拒绝好过静默解析出错误时间。
            DateTime earliest = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime latest = new DateTime(2036, 2, 7, 6, 28, 16, DateTimeKind.Utc);

            if (utcTime < earliest || utcTime > latest)
            {
                result.ValidationError = string.Format(
                    "服务器时间 {0:yyyy-MM-dd HH:mm:ss} 超出可接受范围（{1:yyyy-MM-dd} ~ {2:yyyy-MM-dd}），解析可能有误",
                    utcTime, earliest, latest);
                return result;
            }

            result.UtcTime = utcTime;
            result.IsValid = true;
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

        /// <summary>
        /// UDP 的一次请求-应答，带超时控制。
        ///
        /// 总超时预算 timeoutMs 在多次发包之间分摊，而不是每包各给 timeoutMs，
        /// 否则单次 Query 的最坏耗时会变成 SendAttemptsPerQuery × timeoutMs。
        /// </summary>
        private static int Exchange(Socket socket, EndPoint endPoint, byte[] buffer, int timeoutMs)
        {
            // UDP 无连接，"Connect" 只是绑定默认对端并设置超时
            IAsyncResult connectResult = socket.BeginConnect(endPoint, null, null);
            if (!connectResult.AsyncWaitHandle.WaitOne(timeoutMs, false))
                throw new TimeoutException("连接 NTP 服务器超时");
            socket.EndConnect(connectResult);

            int perAttemptMs = Math.Max(1000, timeoutMs / SendAttemptsPerQuery);
            int received = 0;
            Exception lastError = null;

            for (int attempt = 1; attempt <= SendAttemptsPerQuery; attempt++)
            {
                try
                {
                    IAsyncResult sendResult = socket.BeginSend(buffer, 0, buffer.Length, SocketFlags.None, null, null);
                    if (sendResult.AsyncWaitHandle.WaitOne(perAttemptMs, false))
                        socket.EndSend(sendResult);

                    // Receive 必须重新发起：上一轮的 BeginReceive 在超时后已失效，
                    // 复用同一个 IAsyncResult 会直接抛 InvalidOperationException。
                    IAsyncResult receiveResult = socket.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, null, null);
                    if (receiveResult.AsyncWaitHandle.WaitOne(perAttemptMs, false))
                    {
                        received = socket.EndReceive(receiveResult);
                        if (received > 0)
                            return received;
                    }
                }
                catch (SocketException ex)
                {
                    // 上一包的应答可能已在路上，再发一次即可；记下来全部失败后上报
                    lastError = ex;
                }
                catch (ObjectDisposedException ex)
                {
                    lastError = ex;
                    break;
                }
            }

            if (lastError != null && received == 0)
                throw new TimeoutException("接收 NTP 响应超时（已重发 " + SendAttemptsPerQuery + " 次）: " + lastError.Message);

            throw new TimeoutException("接收 NTP 响应超时（已发送 " + SendAttemptsPerQuery + " 包）");
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
    /// 变成调用自身。现分别命名为 SetToUtc / SetSystemTimeNative。
    ///
    /// 【时区修正 —— 2026-10-04 第二次修正，修掉"多 8 小时"的 bug】
    ///
    /// 曾有两版实现，两次都在"UTC 与本地时间"上出错，值得记下来：
    ///
    /// 第一版：写死 <c>utcTime.AddHours(8)</c>，把 UTC 硬加成东八区本地时间，
    ///         再交给 SetSystemTime。非中国时区用户会得到错误结果。
    ///
    /// 第二版：改用 <c>TimeZoneInfo.ConvertTimeFromUtc</c> 做时区换算，看似更正确，
    ///         实际仍是错的。因为 <c>SetSystemTime</c> 要的参数是 **UTC**，
    ///         而它传的是**本地时间**——于是 Windows 把"本地时间"当成 UTC，
    ///         显示时又加一次时区偏移，结果**整整多出一个时区**（UTC+8 下多 8 小时）。
    ///         实测：NTP 给 UTC 05:45:42，写完后系统显示 21:45:42。
    ///
    /// 两版的共同误判是"拿到 UTC 后总要做点什么转换才能写进系统"。
    /// 事实相反：SetSystemTime 接受的就是 UTC，此处**不能做任何时区换算**。
    /// 时区只影响 Windows 如何*显示*，不影响它如何*存储*——存储始终是 UTC。
    ///
    /// 因此本类只暴露 SetToUtc，用名字把契约写在脸上，避免后来者再"顺手加个转换"。
    /// 需要本地时间时用 ToLocalTime，但那个值**不应该**喂给 SetSystemTime。
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

        /// <summary>
        /// 设置系统时间。
        ///
        /// EntryPoint 必须显式指定：DllImport 默认把 C# 方法名当作导出名，
        /// 而 kernel32.dll 里的真实导出名是 SetSystemTime。上一版为规避与
        /// 包装方法同名而把 C# 方法改名为 SetSystemTimeNative，却漏了 EntryPoint，
        /// 运行时抛 EntryPointNotFoundException（表现为每个服务器都"失败"）。
        ///
        /// 注意导出名是 SetSystemTime（收 UTC），**不是** SetLocalTime（收本地时间）。
        /// </summary>
        [DllImport("kernel32.dll", EntryPoint = "SetSystemTime", SetLastError = true)]
        private static extern bool SetSystemTimeNative(ref SYSTEMTIME st);

        [DllImport("kernel32.dll", EntryPoint = "GetLastError")]
        private static extern uint GetLastErrorNative();

        /// <summary>Win32 错误码：ERROR_ACCESS_DENIED</summary>
        private const uint ErrorAccessDenied = 5;

        /// <summary>Win32 错误码：ERROR_PRIVILEGE_NOT_HELD</summary>
        private const uint ErrorPrivilegeNotHeld = 1314;

        /// <summary>
        /// 把 UTC 时间转换为当前系统的本地时间。
        ///
        /// 【仅供显示与日志使用】不要把这个结果传给 SetToUtc——
        /// 那正是 2026-10-04 那个"多 8 小时" bug 的成因。
        /// </summary>
        public static DateTime ToLocalTime(DateTime utcTime)
        {
            if (utcTime.Kind != DateTimeKind.Utc)
                utcTime = DateTime.SpecifyKind(utcTime, DateTimeKind.Utc);

            return TimeZoneInfo.ConvertTimeFromUtc(utcTime, TimeZoneInfo.Local);
        }

        /// <summary>
        /// 一个待写入系统的时刻，按 UTC 拆成系统时钟所需的各字段。
        ///
        /// 【为什么单独抽出来】这是"多 8 小时" bug 的实际发生地，而它原先藏在
        /// SetToUtc 内部、被 SYSTEMTIME（private）和 SetSystemTimeNative（需管理员）
        /// 挡在后面，无法在单元测试里断言。抽成 public 纯函数后，
        /// "喂进去的秒数是否等于 UTC 的秒数"可以被直接验证，
        /// 不必真的去改测试机的时钟。
        ///
        /// 【契约】所有字段取自传入时间的 UTC 表示。此结构不含任何时区信息，
        /// 时区由 Windows 在显示时应用。
        /// </summary>
        public struct ClockFields
        {
            public int Year, Month, Day, Hour, Minute, Second, Millisecond;

            public override string ToString()
            {
                return string.Format("{0:D4}-{1:D2}-{2:D2} {3:D2}:{4:D2}:{5:D2}.{6:D3}",
                    Year, Month, Day, Hour, Minute, Second, Millisecond);
            }
        }

        /// <summary>
        /// 把传入的 UTC 时刻拆成系统时钟字段。
        ///
        /// 【关键】这里**绝不做时区换算**。SetSystemTime 收的就是 UTC，
        /// 转换一次就会多出一个时区偏移（UTC+8 下多 8 小时）。
        /// 唯一的换算发生在 Kind == Local 时——那是在把调用方误传的
        /// 本地时间纠正回 UTC，方向与"UTC 转本地"相反，不要搞反。
        /// </summary>
        public static ClockFields ToClockFields(DateTime utcTime)
        {
            DateTime utc = utcTime;
            if (utc.Kind == DateTimeKind.Local)
                utc = utc.ToUniversalTime();          // 纠正误传的本地时间
            else if (utc.Kind == DateTimeKind.Unspecified)
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);   // 契约上视为 UTC

            return new ClockFields
            {
                Year = utc.Year,
                Month = utc.Month,
                Day = utc.Day,
                Hour = utc.Hour,
                Minute = utc.Minute,
                Second = utc.Second,
                Millisecond = utc.Millisecond
            };
        }

        /// <summary>
        /// 用 UTC 时间设置系统时钟。
        ///
        /// 入参必须是 UTC。这里**刻意不做任何时区换算**：
        /// SetSystemTime 收的就是 UTC，Windows 会在显示时自行应用当前时区。
        /// 多转换一次就会多出一个时区偏移（UTC+8 下多 8 小时）。
        ///
        /// 结果通过返回值携带，不抛异常。
        /// </summary>
        public static SetTimeResult SetToUtc(DateTime utcTime)
        {
            ClockFields f = ToClockFields(utcTime);

            var result = new SetTimeResult();
            try
            {
                SYSTEMTIME st = new SYSTEMTIME
                {
                    wYear = (short)f.Year,
                    wMonth = (short)f.Month,
                    wDay = (short)f.Day,
                    wHour = (short)f.Hour,
                    wMinute = (short)f.Minute,
                    wSecond = (short)f.Second,
                    wMilliseconds = (short)f.Millisecond
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
