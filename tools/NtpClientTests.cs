using System;
using SeewoOpt.Services;

/// <summary>
/// NtpClient 协议解析的单元测试。
///
/// 【测什么】NtpClient.ParseResponse —— 给定 48 字节报文，产出解析结果。
/// 这是整条校时链路里唯一完全确定的部分，也是历史 bug 最集中的地方：
///   - 纪元叠加错误（解析结果早 70 年，日志表现为 "1956-10-03"）
///   - 合理性校验用本机时钟做上界（循环论证，把正确时间判为非法）
///   - 缺失 Mode / Stratum / LI 校验（把非应答包当应答）
///
/// 【为什么能测】解析已抽为纯函数：不碰网络、不读本机时钟、不抛异常。
/// 因此可以构造任意畸形报文，包括真实服务器绝不会返回的那种。
///
/// 【为什么用链接源码而不是复刻】见 NtpClientTests.csproj 的注释。
/// </summary>
internal static class NtpClientTests
{
    private static int _passed;
    private static int _failed;

    private static int Main()
    {
        Console.WriteLine("=== NtpClient 协议解析测试 ===");
        Console.WriteLine();

        // --- 时间解析与纪元 ---
        TestValidResponseParsesUtcTime();
        TestTimestampIsEpochCorrect();
        TestFractionalSecondPrecision();

        // --- 合理性边界 ---
        TestRejectsBefore2000();
        TestRejectsAtEra0Overflow();
        TestAcceptsEra0LastValidSecond();
        TestAcceptsWithWrongLocalClock();

        // --- 协议字段校验 ---
        TestRejectsNonServerMode();
        TestRejectsUnsyncedStratum();
        TestRejectsAlarmLeapIndicator();

        // --- 系统时钟写入契约（防"多 8 小时"回归）---
        TestClockFieldsAreUtcNotLocal();
        TestClockFieldsHandleKind();

        // --- 日志运行分段 ---
        TestLogTrimKeepsLimit();
        TestLogTrimEdgeCases();
        TestLogSessionCounting();
        TestLogNumberingMonotonic();
        TestLogBaseRoundTrip();

        // --- 畸形输入 ---
        TestRejectsShortPacket();
        TestRejectsNullPacket();
        TestDoesNotThrowOnAllZeroPacket();
        TestDoesNotThrowOnGarbage();

        Console.WriteLine();
        Console.WriteLine($"通过 {_passed} 项，失败 {_failed} 项");
        Console.WriteLine(_failed == 0 ? "全部通过" : "存在失败");
        return _failed == 0 ? 0 : 1;
    }

    // ---------------------------------------------------------------
    // 辅助：构造一个合法的 NTP 服务器响应报文
    // ---------------------------------------------------------------

    /// <summary>
    /// 构造响应报文。
    /// 字节 0 布局：LI(2bit) VN(3bit) Mode(3bit)
    /// </summary>
    private static byte[] BuildPacket(
        byte mode = 4,
        byte stratum = 2,
        byte leapIndicator = 0,
        DateTime? utcTime = null,
        byte version = 4)
    {
        byte[] packet = new byte[48];

        packet[0] = (byte)((leapIndicator << 6) | (version << 3) | mode);
        packet[1] = stratum;

        DateTime t = utcTime ?? new DateTime(2026, 10, 4, 3, 30, 0, DateTimeKind.Utc);
        WriteTimestamp(packet, 40, t);

        return packet;
    }

    /// <summary>
    /// 把 DateTime 写成 NTP 64 位时间戳（32 位秒 + 32 位小数），
    /// 基准为 1900-01-01。这是协议的独立实现，故意不同于产品代码，
    /// 免得"用被测代码验证被测代码"。
    /// </summary>
    private static void WriteTimestamp(byte[] packet, int offset, DateTime utc)
    {
        DateTime epoch1900 = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        TimeSpan delta = utc - epoch1900;

        ulong seconds = (ulong)delta.TotalSeconds;
        ulong fraction = (ulong)((delta.TotalSeconds - Math.Floor(delta.TotalSeconds)) * 4294967296.0);

        packet[offset]     = (byte)(seconds >> 24);
        packet[offset + 1] = (byte)(seconds >> 16);
        packet[offset + 2] = (byte)(seconds >> 8);
        packet[offset + 3] = (byte)(seconds);

        packet[offset + 4] = (byte)(fraction >> 24);
        packet[offset + 5] = (byte)(fraction >> 16);
        packet[offset + 6] = (byte)(fraction >> 8);
        packet[offset + 7] = (byte)(fraction);
    }

    // ---------------------------------------------------------------
    // 时间解析与纪元
    // ---------------------------------------------------------------

    private static void TestValidResponseParsesUtcTime()
    {
        var expected = new DateTime(2026, 10, 4, 3, 30, 0, DateTimeKind.Utc);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: expected));

        Check("合法响应可解析", result.IsValid,
            result.IsValid ? null : result.ValidationError);

        if (result.IsValid)
        {
            Check("解析出的时间与输入一致（误差 < 10ms）",
                Math.Abs((result.UtcTime - expected).TotalMilliseconds) < 10,
                $"期望 {expected:O}，实际 {result.UtcTime:O}");
        }
    }

    /// <summary>
    /// 纪元正确性。历史上的真实 bug：解析结果比真实时间早 70 年
    /// （把 1900 基准与 Unix 纪元偏移同时减了）。这里直接断言年份。
    /// </summary>
    private static void TestTimestampIsEpochCorrect()
    {
        var input = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: input));

        if (!result.IsValid) { Check("纪元测试：应解析成功", false, result.ValidationError); return; }

        Check("纪元正确（不得早 70 年）", result.UtcTime.Year == 2026,
            $"期望年份 2026，实际 {result.UtcTime.Year}");

        // 反向确认：若误减纪元偏移，结果会落在 1956 年左右
        Check("纪元测试：不得落在 1956 年", result.UtcTime.Year != 1956,
            $"实际 {result.UtcTime:O}");
    }

    private static void TestFractionalSecondPrecision()
    {
        var input = new DateTime(2026, 6, 15, 12, 0, 0, DateTimeKind.Utc).AddMilliseconds(500);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: input));

        if (!result.IsValid) { Check("小数秒精度：应解析成功", false, result.ValidationError); return; }

        double actualMs = result.UtcTime.Millisecond;
        Check("小数秒精度（500ms 误差 < 20ms）", Math.Abs(actualMs - 500) < 20,
            $"期望约 500ms，实际 {actualMs}ms");
    }

    // ---------------------------------------------------------------
    // 合理性边界
    // ---------------------------------------------------------------

    private static void TestRejectsBefore2000()
    {
        // 1999-12-31 应被拒绝（下限 2000-01-01）
        var input = new DateTime(1999, 12, 31, 23, 59, 59, DateTimeKind.Utc);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: input));

        Check("拒绝 1999 年（低于下限）", !result.IsValid,
            result.IsValid ? $"却接受了 {result.UtcTime:O}" : null);
    }

    private static void TestRejectsAtEra0Overflow()
    {
        // era 0 终点之后（2036-02-07 06:28:17）应被拒绝
        var input = new DateTime(2036, 2, 7, 6, 28, 17, DateTimeKind.Utc);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: input));

        Check("拒绝 era 0 溢出边界之后的时间", !result.IsValid,
            result.IsValid ? $"却接受了 {result.UtcTime:O}" : null);
    }

    private static void TestAcceptsEra0LastValidSecond()
    {
        // era 0 最后一个有效时刻 2036-02-07 06:28:15 应被接受
        var input = new DateTime(2036, 2, 7, 6, 28, 15, DateTimeKind.Utc);
        var result = NtpClient.ParseResponse(BuildPacket(utcTime: input));

        Check("接受 era 0 最后有效时刻", result.IsValid,
            result.IsValid ? null : result.ValidationError);
    }

    /// <summary>
    /// 时钟无关性：本工具的核心场景就是本机时钟错误。
    /// 若校验里混入 DateTime.UtcNow 作上界，本机时钟停在过去时
    /// 会把正确时间判为"超出范围"——正是曾经的循环论证缺陷。
    ///
    /// 纯函数不读本机时钟，因此本机时钟是什么都该通过。
    ///
    /// 【为什么要跨一个时间跨度断点】
    /// 只测"当下这一刻"是不够的：若上界写成 UtcNow.AddDays(1)，
    /// 而本机时钟恰好正常，那么所有接近当下的用例都会通过，
    /// 循环论证被掩盖。这里刻意用三个跨越 10 年的未来时间点，
    /// 本机时钟正常时它们就会越界——故障才暴露得出来。
    /// </summary>
    private static void TestAcceptsWithWrongLocalClock()
    {
        Console.WriteLine($"    （本机当前 UTC = {DateTime.UtcNow:yyyy-MM-dd HH:mm}）");

        // 未来 10 年内的多个时间点：服务器时间本就该"比本机晚"，
        // 若校验拿本机时钟当上界，这几项必然失败。
        DateTime[] futureMoments =
        {
            new DateTime(2026, 10, 4, 4, 0, 0, DateTimeKind.Utc),   // 当下附近
            new DateTime(2030, 1, 1, 0, 0, 0, DateTimeKind.Utc),    // 中期
            new DateTime(2036, 2, 7, 6, 0, 0, DateTimeKind.Utc)     // era 0 末期
        };

        foreach (DateTime moment in futureMoments)
        {
            var result = NtpClient.ParseResponse(BuildPacket(utcTime: moment));

            Check($"本机时钟无关：接受 {moment:yyyy-MM-dd}（服务器时间可晚于本机）",
                result.IsValid,
                result.IsValid ? null : $"被拒绝：{result.ValidationError}");
        }

        // 更强的断言：解析结果必须与输入**完全相等**，而不只是 IsValid。
        // 这样即使校验放宽了，纪元偏移之类的错误也逃不掉。
        var exact = new DateTime(2028, 5, 20, 14, 30, 45, DateTimeKind.Utc);
        var exactResult = NtpClient.ParseResponse(BuildPacket(utcTime: exact));
        Check("解析结果与输入逐秒一致（比 IsValid 更强的断言）",
            exactResult.IsValid
                && exactResult.UtcTime.Year == exact.Year
                && exactResult.UtcTime.Month == exact.Month
                && exactResult.UtcTime.Day == exact.Day
                && exactResult.UtcTime.Hour == exact.Hour
                && exactResult.UtcTime.Minute == exact.Minute
                && exactResult.UtcTime.Second == exact.Second,
            exactResult.IsValid ? $"期望 {exact:O}，实际 {exactResult.UtcTime:O}" : exactResult.ValidationError);
    }

    // ---------------------------------------------------------------
    // 协议字段校验
    // ---------------------------------------------------------------

    private static void TestRejectsNonServerMode()
    {
        // Mode=3 是客户端请求，不是服务器应答
        foreach (byte mode in new byte[] { 0, 1, 2, 3, 5, 6, 7 })
        {
            var result = NtpClient.ParseResponse(BuildPacket(mode: mode));
            Check($"拒绝非服务器应答 Mode={mode}", !result.IsValid,
                result.IsValid ? "却被接受" : null);
        }
    }

    private static void TestRejectsUnsyncedStratum()
    {
        // Stratum=0（未同步 / KoD）与 >=16（保留未定义）都应拒绝
        foreach (byte stratum in new byte[] { 0, 16, 17, 255 })
        {
            var result = NtpClient.ParseResponse(BuildPacket(stratum: stratum));
            Check($"拒绝 Stratum={stratum}", !result.IsValid,
                result.IsValid ? "却被接受" : null);
        }

        // 边界内应接受
        foreach (byte stratum in new byte[] { 1, 15 })
        {
            var result = NtpClient.ParseResponse(BuildPacket(stratum: stratum));
            Check($"接受 Stratum={stratum}", result.IsValid,
                result.IsValid ? null : result.ValidationError);
        }
    }

    private static void TestRejectsAlarmLeapIndicator()
    {
        // LI=3 表示服务器时钟未同步
        var result = NtpClient.ParseResponse(BuildPacket(leapIndicator: 3));
        Check("拒绝 LI=3（服务器时钟未同步）", !result.IsValid,
            result.IsValid ? "却被接受" : null);

        // LI=0/1/2 应接受
        foreach (byte li in new byte[] { 0, 1, 2 })
        {
            var result2 = NtpClient.ParseResponse(BuildPacket(leapIndicator: li));
            Check($"接受 LI={li}", result2.IsValid,
                result2.IsValid ? null : result2.ValidationError);
        }
    }

    // ---------------------------------------------------------------
    // 畸形输入
    // ---------------------------------------------------------------

    private static void TestRejectsShortPacket()
    {
        foreach (int len in new int[] { 0, 1, 47 })
        {
            var result = NtpClient.ParseResponse(new byte[len]);
            Check($"拒绝长度 {len} 的短报文", !result.IsValid,
                result.IsValid ? "却被接受" : null);
        }

        // 49 字节（多出）应可接受——协议允许扩展字段
        var packet = BuildPacket();
        var longer = new byte[64];
        Array.Copy(packet, longer, packet.Length);
        var result49 = NtpClient.ParseResponse(longer);
        Check("接受超出 48 字节的报文（允许扩展）", result49.IsValid,
            result49.IsValid ? null : result49.ValidationError);
    }

    private static void TestRejectsNullPacket()
    {
        var result = NtpClient.ParseResponse(null);
        Check("拒绝 null 报文（不抛异常）", !result.IsValid,
            result.IsValid ? "却被接受" : null);
    }

    private static void TestDoesNotThrowOnAllZeroPacket()
    {
        // 全零报文：Mode=0, Stratum=0 → 应被拒绝而非抛出
        bool threw = false;
        NtpResult result = null;
        try { result = NtpClient.ParseResponse(new byte[48]); }
        catch { threw = true; }

        Check("全零报文不抛异常", !threw, threw ? "抛出了异常" : null);
        if (!threw)
            Check("全零报文被拒绝", !result.IsValid,
                result.IsValid ? "却被接受" : null);
    }

    private static void TestDoesNotThrowOnGarbage()
    {
        // 随机字节：解析必须稳健，绝不抛异常
        var rng = new Random(20261004);
        int threw = 0;

        for (int i = 0; i < 500; i++)
        {
            byte[] packet = new byte[48];
            rng.NextBytes(packet);

            try { NtpClient.ParseResponse(packet); }
            catch { threw++; }
        }

        Check("500 组随机字节均不抛异常", threw == 0,
            threw > 0 ? $"{threw} 组抛出异常" : null);
    }

    // ---------------------------------------------------------------
    // 系统时钟写入契约（2026-10-04 "多 8 小时" bug 的防回归）
    //
    // 【Bug 回顾】SetSystemTime 接受的是 UTC，代码却传了 ConvertTimeFromUtc
    // 之后的本地时间。Windows 把该值当 UTC，显示时再加一次时区偏移，
    // 结果整整多出一个时区。实测 UTC 05:45:42 写完后显示 21:45:42。
    //
    // 【第一版测试写错了，值得记下来】最初只断言"ToLocalTime 确实应用了偏移"
    // 和"本地小时 ≠ UTC 小时"——这些是辅助函数的性质，与被测函数无关。
    // 实测证明：把 bug 注回产品代码，那版测试**依然全绿**，什么都没抓住。
    //
    // 【现在怎么做】SystemTimeSetter.ToClockFields 是 bug 的实际发生地，
    // 已被抽为 public 纯函数。测试直接断言它的输出字段：
    //   - 喂 UTC 05:45:42，输出的小时必须还是 5，不能是 13
    // 这样注释掉产品代码里的修复、改回 ToLocalTime，断言立刻失败（见下方对照）。
    // ---------------------------------------------------------------

    /// <summary>
    /// 核心契约：ToClockFields 的输出字段必须等于输入 UTC 的字段，不得被时区改写。
    /// </summary>
    private static void TestClockFieldsAreUtcNotLocal()
    {
        DateTime utc = new DateTime(2026, 10, 4, 5, 45, 42, 123, DateTimeKind.Utc);
        TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(utc);
        DateTime local = utc + offset;

        Console.WriteLine($"    （本机偏移 = {offset}，UTC {utc:HH:mm:ss} 对应的本地时间是 {local:HH:mm:ss}）");

        SystemTimeSetter.ClockFields f = SystemTimeSetter.ToClockFields(utc);

        // 这一条是本次 bug 的正身。若产品代码里误把 utc 换成本地时间，
        // 在 UTC+8 下 Hour 会变成 13 而失败。
        Check($"Hour 取自 UTC，未被时区改写（得 {f.Hour}，期望 {utc.Hour}）",
            f.Hour == utc.Hour,
            $"得到 {f.Hour}，期望 {utc.Hour}"
            + (f.Hour == local.Hour && offset != TimeSpan.Zero
               ? $"  ← 看起来像被当成本地时间了（本地 {local.Hour}）" : ""));

        Check($"Minute 取自 UTC（得 {f.Minute}，期望 {utc.Minute}）", f.Minute == utc.Minute, null);
        Check($"Second 取自 UTC（得 {f.Second}，期望 {utc.Second}）", f.Second == utc.Second, null);
        Check($"Millisecond 取自 UTC（得 {f.Millisecond}，期望 {utc.Millisecond}）", f.Millisecond == utc.Millisecond, null);
        Check($"Year/Month/Day 取自 UTC（得 {f.Year}-{f.Month:D2}-{f.Day:D2}）",
            f.Year == utc.Year && f.Month == utc.Month && f.Day == utc.Day, null);

        // 边界：跨越 UTC 日界时，本地日期已变而 UTC 日期未变——最能暴露时区换算
        DateTime edge = new DateTime(2026, 10, 4, 17, 30, 0, DateTimeKind.Utc); // UTC+8 → 次日 01:30
        SystemTimeSetter.ClockFields fe = SystemTimeSetter.ToClockFields(edge);
        if (offset == TimeSpan.FromHours(8))
        {
            Check("跨日界：UTC 17:30 仍记为 4 日 17:30，不得变成 5 日 01:30",
                fe.Day == 4 && fe.Hour == 17,
                $"得到 {fe}");
        }
        else
        {
            Check($"跨日界检查（本机偏移 {offset}，仅校验等于 UTC）",
                fe.Day == edge.Day && fe.Hour == edge.Hour,
                $"得到 {fe}");
        }
    }

    /// <summary>
    /// 契约：入参 Kind 决定是否换算——Local 会被纠正回 UTC，Utc/Unspecified 原样使用。
    /// </summary>
    private static void TestClockFieldsHandleKind()
    {
        DateTime utcValue = new DateTime(2026, 10, 4, 5, 45, 42, DateTimeKind.Utc);
        SystemTimeSetter.ClockFields expected = SystemTimeSetter.ToClockFields(utcValue);

        // Utc 与 Unspecified 结果应完全一致
        var unspecified = SystemTimeSetter.ToClockFields(
            new DateTime(2026, 10, 4, 5, 45, 42, DateTimeKind.Unspecified));
        Check("Unspecified 视为 UTC，不叠加时区",
            unspecified.Hour == expected.Hour && unspecified.Day == expected.Day,
            $"得到 {unspecified}，期望 {expected}");

        // Local 应被纠正回 UTC：本地时刻 → ToUniversalTime → 字段
        DateTime localMoment = new DateTime(2026, 10, 4, 13, 45, 42, DateTimeKind.Local);
        var fromLocal = SystemTimeSetter.ToClockFields(localMoment);
        DateTime asUtc = localMoment.ToUniversalTime();
        Check($"Local 被纠正回 UTC（13:45:42 本地 → {asUtc:HH:mm:ss} UTC）",
            fromLocal.Hour == asUtc.Hour,
            $"得到 {fromLocal.Hour}，期望 {asUtc.Hour}");
    }

    // ---------------------------------------------------------------
    // 日志运行分段（2026-10-04 新增功能：把每次启动的日志分开）
    //
    // 【测什么】LogService.TrimContent —— 给定日志全文，返回应保留的部分。
    // 它是纯函数，直接链接产品源码测试，不碰真实日志文件。
    //
    // 【为什么必须测】裁剪边界连续出过两次 off-by-one：
    //   1. 分隔头同时用作首尾边框 → 同一标记每段出现两次 → 计数翻倍
    //   2. 按 KeepSessions 裁剪 → 每次固定多留一段（8 次跑出 6 段）
    // 两次都是"看起来对、跑起来多一段"的类型，肉眼极难发现。
    // ---------------------------------------------------------------

    private const string LogBegin = "======== RUN 开始 ========";
    private const string LogEnd = "======== RUN 结束 ========";

    /// <summary>构造一段模拟的日志运行段落，结构同 LogService 实际写出的内容</summary>
    private static string MakeLogSession(int n)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.AppendLine();
        sb.AppendLine(LogBegin);
        sb.AppendLine($"启动时间: 2026-10-04 14:0{n % 10}:00.000");
        sb.AppendLine($"运行序号: 第 {n} 次");
        // 基准行：真实格式里每段都有，值是"这一段自己的序号"。
        // fixture 必须跟着写，否则测不出"基准丢失导致序号卡住"这类问题。
        sb.AppendLine($"序号基准: {n}");
        // 分隔线必须引用产品常量，不能把字面量抄一遍——
        // 抄一遍的话，产品改了分隔线而测试没跟着改，测试依然全绿，
        // 就等于没测。这里刻意暴露 LogService.Divider 就是为断掉这条退路。
        sb.AppendLine(LogService.Divider);
        sb.AppendLine($"2026-10-04 14:0{n % 10}:00.100 - 第 {n} 次运行的日志");
        sb.AppendLine(LogEnd);
        return sb.ToString();
    }

    private static string MakeLogContent(int sessions)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        for (int i = 1; i <= sessions; i++) sb.Append(MakeLogSession(i));
        return sb.ToString();
    }

    /// <summary>
    /// 核心契约：裁剪后原有段数不得超过 Keep-1（因为紧接着还要追加新段）。
    /// 这条测试直接对应"8 次跑出 6 段"那个 bug。
    /// </summary>
    private static void TestLogTrimKeepsLimit()
    {
        const int Keep = 5;

        for (int sessions = 1; sessions <= 12; sessions++)
        {
            string content = MakeLogContent(sessions);
            string kept = LogService.TrimContent(content, Keep);
            int beforeAppend = LogService.CountSessions(kept);

            // 裁剪后 + 即将追加的 1 段，总数不得超过 Keep
            bool ok = beforeAppend <= Keep - 1;
            Check($"{sessions} 次运行裁剪后剩 {beforeAppend} 段（加上新段共 {beforeAppend + 1} <= {Keep}）",
                ok,
                ok ? null : $"裁剪后 {beforeAppend} 段，加新段将变成 {beforeAppend + 1} 段，超出上限 {Keep}");
        }

        // 裁剪后追加新段，验证最终正好等于 Keep
        string full = MakeLogContent(12);
        string afterTrim = LogService.TrimContent(full, Keep);
        string afterAppend = afterTrim + MakeLogSession(13);
        Check($"裁剪后追加新段，最终恰好 {Keep} 段",
            LogService.CountSessions(afterAppend) == Keep,
            $"得到 {LogService.CountSessions(afterAppend)} 段");

        Check("追加后保留的是最新几次（第 13 次在内）",
            afterAppend.Contains("第 13 次运行的日志"));
        Check("追加后已丢弃最早的（第 1 次）",
            !afterAppend.Contains("第 1 次运行的日志"));
    }

    /// <summary>边界：未超上限时不得改动内容；开头不得残留空行</summary>
    private static void TestLogTrimEdgeCases()
    {
        const int Keep = 5;

        // 未超上限：原样返回
        string under = MakeLogContent(3);
        Check("3 段 < 上限，内容原样不动",
            LogService.TrimContent(under, Keep) == under);

        // 空内容
        Check("空内容返回空", LogService.TrimContent("", Keep) == "");

        // 恰好在上限边界：Keep-1 段时不动
        string exact = MakeLogContent(Keep - 1);
        Check($"{Keep - 1} 段时不裁剪",
            LogService.TrimContent(exact, Keep) == exact);

        // 多一段：应裁掉最旧一段
        string over = MakeLogContent(Keep);
        string trimmed = LogService.TrimContent(over, Keep);
        Check($"{Keep} 段时裁掉最旧一段，剩 {Keep - 1} 段",
            LogService.CountSessions(trimmed) == Keep - 1,
            $"得到 {LogService.CountSessions(trimmed)} 段");

        // 裁剪结果不得以空行/空白开头
        Check("裁剪结果不以空行开头",
            trimmed.Length > 0 && trimmed[0] != '\r' && trimmed[0] != '\n' && trimmed[0] != ' ',
            "首字符 = " + (trimmed.Length > 0 ? ((int)trimmed[0]).ToString() : "空"));

        Check("裁剪结果首行就是开始标记",
            trimmed.StartsWith(LogBegin),
            "首行 = " + trimmed.Split('\n')[0].Trim());

        // 升级前的历史日志（无标记）不得被破坏
        string legacy = "2000-01-01 00:00:03.400 - 旧日志\n2000-01-01 00:00:04.000 - 更多旧日志\n";
        Check("无标记的历史日志原样保留",
            LogService.TrimContent(legacy, Keep) == legacy);
    }

    /// <summary>
    /// 计数契约：分隔头每次运行必须只出现一次。
    ///
    /// 曾把同一串标记同时用作首尾边框，导致每段被数两次、裁剪边界全错。
    /// 这条测试守住"一段 = 一个标记"这个前提。
    /// </summary>
    private static void TestLogSessionCounting()
    {
        Check("空内容计 0 段", LogService.CountSessions("") == 0);

        for (int n = 1; n <= 6; n++)
        {
            string c = MakeLogContent(n);
            Check($"{n} 段内容计为 {n}", LogService.CountSessions(c) == n,
                $"得到 {LogService.CountSessions(c)}");
        }

        // 一段里标记只出现一次（防止再次引入"首尾都用同一标记"的错误）
        string one = MakeLogSession(1);
        int markerCount = 0, idx = 0;
        while ((idx = one.IndexOf(LogBegin, idx, StringComparison.Ordinal)) >= 0)
        {
            markerCount++;
            idx += LogBegin.Length;
        }
        Check("单段内开始标记只出现 1 次（不得同时用作边框）",
            markerCount == 1,
            $"出现 {markerCount} 次——会把段数数成两倍");

        // 把"计数锚点只能绑在开始标记上"这条约束钉进测试。
        //
        // 背景：曾设想用 LogService.Divider 替换计数锚点来验证测试的鉴别力，
        // 实测发现——分隔线在每段里也恰好出现一次，段数照样数得对，
        // 所以"锚点换成 Divider"本身并不构成 bug。这条断言因此改为
        // 验证那件真正重要的事：计数**与分隔线无关**。
        //
        // 若哪天有人把计数锚点绑到分隔线上，一旦产品修改分隔线文案，
        // 日志归档就会静默失效（段数恒为 0，永不再裁剪），
        // 而所有既有测试仍会全绿。下面模拟的正是这种漂移。
        string drifted = LogService.Divider.Replace('-', '=');
        string oneDrifted = one.Replace(LogService.Divider, drifted);
        Check("计数与分隔线无关：分隔线被改动，段数不变",
            LogService.CountSessions(oneDrifted) == 1,
            $"得到 {LogService.CountSessions(oneDrifted)} 段——计数被分隔线影响了");

        // 同一条约束的反面：开始标记一旦被改动，计数必须立刻失真。
        // 若这条也"通过"，说明计数锚点根本没绑在开始标记上。
        string beginDrifted = one.Replace(LogBegin, "======== RUN 启动 ========");
        Check("计数锚定开始标记：开始标记被改动，计数失真",
            LogService.CountSessions(beginDrifted) == 0,
            $"得到 {LogService.CountSessions(beginDrifted)} 段——计数锚点没绑在开始标记上");
    }

    // ---------------------------------------------------------------
    // 运行序号必须持续递增（2026-10-04 端到端实跑才发现的问题）
    //
    // 【症状】跑满 KeepSessions 次之后，每一段的"运行序号"都显示成
    // 同一个数字（实测 5 段全是"第 6 次启动"）。
    //
    // 【根因】序号曾经等于 CountSessions(文件) + 1。文件裁剪到上限后
    // 段数恒为 KeepSessions，序号自然就卡死在同一个值。
    //
    // 【为什么单测没抓到】既有测试只验证了"裁剪后剩几段"，
    // 从没验证过"写进日志里的那个序号"。段数对、序号错，
    // 全绿却毫无察觉。下面这两组测试专门盯序号本身。
    // ---------------------------------------------------------------

    /// <summary>序号契约：模拟真实写入循环，序号必须严格递增且不重复</summary>
    private static void TestLogNumberingMonotonic()
    {
        const int Keep = 5;
        const int Runs = 20;

        string content = "";
        int prev = 0;
        System.Collections.Generic.List<int> seen = new System.Collections.Generic.List<int>();
        bool strictlyIncreasing = true;
        string failureDetail = null;

        for (int run = 1; run <= Runs; run++)
        {
            // 严格照 BeginSession 的顺序：先读序号，再裁剪，再写入
            int number = LogService.NextSessionNumber(content);

            if (number != prev + 1)
            {
                strictlyIncreasing = false;
                if (failureDetail == null)
                    failureDetail = $"第 {run} 次应得序号 {prev + 1}，实得 {number}";
            }
            prev = number;
            seen.Add(number);

            // 裁剪（与产品同样的调用顺序：算完序号才裁）
            string trimmed = LogService.TrimContent(content, Keep);

            // 追加本次段。基准写的就是本次序号本身。
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendLine(LogBegin);
            sb.AppendLine($"启动时间: 2026-10-04 14:00:{run % 60:D2}.000");
            sb.AppendLine($"运行序号: 第 {number} 次");
            sb.AppendLine($"序号基准: {number}");
            sb.AppendLine(LogService.Divider);
            sb.AppendLine(LogEnd);
            content = trimmed + sb.ToString();
        }

        Check($"连续 {Runs} 次运行，序号严格递增（{seen[0]}..{seen[seen.Count - 1]}）",
            strictlyIncreasing, failureDetail);

        Check($"连续 {Runs} 次运行，序号无重复",
            new System.Collections.Generic.HashSet<int>(seen).Count == seen.Count,
            $"出现重复：{string.Join(",", seen)}");

        Check($"序号到达上限后仍在增长（末次 {seen[seen.Count - 1]} > 上限 {Keep}）",
            seen[seen.Count - 1] > Keep,
            $"末次序号 {seen[seen.Count - 1]}，说明裁剪后序号卡死了");

        // 段数仍必须被压在 Keep 以内——修序号不能把保留策略弄坏
        Check($"修复序号后段数仍不超过 {Keep}",
            LogService.CountSessions(content) <= Keep,
            $"得到 {LogService.CountSessions(content)} 段");
    }

    /// <summary>基准往返契约：写进去的基准必须能被读回来，否则序号会重置</summary>
    private static void TestLogBaseRoundTrip()
    {
        // 首次运行：空文件、无基准 → 序号必须是 1
        Check("首次运行（空文件）序号为 1",
            LogService.NextSessionNumber("") == 1,
            $"得到 {LogService.NextSessionNumber("")}");

        // 无基准的旧版日志：退化为"段数 + 1"，保证升级后不会突然跳到天文数字
        string legacy = MakeLogContent(3);
        Check("无基准的旧日志退化为段数 + 1（得 4）",
            LogService.NextSessionNumber(legacy) == 4,
            $"得到 {LogService.NextSessionNumber(legacy)}");

        // 基准行可读回：最后一条生效。
        // 语义是"基准 = 上一次运行写下的序号"，所以本次 = 250 + 1 = 251。
        // 注意结果与段数无关——这正是它能扛住裁剪的原因。
        string withBase = "序号基准: 100\n" + MakeLogContent(2) + "序号基准: 250\n";
        Check("基准取最后一条且与段数无关（得 250 + 1 = 251）",
            LogService.NextSessionNumber(withBase) == 251,
            $"得到 {LogService.NextSessionNumber(withBase)}");

        // 同上内容的段数变了，序号必须不变——这是"抗裁剪"的核心证据
        string withBaseMore = "序号基准: 100\n" + MakeLogContent(9) + "序号基准: 250\n";
        Check("基准相同时，段数从 2 增到 9 也不影响序号",
            LogService.NextSessionNumber(withBaseMore) == 251,
            $"得到 {LogService.NextSessionNumber(withBaseMore)}");

        // 基准行损坏时退化为"段数 + 1"，不得把序号重置成怪值
        string broken = MakeLogContent(2) + "序号基准: 不是数字\n";
        Check("基准行损坏时退化为段数 + 1（得 3）",
            LogService.NextSessionNumber(broken) == 3,
            $"得到 {LogService.NextSessionNumber(broken)}");
    }

    // ---------------------------------------------------------------

    /// <summary>便捷重载：不需要附加说明时省略 detail</summary>
    private static void Check(string name, bool ok)
    {
        Check(name, ok, null);
    }

    private static void Check(string name, bool ok, string detail)
    {
        if (ok)
        {
            _passed++;
            Console.WriteLine($"  [通过] {name}");
        }
        else
        {
            _failed++;
            Console.WriteLine($"  [失败] {name}");
            if (!string.IsNullOrEmpty(detail))
                Console.WriteLine($"         {detail}");
        }
    }
}
