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
