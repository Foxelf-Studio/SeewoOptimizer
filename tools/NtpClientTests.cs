using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
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

        // --- 定时关机调度（2026-10-10 新增功能）---
        TestShutdownDayMaskBitOrder();
        TestShutdownPackedRoundTrip();
        TestShutdownPackedSurvivesGarbage();
        TestShutdownMatchesExactMinute();
        TestShutdownEveryDayRule();
        TestShutdownEmptyMaskNeverMatches();
        TestShutdownNextOccurrence();
        TestShutdownWarnWindow();
        TestShutdownLateWarnSlipsDeliberately();
        TestShutdownCountdownIsImmediate();
        TestShutdownUsesAbsolutePath();
        TestShutdownDecidePriority();
        TestShutdownPromiseNormalPath();
        TestShutdownPromiseAfterRestart();
        TestShutdownPromiseSkipDelayed();
        TestShutdownNoPromiseNoShutdown();
        TestShutdownWarnsOncePerWindow();
        TestShutdownWarnsAgainNextDay();
        TestShutdownSkipResumesNextDay();
        // 补执行与状态持久化：覆盖“休眠/重启跨过关机点”的漏关场景
        TestShutdownCatchUpWithinGrace();
        TestShutdownCatchUpBeyondGrace();
        TestShutdownStateRoundTrip();
        TestShutdownStateKeepsSkip();
        TestShutdownStateKeepsDelayedSkip();
        TestShutdownStateExpiresStalePromise();
        TestShutdownStateRejectsGarbage();
        // 【顺序约束】TestSkipAbortDecision 会永久置位进程内的 _shutdownIssued，
        // 一旦置位，之后任何 SkipOnce 都会真的去调 shutdown /a。
        // 因此它必须排在所有会调用 SkipOnce 的测试之后
        // （TestShutdownPromiseSkipDelayed / TestShutdownSkipResumesNextDay），
        // 否则那些测试会对宿主机产生系统级副作用。
        TestSkipAbortDecision();
        TestExitCodeFormatting();
        TestShutdownSerializationRoundTrip();
        TestShutdownSerializationToleratesGarbage();

        // --- 守护天数（"已守护 x 天"的计数契约）---
        TestGuardedDaysFirstDayIsZero();
        TestGuardedDaysCrossesMidnight();
        TestGuardedDaysUnrecordedFallsBackToZero();
        TestGuardedDaysNeverGoesBackwards();
        TestGuardedDaysIgnoresTimeOfDay();
        TestFirstRunDateTrust();

        // --- 升级迁移（老用户补勾"启动时自动执行"，新用户保持默认）---
        TestAutoStartUpgradeIdempotent();
        TestAutoStartUpgradeForLegacyUser();
        TestAutoStartUpgradeForFreshInstall();

        // --- 设置规整（注册表可手改，界面限不住）---
        TestSettingsNormalizeClampsRuleCount();

        // --- 签名信任链（换钥过渡）---
        TestSignatureVerifierAcceptsSecondKey();

        // --- 多更新源选源（Gitee / GitHub）---
        TestPickNewerPrefersHigherVersion();
        TestPickNewerPrefersDomesticOnTie();
        TestPickNewerToleratesMissingSource();
        TestBuildDownloadUrlsOnlySameVersion();
        TestVersionMatchesIgnoresMissingRevision();
        TestVersionMatchesRejectsOlderBuild();
        TestGiteeReleaseJsonParsesWithoutSize();

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
    // 定时关机调度（2026-10-10 新增功能：到点提醒 + 关机）
    //
    // 【测什么】ShutdownScheduleLogic / ShutdownTimeRule 的纯逻辑部分：
    // 星期掩码匹配、时分匹配、下一次触发时刻、提醒窗口、优先级。
    // 全是纯函数，直接链接产品源码，不碰 UI、不读系统时钟。
    //
    // 【为什么必须测】关机是**破坏性动作**。逻辑错一格的后果是
    // "该关的日子没关"或"不该关的时刻把电脑关了"——后者会直接切掉
    // 教室一体机上正在讲的内容。这类错误不能靠人工等到那个点试出来。
    //
    // 【历史坑位，测试要钉住的】
    //   · DayOfWeek 枚举是 Sunday=0，而位序约定是 Monday=0。
    //     直接 (int)DayOfWeek 移位会整体错开一位 → 设周一却在周二执行。
    //   · 掩码为 0 时规则永远不命中。这类"静默失效"最难排查。
    //   · 提醒与关机若各算各的时刻，会慢慢错开 → 提醒了却不关，或关了没提醒。
    // ---------------------------------------------------------------

    private static ShutdownTimeRule MakeRule(int hour, int minute, params DayOfWeek[] days)
    {
        var rule = new ShutdownTimeRule { Hour = hour, Minute = minute };
        foreach (DayOfWeek d in days) rule.SetDay(d, true);
        return rule;
    }

    /// <summary>
    /// 位序契约：位 0 必须对应周一、位 6 对应周日。
    ///
    /// 这条是最容易错、也最难发现的一条：DayOfWeek 枚举是 Sunday=0，
    /// 若产品代码图省事直接 (int)day 移位，掩码整体错开一位，
    /// 表现为"设了周一却在周二关机"。这种偏差每天只差一天，肉眼极难察觉。
    /// </summary>
    private static void TestShutdownDayMaskBitOrder()
    {
        Check("周一对应位 0",
            ShutdownTimeRule.BitForDayOfWeek(DayOfWeek.Monday) == 0,
            $"得到位 {ShutdownTimeRule.BitForDayOfWeek(DayOfWeek.Monday)}");

        Check("周日对应位 6（不得受 DayOfWeek=0 影响）",
            ShutdownTimeRule.BitForDayOfWeek(DayOfWeek.Sunday) == 6,
            $"得到位 {ShutdownTimeRule.BitForDayOfWeek(DayOfWeek.Sunday)}");

        // 七个位必须两两不同，且恰好覆盖 0..6
        var bits = new System.Collections.Generic.HashSet<int>();
        for (int i = 0; i < 7; i++)
            bits.Add(ShutdownTimeRule.BitForDayOfWeek(ShutdownTimeRule.DayOfWeekForBit(i)));

        Check("七个位互不重复且落在 0..6", bits.Count == 7,
            $"得到 {bits.Count} 个不同的位");

        // 逐天验证：只勾周一，则只有周一生效
        var onlyMonday = MakeRule(10, 0, DayOfWeek.Monday);
        Check("只勾周一：周一生效",
            onlyMonday.IsDayEnabled(DayOfWeek.Monday));
        Check("只勾周一：周二不生效",
            !onlyMonday.IsDayEnabled(DayOfWeek.Tuesday));

        // 逐天验证全部七天（防止只有首尾两天碰巧对）
        bool allCorrect = true;
        string detail = null;
        foreach (DayOfWeek d in Enum.GetValues(typeof(DayOfWeek)))
        {
            var r = MakeRule(0, 0, d);
            if (!r.IsDayEnabled(d))
            {
                allCorrect = false;
                detail = $"{d} 单独设置后却判定为不生效";
                break;
            }
            // 其余六天必须都不生效
            foreach (DayOfWeek other in Enum.GetValues(typeof(DayOfWeek)))
            {
                if (other == d) continue;
                if (r.IsDayEnabled(other))
                {
                    allCorrect = false;
                    detail = $"只设 {d}，却 {other} 也生效";
                    break;
                }
            }
            if (!allCorrect) break;
        }
        Check("逐天验证：任意单天设置只对该天生效", allCorrect, detail);
    }

    /// <summary>打包往返：ToPacked → FromPacked 必须完全还原</summary>
    private static void TestShutdownPackedRoundTrip()
    {
        bool allOk = true;
        string detail = null;

        // 覆盖边界：0:00 / 23:59、掩码 0 / 全选 / 单日
        int[][] cases = new int[][]
        {
            new int[] { 0, 0, 0 },
            new int[] { 23, 59, ShutdownTimeRule.EveryDayMask },
            new int[] { 12, 30, 1 },          // 只有周一
            new int[] { 7, 5, 127 },          // 每天
            new int[] { 18, 0, 0b0101010 },
            new int[] { 6, 45, 64 },          // 只有周日（位 6）
        };

        foreach (int[] c in cases)
        {
            var rule = new ShutdownTimeRule { Hour = c[0], Minute = c[1], DayMask = c[2] };
            ShutdownTimeRule back = ShutdownTimeRule.FromPacked(rule.ToPacked());

            if (back.Hour != c[0] || back.Minute != c[1] || back.DayMask != c[2])
            {
                allOk = false;
                detail = $"原 {c[0]}:{c[1]} mask={c[2]} → 还原为 "
                       + $"{back.Hour}:{back.Minute} mask={back.DayMask}";
                break;
            }
        }
        Check("打包往返：时间与掩码完全还原", allOk, detail);

        // 全枚举校验编码是单射的（不同规则不得编码成同一个值）。
        // 【为什么必须查这一条】编码用 "分钟数 * 128 + 掩码"，
        // 若乘数写成了 127，掩码 127 就会与下一档的掩码 0 撞上，
        // 表现为"某两条规则互相覆盖"，而单一的往返测试抓不到。
        var seen = new System.Collections.Generic.Dictionary<int, string>();
        string collision = null;
        for (int h = 0; h < 24 && collision == null; h++)
        {
            for (int m = 0; m < 60 && collision == null; m++)
            {
                for (int mask = 0; mask <= ShutdownTimeRule.EveryDayMask; mask++)
                {
                    var r = new ShutdownTimeRule { Hour = h, Minute = m, DayMask = mask };
                    int packed = r.ToPacked();
                    string key = $"{h}:{m}/{mask}";

                    if (seen.ContainsKey(packed))
                    {
                        collision = $"{seen[packed]} 与 {key} 都编码为 {packed}";
                        break;
                    }
                    seen[packed] = key;
                }
            }
        }
        Check("编码是单射的：任意两条不同规则不得编成同一个值（乘数必须是 128）",
            collision == null, collision);
    }

    /// <summary>损坏的注册表值不得产生非法时刻，也不得抛异常</summary>
    private static void TestShutdownPackedSurvivesGarbage()
    {
        int[] garbage = { -1, -999999, 0, int.MaxValue, int.MinValue, 999999 };
        bool allOk = true;
        string detail = null;

        foreach (int g in garbage)
        {
            try
            {
                ShutdownTimeRule r = ShutdownTimeRule.FromPacked(g);

                if (r.Hour < 0 || r.Hour > 23 || r.Minute < 0 || r.Minute > 59)
                {
                    allOk = false;
                    detail = $"输入 {g} 得到非法时刻 {r.Hour}:{r.Minute}";
                    break;
                }
                if ((r.DayMask & ~ShutdownTimeRule.EveryDayMask) != 0)
                {
                    allOk = false;
                    detail = $"输入 {g} 得到越界掩码 {r.DayMask}";
                    break;
                }
            }
            catch (Exception ex)
            {
                allOk = false;
                detail = $"输入 {g} 抛出了 {ex.GetType().Name}";
                break;
            }
        }
        Check("损坏的编码不产生非法时刻、不抛异常", allOk, detail);
    }

    /// <summary>
    /// 时分匹配：必须**精确到分钟**，同一分钟内任何秒数都算命中，
    /// 差一分钟都不算。
    /// </summary>
    private static void TestShutdownMatchesExactMinute()
    {
        var rule = MakeRule(23, 30, DayOfWeek.Monday);
        DateTime monday = new DateTime(2026, 10, 12, 23, 30, 0);   // 2026-10-12 是周一

        Check("前提：2026-10-12 是周一",
            monday.DayOfWeek == DayOfWeek.Monday, $"实际 {monday.DayOfWeek}");

        Check("23:30:00 命中", rule.Matches(monday));
        Check("23:30:59 同样命中（按分钟粒度）",
            rule.Matches(monday.AddSeconds(59)));
        Check("23:29:59 不命中（早一分钟）",
            !rule.Matches(monday.AddSeconds(-1)));
        Check("23:31:00 不命中（晚一分钟）",
            !rule.Matches(monday.AddMinutes(1)));
        Check("同一天 22:30 不命中（小时不同）",
            !rule.Matches(new DateTime(2026, 10, 12, 22, 30, 0)));

        // 跨周：周二同一时刻不命中
        Check("周二同一时刻不命中（星期不匹配）",
            !rule.Matches(new DateTime(2026, 10, 13, 23, 30, 0)));
    }

    /// <summary>"每天"应当等价于七天全选，且必须是真的七天</summary>
    private static void TestShutdownEveryDayRule()
    {
        var rule = new ShutdownTimeRule { Hour = 8, Minute = 0, DayMask = ShutdownTimeRule.EveryDayMask };

        Check("每天模式：IsEveryDay 为真", rule.IsEveryDay);

        bool allSeven = true;
        string detail = null;
        foreach (DayOfWeek d in Enum.GetValues(typeof(DayOfWeek)))
        {
            if (!rule.IsDayEnabled(d))
            {
                allSeven = false;
                detail = $"{d} 未被覆盖";
                break;
            }
        }
        Check("每天模式：七天全部命中", allSeven, detail);

        // 逐日跑到一周，验证每天都能命中（用连续 7 天模拟）
        DateTime start = new DateTime(2026, 10, 12, 8, 0, 0);
        bool everyDayHit = true;
        for (int i = 0; i < 14; i++)
        {
            if (!rule.Matches(start.AddDays(i)))
            {
                everyDayHit = false;
                detail = $"{start.AddDays(i):yyyy-MM-dd}（{start.AddDays(i).DayOfWeek}）未命中";
                break;
            }
        }
        Check("每天模式：连续 14 天全部命中", everyDayHit, detail);
    }

    /// <summary>
    /// 空掩码：规则永远不命中，且 NextOccurrence 返回 null。
    ///
    /// 【为什么单独测】这是最典型的"静默失效"——用户设了时间点，
    /// 程序不报错、不提示，就是永远不关机。必须能自动发现。
    /// </summary>
    private static void TestShutdownEmptyMaskNeverMatches()
    {
        var empty = new ShutdownTimeRule { Hour = 12, Minute = 0, DayMask = 0 };

        bool never = true;
        string detail = null;
        DateTime start = new DateTime(2026, 10, 12, 12, 0, 0);
        for (int i = 0; i < 21; i++)
        {
            if (empty.Matches(start.AddDays(i)))
            {
                never = false;
                detail = $"{start.AddDays(i):yyyy-MM-dd} 竟然命中";
                break;
            }
        }
        Check("空掩码规则：三周内一次都不命中", never, detail);

        Check("空掩码规则：NextOccurrence 返回 null（不得死循环找下去）",
            !ShutdownScheduleLogic.NextOccurrence(empty, start).HasValue);

        Check("空掩码规则：IsEveryDay 为假", !empty.IsEveryDay);
    }

    /// <summary>
    /// 下一次触发：必须是 from **之后**的第一个命中时刻，
    /// 且要能正确跨天、跨周。
    /// </summary>
    private static void TestShutdownNextOccurrence()
    {
        var daily = MakeRule(23, 0, DayOfWeek.Monday, DayOfWeek.Tuesday,
                                        DayOfWeek.Wednesday, DayOfWeek.Thursday,
                                        DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday);

        DateTime noon = new DateTime(2026, 10, 12, 12, 0, 0);
        DateTime? next = ShutdownScheduleLogic.NextOccurrence(daily, noon);
        Check("每天 23:00：从中午算出当天 23:00",
            next.HasValue && next.Value == new DateTime(2026, 10, 12, 23, 0, 0),
            next.HasValue ? $"得到 {next.Value:yyyy-MM-dd HH:mm:ss}" : "返回 null");

        // 已过今天的点 → 应算到明天
        DateTime late = new DateTime(2026, 10, 12, 23, 30, 0);
        DateTime? next2 = ShutdownScheduleLogic.NextOccurrence(daily, late);
        Check("每天 23:00：从 23:30 算出次日 23:00",
            next2.HasValue && next2.Value == new DateTime(2026, 10, 13, 23, 0, 0),
            next2.HasValue ? $"得到 {next2.Value:yyyy-MM-dd HH:mm:ss}" : "返回 null");

        // 【关键边界】正好落在目标分钟上时，必须返回**下一次**而不是自身。
        // 若返回自身，调用方会据此排出"此刻提醒、5 分钟后关机"，
        // 但那条规则其实刚刚已经触发过 —— 会重复关机。
        DateTime exactly = new DateTime(2026, 10, 12, 23, 0, 0);
        DateTime? next3 = ShutdownScheduleLogic.NextOccurrence(daily, exactly);
        Check("正好在目标分钟上：返回的是明天同一时刻，不是自身",
            next3.HasValue && next3.Value == new DateTime(2026, 10, 13, 23, 0, 0),
            next3.HasValue ? $"得到 {next3.Value:yyyy-MM-dd HH:mm:ss}" : "返回 null");

        // 跨周：只设周一，从周二算起应落到下周一（跨 6 天）
        var onlyMonday = MakeRule(9, 0, DayOfWeek.Monday);
        DateTime tue = new DateTime(2026, 10, 13, 10, 0, 0);
        DateTime? next4 = ShutdownScheduleLogic.NextOccurrence(onlyMonday, tue);
        Check("只设周一：从周二算出下周一 09:00（跨周）",
            next4.HasValue && next4.Value == new DateTime(2026, 10, 19, 9, 0, 0),
            next4.HasValue ? $"得到 {next4.Value:yyyy-MM-dd}（{next4.Value.DayOfWeek}）" : "返回 null");
    }

    /// <summary>
    /// 提醒窗口：关机前 5 分钟**整段**（[触发-5min, 触发)）都必须命中提醒。
    ///
    /// 【为什么窗口是一段而不是一分钟】
    /// 原先只认"提前量那一分钟"（23:00 的规则只认 22:55 这一分钟），
    /// 因为正常运行下 22:55 一定会被 20 秒轮询捕到。但重启会打破这个假设：
    /// 22:56 重启后 22:55 早已过去，若只认那一分钟，这次提醒永远补不回来。
    /// 放宽成一段后，22:56~22:59 之间重启都能补弹，用户仍能选"本次不关机"。
    /// </summary>
    private static void TestShutdownWarnWindow()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        DateTime shutdownAt = new DateTime(2026, 10, 12, 23, 0, 0);

        // 提前量必须是 5 分钟（写在产品常量里，测试引用它而不是抄字面量）
        Check($"提前量常量为 5 分钟（实际 {ShutdownScheduleLogic.WarnMinutesAhead}）",
            ShutdownScheduleLogic.WarnMinutesAhead == 5);

        DateTime warnAt = shutdownAt.AddMinutes(-ShutdownScheduleLogic.WarnMinutesAhead);

        Check("关机前 5 分钟：命中提醒",
            ShutdownScheduleLogic.FindRuleToWarn(rules, warnAt) != null,
            $"检查时刻 {warnAt:HH:mm}");

        Check("关机前 5 分钟那一整分钟内都命中（含第 59 秒）",
            ShutdownScheduleLogic.FindRuleToWarn(rules, warnAt.AddSeconds(59)) != null);

        // 窗口是一整段：22:56 / 22:57 / 22:58 / 22:59 全部命中。
        // 这正是"重启后补弹"赖以成立的前提——如果这里只认 22:55，
        // 22:56 重启的机器就永远收不到提醒，会静默地准点关机。
        for (int offset = 1; offset <= 4; offset++)
        {
            DateTime probe = warnAt.AddMinutes(offset);
            Check($"关机前 {5 - offset} 分钟（{probe:HH:mm}）：仍在窗口内，命中",
                ShutdownScheduleLogic.FindRuleToWarn(rules, probe) != null,
                $"检查时刻 {probe:HH:mm}");
        }

        Check("关机前 6 分钟：不提醒（窗口尚未开始）",
            ShutdownScheduleLogic.FindRuleToWarn(rules, warnAt.AddMinutes(-1)) == null);

        Check("关机时刻本身：不触发提醒（该走关机分支）",
            ShutdownScheduleLogic.FindRuleToWarn(rules, shutdownAt) == null);
    }

    /// <summary>
    /// 在提醒窗口里才被发现时，关机时刻会**顺延**——这是被明确批准的取舍。
    ///
    /// 【为什么要在测试里把它钉成"预期行为"】
    /// 2026-10-11 实机日志：规则配的是 08:50，实际 08:51:15 才关机。
    /// 看上去像 bug，其实是承诺时刻的公式取向所致：
    ///     promised = max(发现时刻所在分钟 + 提前量, 规则时刻)
    /// 即"**保证给满提前量**"，规则时刻只是下限。
    ///
    /// 触发条件：程序在 5 分钟窗口内才启动，或改设置重启了巡检
    /// （WinForms Timer 首次触发要等满一个轮询间隔，本例 20 秒）。
    /// 最坏可顺延到 规则时刻 + (提前量 - 1) 分钟。
    ///
    /// 【用户决定：维持现状（2026-10-11）】
    /// 提醒框是**唯一**能点"本次不关机"的机会。若不顺延，窗口内才启动的机器
    /// 会弹出一个一闪而过的框，等于没有拦截机会；而对话框上写的就是顺延后的
    /// 真实时刻，用户被如实告知，不存在"悄悄晚了"。
    ///
    /// 所以这个测试的用意与常见相反：**它保护的是"不要被修掉"**。
    /// 将来若有人把顺延当 bug 改掉，这里会变红，逼他重新面对那个取舍
    /// 并重新征求用户决定，而不是默默改掉。
    /// </summary>
    private static void TestShutdownLateWarnSlipsDeliberately()
    {
        // 2026-10-12 是周一
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(8, 50, DayOfWeek.Monday)
        };
        DateTime ruleTime = new DateTime(2026, 10, 12, 8, 50, 0);

        // 窗口是 [08:45, 08:50)。这里模拟"程序在 08:46 才第一次巡检到"。
        ShutdownService.ResetRuntimeState();
        ShutdownService.PollResult late =
            ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 8, 46, 0));

        Check("窗口内才被发现（08:46）时仍然弹提醒",
            late.ShouldWarn,
            late.ShouldWarn ? "" : "不弹提醒会让这台机器静默准点关机，用户毫无拦截机会");

        Check("承诺时刻顺延为 08:51（发现时刻 + 5 分钟），而非规则时刻 08:50",
            late.DueAt == new DateTime(2026, 10, 12, 8, 51, 0),
            $"得到 {late.DueAt:HH:mm}——这里就该晚于规则时刻，"
          + "改掉它等于取消『保证给满 5 分钟提醒』的取舍");

        Check("顺延后的时刻确实晚于用户配置的 08:50",
            late.DueAt > ruleTime,
            $"得到 {late.DueAt:HH:mm} 与规则 {ruleTime:HH:mm}");

        // 反向契约：窗口一开始就被发现时**不**顺延。
        // 有这一条才能证明上面测的是"顺延"而不是"公式整体偏移了"。
        ShutdownService.ResetRuntimeState();
        ShutdownService.PollResult onTime =
            ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 8, 45, 0));

        Check("窗口一开始就被发现（08:45）时承诺恰为准点 08:50（不顺延）",
            onTime.ShouldWarn && onTime.DueAt == ruleTime,
            $"得到 {(onTime.ShouldWarn ? onTime.DueAt.ToString("HH:mm") : "未提醒")}");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 到点执行时**不再带系统倒计时**（该常量必须是 0）。
    ///
    /// 【为什么这条要单独钉住】这是用户明确要求的行为变更，而且方向
    /// 与"直觉上的安全做法"相反——直觉会认为"留 5 分钟撤销窗口更安全"。
    /// 实机证明恰恰相反：5 分钟倒计时结束后 Windows 会自己弹出
    /// "您将要被注销"的系统框，与我们的提醒框叠在一起；而那 5 分钟里
    /// 用户唯一的撤销手段是敲命令行 shutdown /a，教室一体机上没人会做。
    ///
    /// 该设计下"提前 5 分钟的提醒框"就是唯一的确认窗口：
    /// 框里问过，用户不点"本次不关机"就到点直接关。
    ///
    /// 若将来有人把这个值改回 300，本测试会立刻变红，
    /// 逼他重新面对上面这段权衡——而不是默默把系统框又引回来。
    /// </summary>
    private static void TestShutdownCountdownIsImmediate()
    {
        Check($"系统倒计时为 0（到点立即关机，实际 {ShutdownService.SystemCountdownSeconds}）",
            ShutdownService.SystemCountdownSeconds == 0,
            $"得到 {ShutdownService.SystemCountdownSeconds}——"
          + "非 0 会引入系统自带的'您将要被注销'框，与提醒框叠加");

        // 提醒提前量仍必须是 5 分钟：它是唯一的确认窗口，不能被改小成 0
        Check($"提醒提前量仍为 5 分钟（唯一确认窗口，实际 {ShutdownScheduleLogic.WarnMinutesAhead}）",
            ShutdownScheduleLogic.WarnMinutesAhead == 5,
            $"得到 {ShutdownScheduleLogic.WarnMinutesAhead}");
    }

    /// <summary>
    /// 关机命令必须用 shutdown.exe 的**绝对路径**启动，且工作目录被钉死。
    ///
    /// 【为什么单独钉这一条——实机报错 0xc0000142 的直接对策】
    /// 实机截图出现 `shutdown.exe - 应用程序错误 / 应用程序无法正常启动(0xc0000142)`。
    /// 0xc0000142 是 STATUS_DLL_INIT_FAILED，含义是**进程自己没起来**，
    /// 而非"关机被程序挡住"。
    ///
    /// 根因是裸文件名 "shutdown.exe" 走 CreateProcess 的搜索顺序
    /// （当前目录 → PATH）。教室机器 PATH 常被工具塞进奇怪条目，
    /// 一旦有同名/损坏副本先被命中，就会 0xc0000142。
    ///
    /// 修法是改用 %SystemRoot%\System32\shutdown.exe 绝对路径，
    /// 并显式设置 WorkingDirectory，不继承本程序可能异常的工作目录。
    ///
    /// 本测试若变红（比如有人改回裸名），说明上述回归被重新引入。
    /// </summary>
    private static void TestShutdownUsesAbsolutePath()
    {
        string exe = ShutdownService.ResolveShutdownExe();

        // 1. 必须是绝对路径，不能是裸名
        Check($"shutdown.exe 使用绝对路径（实际 {exe}）",
            !string.IsNullOrEmpty(exe) && System.IO.Path.IsPathRooted(exe),
            "裸文件名会走 PATH/当前目录搜索，是 0xc0000142 的根源");

        // 2. 扩展名必须是 shutdown.exe（大小写不敏感）
        Check($"路径末端是 shutdown.exe（实际 {System.IO.Path.GetFileName(exe)}）",
            string.Equals(System.IO.Path.GetFileName(exe), "shutdown.exe",
                StringComparison.OrdinalIgnoreCase));

        // 3. 该路径在当前机器上确实存在（若为 Sysnative 回退也算）
        Check($"解析出的 shutdown.exe 存在（{exe}）",
            System.IO.File.Exists(exe));

        // 4. 通过 BuildShutdownStartInfo 产出的对象同样满足以上契约，
        //    且工作目录非空——确保两处调用不会各自退化
        var psi = ShutdownService.BuildShutdownStartInfo("/a");
        Check("BuildShutdownStartInfo 用绝对路径",
            System.IO.Path.IsPathRooted(psi.FileName));
        Check("BuildShutdownStartInfo 钉死了工作目录",
            !string.IsNullOrEmpty(psi.WorkingDirectory)
            && System.IO.Path.IsPathRooted(psi.WorkingDirectory));
        Check("BuildShutdownStartInfo 参数原样透传",
            psi.Arguments == "/a",
            $"得到 {psi.Arguments}");
    }

    /// <summary>
    /// "本次不关机"到底该不该去撤销系统关机。
    ///
    /// 【实机日志提出的问题（2026-10-11）】
    /// 用户在提醒框里点了"本次不关机"，日志却写着
    /// 「撤销关机未生效（shutdown /a 退出码 1116，通常表示当前没有进行中的关机）」，
    /// 用户读到"未生效"便怀疑是不是没取消成功。
    ///
    /// 查下来是**日志措辞与实际动作不符**：/t 0 的设计下，提醒阶段根本不会
    /// 下发任何关机命令（见 SystemCountdownSeconds 的注释），命令只在到点才发。
    /// 所以"没有需要撤销的系统关机"是正常且正确的——机器确实没关。
    ///
    /// 修法：SkipOnce 先看 <see cref="ShutdownService.ShouldAbortOnSkip"/>，
    /// 没下发过就只记账、不调 shutdown /a，日志也直说"无需撤销"。
    ///
    /// 【为什么本测试不真的走 SkipOnce 的"已下发"分支】
    /// 那条分支会真的启动 shutdown.exe /a。单元测试不该对宿主机产生
    /// 系统级副作用（万一测试机上恰好有进行中的关机，就被它撤销了）。
    /// 因此这里只断言**决策**（ShouldAbortOnSkip），
    /// 决策与动作之间的连线由 SkipOnce 里那一行直白代码保证。
    ///
    /// 【顺序约束】NoteShutdownIssued 会永久置位进程内的静态标志，
    /// 所以两个方向都写在同一个测试里。若将来新增依赖
    /// "未下发过"这一前提的断言，必须排在调用本测试之前。
    /// </summary>
    private static void TestSkipAbortDecision()
    {
        DateTime due = new DateTime(2026, 10, 12, 23, 40, 0);

        // 方向一：从未下发过关机 → 跳过时不该去撤销
        Check("未下发过关机时：ShouldAbortOnSkip 为 false",
            !ShutdownService.ShouldAbortOnSkip(),
            "为 true 会让每次跳过都白调一次 shutdown /a，并在日志里留下"
          + "让人误以为失败的『撤销未生效』");

        Check("未下发过关机时：SkipOnce 只记账、不撤销",
            ShutdownService.SkipOnce(due) == ShutdownService.SkipOutcome.MarkedOnly,
            "MarkedOnly 表示没有去动系统侧的关机；用户看到的是『本次跳过已生效』");

        // 方向二：已经下发过 → 必须去撤销（真实时序：到点与点按钮撞在一起）
        ShutdownService.NoteShutdownIssued();
        Check("已下发过关机时：ShouldAbortOnSkip 为 true",
            ShutdownService.ShouldAbortOnSkip(),
            "为 false 会让『本次不关机』挡不住已发出的关机命令——"
          + "用户以为取消了，机器还是关了");
    }

    /// <summary>
    /// 退出码的显示格式。
    ///
    /// 【为什么值得单独钉】实机日志里出现过 -1073741502，
    /// 十进制看不出任何含义；它其实是 0xC0000142（DLL 初始化失败），
    /// 与上一批修的 shutdown.exe 启动失败是同一个系统状态码。
    /// 负数按十六进制打印，排查时才能一眼对上号。
    /// </summary>
    private static void TestExitCodeFormatting()
    {
        Check("负数退出码按十六进制显示（-1073741502 → 0xC0000142）",
            ShutdownService.DescribeExitCode(-1073741502) == "0xC0000142",
            $"得到 {ShutdownService.DescribeExitCode(-1073741502)}");

        Check("正数退出码保持十进制（1116 = 没有进行中的关机）",
            ShutdownService.DescribeExitCode(1116) == "1116",
            $"得到 {ShutdownService.DescribeExitCode(1116)}");

        Check("零退出码原样（0 = 撤销成功）",
            ShutdownService.DescribeExitCode(0) == "0",
            $"得到 {ShutdownService.DescribeExitCode(0)}");
    }

    /// <summary>
    /// 优先级契约：关机与提醒落在同一分钟时，必须选择**关机**。
    ///
    /// 【为什么重要】反过来的话，程序会在该关机的时刻弹出一个提醒框
    /// 把关机顶掉——而那个框还写着"5 分钟后关机"，语义完全错乱。
    /// </summary>
    private static void TestShutdownDecidePriority()
    {
        ShutdownTimeRule hit;

        // 构造：规则 A 在 23:00 关机，规则 B 在 23:05 关机。
        // 那么 23:00 这一刻既是 A 的关机点，又是 B 的提醒点。
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday),
            MakeRule(23, 5, DayOfWeek.Monday)
        };

        DateTime conflict = new DateTime(2026, 10, 12, 23, 0, 0);

        // 先确认这确实是个冲突时刻（B 的提醒点 = 23:05 - 5 = 23:00）
        Check("前提：23:00 同时是规则 B 的提醒点",
            ShutdownScheduleLogic.FindRuleToWarn(rules, conflict) != null);

        ShutdownScheduleLogic.ActionKind kind =
            ShutdownScheduleLogic.Decide(rules, conflict, out hit);

        Check("冲突时刻选择「关机」而非「提醒」",
            kind == ShutdownScheduleLogic.ActionKind.Shutdown,
            $"得到 {kind}");

        Check("冲突时刻命中的是 23:00 那条规则",
            hit != null && hit.Hour == 23 && hit.Minute == 0,
            hit == null ? "未命中任何规则" : $"命中 {hit.Describe()}");

        // 平常时刻什么也不做
        DateTime idle = new DateTime(2026, 10, 12, 15, 0, 0);
        ShutdownScheduleLogic.ActionKind idleKind =
            ShutdownScheduleLogic.Decide(rules, idle, out hit);
        Check("无关时刻返回 None", idleKind == ShutdownScheduleLogic.ActionKind.None,
            $"得到 {idleKind}");
    }

    /// <summary>
    /// 承诺时刻机制：正常路径下，22:55 弹提醒时承诺的应当是准点 23:00，
    /// 而 23:00 那一刻必须命中关机。
    ///
    /// 【为什么必须测】这是"不弹二次确认框"能成立的前提——
    /// 删掉二次确认后，唯一把关机触发出去的就是这条承诺链。
    /// 如果 22:55 承诺的不是 23:00，或 23:00 不认这个承诺，关机就再也不会发生。
    /// </summary>
    private static void TestShutdownPromiseNormalPath()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        DateTime warnMoment = new DateTime(2026, 10, 12, 22, 55, 0);

        ShutdownService.ResetRuntimeState();

        // 22:55 这一刻巡检：正常路径下承诺时刻应当恰好等于准点 23:00。
        // 同样从 Poll 返回值读，不自己重算。
        var warn = ShutdownService.Poll(rules, warnMoment);

        Check("22:55 命中提醒",
            warn.ShouldWarn, warn.ShouldWarn ? "" : "未命中提醒");

        Check("22:55 弹提醒时承诺的时刻是准点 23:00（正常路径不顺延）",
            warn.DueAt == new DateTime(2026, 10, 12, 23, 0, 0),
            $"得到 {warn.DueAt:yyyy-MM-dd HH:mm}");

        // 在 23:00 那一刻，承诺应当命中
        ShutdownService.MarkWarned(warnMoment);
        ShutdownService.PromiseShutdown(warn.DueAt);
        var r = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("23:00 命中关机（承诺已生效）",
            r.ShouldShutdown, r.ShouldShutdown ? "" : "未命中关机");
        Check("23:00 不再弹提醒",
            !r.ShouldWarn, r.ShouldWarn ? "错误地弹了提醒" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 重启补弹：22:56 重启后（提醒状态已丢），在 22:56 这一刻应当
    /// **重新命中提醒**，且承诺时刻**顺延为 23:01**（现在 + 5 分钟）。
    ///
    /// 【为什么这是本轮最关键的一条】
    /// 用户明确要求：22:56 重启时应重新弹框、重新计时、仍可选"本次不关机"。
    /// 同时原规则的 23:00 必须被顶替掉——否则 23:00 与 23:01 会各关一次。
    /// 本测试同时钉住这两面。
    /// </summary>
    private static void TestShutdownPromiseAfterRestart()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        DateTime restartMoment = new DateTime(2026, 10, 12, 22, 56, 0);

        // 重启：所有内存记账归零（这正是重启丢状态的效果）
        ShutdownService.ResetRuntimeState();

        // 22:56 这一刻巡检，应当**直接产出提醒**，且 DueAt 就是顺延后的承诺时刻。
        // 【关键】必须从 Poll 的返回值读承诺时刻，不能自己重算一遍——
        // 自己重算等于把产品逻辑抄进测试，产品改了测试也不会红。
        var atRestart = ShutdownService.Poll(rules, restartMoment);

        Check("22:56（重启后）应当命中提醒",
            atRestart.ShouldWarn,
            atRestart.ShouldWarn ? "" : "未命中提醒 → 重启后用户收不到提示");

        Check("22:56 补弹时承诺时刻顺延为 23:01（重新计时，非准点 23:00）",
            atRestart.DueAt == new DateTime(2026, 10, 12, 23, 1, 0),
            $"得到 {atRestart.DueAt:yyyy-MM-dd HH:mm}");

        // 模拟窗体弹框：把承诺时刻写入（这是窗体的职责）
        ShutdownService.MarkWarned(restartMoment);
        ShutdownService.PromiseShutdown(atRestart.DueAt);

        // 关键：原规则的 23:00 必须**不**触发关机（被顺延承诺顶替）
        var atOriginal = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("原规则时刻 23:00 不再关机（已被顺延顶替，避免关两次）",
            !atOriginal.ShouldShutdown,
            atOriginal.ShouldShutdown ? "23:00 仍然触发了关机 → 会关两次" : "");

        // 顺延后的 23:01 才关门
        var atDelayed = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 1, 0));
        Check("顺延时刻 23:01 命中关机",
            atDelayed.ShouldShutdown,
            atDelayed.ShouldShutdown ? "" : "23:01 未命中关机 → 会漏关");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 重启补弹后选"本次不关机"，则顺延的那次也不得执行。
    /// </summary>
    private static void TestShutdownPromiseSkipDelayed()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        DateTime delayed = new DateTime(2026, 10, 12, 23, 1, 0);

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(delayed);
        ShutdownService.SkipOnce(delayed);

        var r = ShutdownService.Poll(rules, delayed);
        Check("顺延时刻选了「本次不关机」后不执行关机",
            !r.ShouldShutdown, r.ShouldShutdown ? "仍然关机了" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 没有承诺时不关机：仅仅"到点是某个规则的时刻"不足以触发。
    ///
    /// 【为什么】这是行为变更的核心——旧版每轮从规则实时推导，
    /// 只要时刻对上就关。新版改由"承诺时刻"驱动，没有承诺就不该动作。
    /// 若不测这条，将来有人把判定改回"实时推导"也不会有测试报警，
    /// 而那个改法会让重启后 23:00 与 23:01 各关一次。
    /// </summary>
    private static void TestShutdownNoPromiseNoShutdown()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();   // 清空承诺

        var r = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("无承诺时，即使到点也不关机",
            !r.ShouldShutdown, r.ShouldShutdown ? "无承诺却关机了" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 整个提醒窗口内只弹一次提醒——用户不点按钮也不能每分钟重弹。
    ///
    /// 【这是实机反馈直接暴露的缺陷，务必钉死】
    /// 用户设置 23:40 关机，实机表现是 23:40 / 23:41 / 23:42 每分钟弹一个框。
    /// 根因：旧版去重记的是"已提醒的那一分钟"（`_lastWarnedMinute != 当前分钟`），
    /// 只能挡住同一分钟内的三跳（20s/40s/60s），挡不住跨分钟。
    /// 而提醒窗口本身是一整段 [23:35, 23:40)，共 5 分钟，
    /// 于是 23:35~23:39 每分钟各弹一次，越弹越晚（各带各的承诺时刻），
    /// 这正是截图里 23:40/23:41/23:42 三个框叠在一起的来源。
    ///
    /// 正确行为：同一个 DueAt 只弹一次；重启后内存态清空，才会重新弹。
    /// </summary>
    private static void TestShutdownWarnsOncePerWindow()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 40, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();

        // 窗口第一分钟：应当命中提醒
        var first = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 35, 0));
        Check("23:40 关机的提醒窗口第一分钟（23:35）命中提醒",
            first.ShouldWarn, first.ShouldWarn ? "" : "未命中提醒");

        Check("23:35 弹提醒时承诺的是准点 23:40",
            first.DueAt == new DateTime(2026, 10, 12, 23, 40, 0),
            $"得到 {first.DueAt:yyyy-MM-dd HH:mm}");

        // 后续每一分钟都不得再弹——这正是实机暴雷的地方
        for (int m = 36; m <= 39; m++)
        {
            var again = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, m, 0));
            Check($"窗口内 {m} 分不再重复弹提醒（同一轮关机只提醒一次）",
                !again.ShouldWarn,
                again.ShouldWarn
                    ? $"在 23:{m:D2} 又弹了一个（会叠加成多个框）——承诺被改写成 {again.DueAt:HH:mm}"
                    : "");
        }

        // 到点仍应关机：去重不能把主链路一起挡掉
        var due = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 40, 0));
        Check("窗口去重后，23:40 仍然正常触发关机",
            due.ShouldShutdown, due.ShouldShutdown ? "" : "去重把关机也挡掉了");

        // 重启（内存态归零）后，若仍在窗口内应能补弹——不能因为去重而永远不会再提醒
        ShutdownService.ResetRuntimeState();
        var afterRestart = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 37, 0));
        Check("重启后仍在窗口内可以重新弹提醒（去重不该永久生效）",
            afterRestart.ShouldWarn, afterRestart.ShouldWarn ? "" : "重启后收不到提醒");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 跨天必须重新提醒：今天的去重记录不能把明天的提醒一起吃掉。
    ///
    /// 【为什么单独测】_warnedForDue 记的是绝对时刻（含日期），
    /// 昨天的 23:40 与今天的 23:40 是两个不同值，所以自然会重新提醒。
    /// 但若将来有人图省事把它改存成"时分"（如 23:40 不带日期），
    /// 跨天就会永久静默——那么"每天 23:40 关机"从第二天起再也不提醒、也不关机。
    /// 这是一条极其隐蔽的失效路径，必须钉住。
    /// </summary>
    private static void TestShutdownWarnsAgainNextDay()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 40, DayOfWeek.Monday, DayOfWeek.Tuesday,
                     DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday,
                     DayOfWeek.Saturday, DayOfWeek.Sunday)
        };

        ShutdownService.ResetRuntimeState();

        // 第一天：10/12（周一）走完整轮
        var day1 = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 35, 0));
        Check("第一天 23:35 命中提醒（前提）",
            day1.ShouldWarn, day1.ShouldWarn ? "" : "第一天就没提醒");

        var day1Due = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 40, 0));
        Check("第一天 23:40 执行关机（前提）",
            day1Due.ShouldShutdown, day1Due.ShouldShutdown ? "" : "第一天没关机");

        // 第二天：10/13（周二）必须重新提醒、重新关机
        var day2 = ShutdownService.Poll(rules, new DateTime(2026, 10, 13, 23, 35, 0));
        Check("第二天 23:35 仍然命中提醒（去重记录不得跨天静默）",
            day2.ShouldWarn, day2.ShouldWarn ? "第二天被昨天的记录吃掉了，不会再提醒" : "");

        var day2Due = ShutdownService.Poll(rules, new DateTime(2026, 10, 13, 23, 40, 0));
        Check("第二天 23:40 仍然执行关机",
            day2Due.ShouldShutdown, day2Due.ShouldShutdown ? "第二天不再关机了" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 选了"本次不关机"后，同一天不得再问、也不得关机；
    /// 但第二天必须恢复正常提醒与关机。
    ///
    /// 【为什么必须测 skip 之后的跨天】_skipUntil 与 _warnedForDue 都是内存态。
    /// 若不验证"第二天恢复"，就可能出现"跳过一次后永远不再关机"——
    /// 那对教室一体机是致命的：以为设了定时关机，实际从此再没关过。
    /// </summary>
    private static void TestShutdownSkipResumesNextDay()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 40, DayOfWeek.Monday, DayOfWeek.Tuesday,
                     DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday,
                     DayOfWeek.Saturday, DayOfWeek.Sunday)
        };

        ShutdownService.ResetRuntimeState();

        // 第一天提醒后选"本次不关机"
        var warn = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 35, 0));
        Check("第一天 23:35 命中提醒（前提）",
            warn.ShouldWarn, warn.ShouldWarn ? "" : "第一天没提醒");

        ShutdownService.SkipOnce(warn.DueAt);

        // 同一天到点不关机
        var sameDay = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 40, 0));
        Check("选了「本次不关机」后，当天 23:40 不关机",
            !sameDay.ShouldShutdown, sameDay.ShouldShutdown ? "跳过了却仍然关机" : "");

        // 同一天窗口内也不再问
        var sameDayAgain = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 38, 0));
        Check("选了「本次不关机」后，当天窗口内不再重复询问",
            !sameDayAgain.ShouldWarn,
            sameDayAgain.ShouldWarn ? "跳过之后又弹了提醒" : "");

        // 第二天必须恢复：提醒 + 关机都要回来
        var nextDay = ShutdownService.Poll(rules, new DateTime(2026, 10, 13, 23, 35, 0));
        Check("第二天 23:35 恢复提醒（跳过只影响当天）",
            nextDay.ShouldWarn, nextDay.ShouldWarn ? "" : "跳过之后第二天不再提醒");

        var nextDayDue = ShutdownService.Poll(rules, new DateTime(2026, 10, 13, 23, 40, 0));
        Check("第二天 23:40 恢复关机（跳过只影响当天）",
            nextDayDue.ShouldShutdown, nextDayDue.ShouldShutdown ? "" : "跳过之后第二天不再关机");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 关机点之后、宽限期之内仍应补执行。
    ///
    /// 旧实现只认 `当前分钟 == 承诺时刻`：休眠唤醒、重启、UI 线程卡顿
    /// 只要跨过那一分钟，当天就再也不会关机。这是本组测试存在的唯一理由。
    /// </summary>
    private static void TestShutdownCatchUpWithinGrace()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(new DateTime(2026, 10, 12, 23, 0, 0));

        var late = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 4, 30));
        Check("关机点后 4 分 30 秒仍应补执行",
            late.ShouldShutdown,
            late.ShouldShutdown ? "" : "错过一分钟就整天不关");
        Check("补执行的目标时刻仍是原承诺 23:00",
            late.DueAt == new DateTime(2026, 10, 12, 23, 0, 0),
            $"得到 {late.DueAt:HH:mm}");

        ShutdownService.MarkShutdown(late.DueAt);
        var again = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 5, 0));
        Check("补执行后不得重复执行",
            !again.ShouldShutdown,
            again.ShouldShutdown ? "同一轮关机会每 20 秒重复下发" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 超过宽限就放弃：不能让上午开机的机器去补执行前一晚的关机。
    /// </summary>
    private static void TestShutdownCatchUpBeyondGrace()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(new DateTime(2026, 10, 12, 23, 0, 0));

        var tooLate = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 11, 0));
        Check("超过 10 分钟宽限不再补执行",
            !tooLate.ShouldShutdown,
            tooLate.ShouldShutdown ? "上午开机会补关前一晚的机器" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 状态持久化往返：承诺写盘、重新读回后仍能在宽限内补执行。
    /// </summary>
    private static void TestShutdownStateRoundTrip()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(new DateTime(2026, 10, 12, 23, 0, 0));
        string text = ShutdownService.ExportStateText();

        ShutdownService.ResetRuntimeState();
        string summary = ShutdownService.ImportStateText(text,
            new DateTime(2026, 10, 12, 22, 57, 0));

        Check("恢复状态应给出一句可读摘要",
            !string.IsNullOrEmpty(summary), "摘要为空，日志里无法判断恢复到了什么");

        var late = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 4, 0));
        Check("重启后恢复的承诺仍能在宽限内补执行",
            late.ShouldShutdown,
            late.ShouldShutdown ? "" : "落盘了却没用上");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// “本次不关机”必须跨重启保留：用户已经明确拒绝过，
    /// 重启后再问一次就是打扰，再关一次就是事故。
    /// </summary>
    private static void TestShutdownStateKeepsSkip()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(new DateTime(2026, 10, 12, 23, 0, 0));
        ShutdownService.SkipOnce(new DateTime(2026, 10, 12, 23, 0, 0));
        string text = ShutdownService.ExportStateText();

        ShutdownService.ResetRuntimeState();
        ShutdownService.ImportStateText(text, new DateTime(2026, 10, 12, 22, 57, 0));

        var r = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("重启后仍记得「本次不关机」",
            !r.ShouldShutdown,
            r.ShouldShutdown ? "用户说过不关，重启后还是关了" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 顺延承诺上的“本次不关机”也必须跨重启保留。
    ///
    /// 这是最容易漏的一条：规则 23:00、22:56 重启补弹、承诺 23:01。
    /// 用户跳过的是 23:01，而提醒窗口算的是 23:00。进程内靠 _warnedForDue
    /// 挡住重复询问；重启后它不落盘，只有 skip 被恢复，必须另有一道
    /// `_skipUntil != _promisedAt` 才能不重复问。
    /// </summary>
    private static void TestShutdownStateKeepsDelayedSkip()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();

        var warn = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 22, 56, 0));
        Check("22:56 重启补弹（前提）", warn.ShouldWarn,
            warn.ShouldWarn ? "" : "没弹提醒，后续断言无意义");
        Check("顺延承诺为 23:01（前提）",
            warn.DueAt == new DateTime(2026, 10, 12, 23, 1, 0),
            $"得到 {warn.DueAt:HH:mm}");

        ShutdownService.SkipOnce(warn.DueAt);
        string text = ShutdownService.ExportStateText();

        ShutdownService.ResetRuntimeState();
        ShutdownService.ImportStateText(text, new DateTime(2026, 10, 12, 22, 57, 0));

        var afterRestart = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 22, 57, 0));
        Check("重启后不得再次询问已跳过的顺延承诺",
            !afterRestart.ShouldWarn,
            afterRestart.ShouldWarn ? "用户刚说过不关，重启后又弹一次" : "");

        var atDue = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 1, 0));
        Check("重启后仍不得执行已跳过的顺延承诺",
            !atDue.ShouldShutdown,
            atDue.ShouldShutdown ? "跳过标记没跨过重启" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 过期承诺必须丢弃，不能在下次启动时突然关机。
    /// </summary>
    private static void TestShutdownStateExpiresStalePromise()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        ShutdownService.PromiseShutdown(new DateTime(2026, 10, 12, 22, 0, 0));
        string text = ShutdownService.ExportStateText();

        ShutdownService.ResetRuntimeState();
        string summary = ShutdownService.ImportStateText(text,
            new DateTime(2026, 10, 12, 23, 0, 0));

        Check("超过宽限的旧承诺应被丢弃",
            summary.Contains("放弃"),
            $"摘要未说明放弃原因：{summary}");

        var r = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("旧承诺不得在下次启动时突然关机",
            !r.ShouldShutdown,
            r.ShouldShutdown ? "上次运行的承诺把这次启动的机器关了" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 状态文件损坏时宁可当没有，也不能猜。
    /// </summary>
    private static void TestShutdownStateRejectsGarbage()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(23, 0, DayOfWeek.Monday)
        };

        ShutdownService.ResetRuntimeState();
        string summary = ShutdownService.ImportStateText("not-a-state-file",
            new DateTime(2026, 10, 12, 23, 0, 0));

        Check("损坏的状态文件应被明确拒绝",
            summary.Contains("格式"), $"实际摘要：{summary}");

        var r = ShutdownService.Poll(rules, new DateTime(2026, 10, 12, 23, 0, 0));
        Check("损坏状态不得产生关机承诺",
            !r.ShouldShutdown, r.ShouldShutdown ? "损坏文件被当成了有效承诺" : "");

        ShutdownService.ResetRuntimeState();
    }

    /// <summary>
    /// 序列化往返：规则写进注册表再读回来必须完全一致。
    /// 这条把界面 → 注册表 → 下次启动这条完整链路的中间段固定住。
    /// </summary>
    private static void TestShutdownSerializationRoundTrip()
    {
        var rules = new System.Collections.Generic.List<ShutdownTimeRule>
        {
            MakeRule(7, 30, DayOfWeek.Monday, DayOfWeek.Friday),
            new ShutdownTimeRule { Hour = 22, Minute = 0, DayMask = ShutdownTimeRule.EveryDayMask },
            MakeRule(12, 15, DayOfWeek.Sunday)
        };

        string serialized = SettingsStore.SerializeRules(rules);
        var back = SettingsStore.DeserializeRules(serialized);

        Check($"序列化后条数不变（{rules.Count} 条）",
            back.Count == rules.Count, $"得到 {back.Count} 条");

        bool identical = back.Count == rules.Count;
        string detail = null;
        for (int i = 0; identical && i < rules.Count; i++)
        {
            if (back[i].Hour != rules[i].Hour
                || back[i].Minute != rules[i].Minute
                || back[i].DayMask != rules[i].DayMask)
            {
                identical = false;
                detail = $"第 {i + 1} 条：原 {rules[i].Describe()} → "
                       + $"还原为 {back[i].Describe()}";
            }
        }
        Check("序列化往返：每条规则的时间与星期完全一致", identical, detail);

        // 空列表往返
        Check("空列表序列化为空串",
            SettingsStore.SerializeRules(new System.Collections.Generic.List<ShutdownTimeRule>()) == "");
        Check("空串反序列化为空列表",
            SettingsStore.DeserializeRules("").Count == 0);
        Check("null 反序列化为空列表",
            SettingsStore.DeserializeRules(null).Count == 0);
    }

    /// <summary>
    /// 序列化容错：一条损坏不得连累其余规则。
    ///
    /// 【为什么逐条隔离而不是整体回退】规则是用户手配的。
    /// 一条写坏就让全部关机设置消失，代价太大——用户会以为程序坏了。
    /// </summary>
    private static void TestShutdownSerializationToleratesGarbage()
    {
        // 中间夹一条非数字
        var good = new ShutdownTimeRule { Hour = 23, Minute = 0, DayMask = ShutdownTimeRule.EveryDayMask };
        string mixed = good.ToPacked() + ";这不是数字;" + good.ToPacked();

        var parsed = SettingsStore.DeserializeRules(mixed);
        Check("损坏项被跳过，其余规则仍保留（2 条）",
            parsed.Count == 2, $"得到 {parsed.Count} 条");

        // 全损坏
        var allBad = SettingsStore.DeserializeRules("abc;def;;ghi");
        Check("全部损坏时返回空列表且不抛异常", allBad.Count == 0,
            $"得到 {allBad.Count} 条");

        // 空项（末尾多余分隔符）不得产生幽灵规则
        var trailing = SettingsStore.DeserializeRules(good.ToPacked() + ";;");
        Check("末尾多余分隔符不产生多余规则",
            trailing.Count == 1, $"得到 {trailing.Count} 条");

        // 超量输入必须被截到上限
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < 20; i++)
        {
            if (sb.Length > 0) sb.Append(';');
            sb.Append(new ShutdownTimeRule { Hour = i % 24, Minute = i, DayMask = 127 }.ToPacked());
        }
        var clamped = SettingsStore.DeserializeRules(sb.ToString());
        Check($"超过上限时截断到 {SeewoOpt.Services.AppSettings.MaxShutdownRules} 条",
            clamped.Count == SeewoOpt.Services.AppSettings.MaxShutdownRules,
            $"得到 {clamped.Count} 条");
    }

    // ---------------------------------------------------------------
    // 守护天数
    //
    // 【为什么这一组必须存在】
    // GuardedDays 是界面上唯一一处"用户能一眼看出对错"的数字。
    // 它没有外部依赖、没有异常路径，极易被当成"这么简单不用测"——
    // 而它恰恰错得最隐蔽：写成 (today - FirstRunDate).TotalDays 时，
    // 同一天的多数时刻仍会算出看起来合理的值，只有跨午夜的那一小段
    // 才暴露为 0 天。测试必须钉住"跨午夜"和"时刻无关"这两条。
    //
    // 【2026-10-11 用户改口径：首次运行当天算第 0 天】
    // 原契约是"首次运行当天为第 1 天"（+1 补偿）。用户明确要求改为
    // "第一次运行程序的那天不算 1 天，算 0 天"——那天结束时才算守护满
    // 一天，开口就说"已守护 1 天"是虚报。整组断言随之改口径。
    // ---------------------------------------------------------------

    /// <summary>首次运行当天：第 0 天（用户 2026-10-11 明确要求）。</summary>
    private static void TestGuardedDaysFirstDayIsZero()
    {
        var s = new SeewoOpt.Services.AppSettings
        {
            FirstRunDate = new DateTime(2026, 10, 10, 9, 30, 0)
        };

        int days = s.GuardedDays(new DateTime(2026, 10, 10, 9, 31, 0));
        Check("首次运行当天为第 0 天", days == 0, $"得到 {days}");

        // 当天最后一刻仍是第 0 天——跨过午夜才变成第 1 天
        int lastMoment = s.GuardedDays(new DateTime(2026, 10, 10, 23, 59, 59));
        Check("首次运行当天 23:59 仍是第 0 天", lastMoment == 0, $"得到 {lastMoment}");

        int nextDay = s.GuardedDays(new DateTime(2026, 10, 11, 0, 0, 0));
        Check("跨过午夜变成第 1 天", nextDay == 1, $"得到 {nextDay}");
    }

    /// <summary>
    /// 跨午夜即算新的一天——这条专门抓"按时刻相减"的错误写法。
    /// 昨天 23:00 首次运行、今天 08:00 打开，必须是第 1 天。
    /// </summary>
    private static void TestGuardedDaysCrossesMidnight()
    {
        var s = new SeewoOpt.Services.AppSettings
        {
            FirstRunDate = new DateTime(2026, 10, 10, 23, 0, 0)
        };

        int days = s.GuardedDays(new DateTime(2026, 10, 11, 8, 0, 0));
        Check("昨天 23:00 → 今天 08:00 为第 1 天（跨午夜）",
            days == 1, $"得到 {days}｜按时刻相减会算成 {(new DateTime(2026,10,11,8,0,0) - s.FirstRunDate).TotalDays.ToString("0.####")}");

        // 反向确认：同一天内不论怎么过，都是第 0 天
        int sameDayLate = s.GuardedDays(new DateTime(2026, 10, 10, 23, 59, 59));
        Check("同一天 23:00 → 23:59 仍为第 0 天", sameDayLate == 0, $"得到 {sameDayLate}");
    }

    /// <summary>未记录首次运行日期时（旧版本升级上来）回退为第 0 天。</summary>
    private static void TestGuardedDaysUnrecordedFallsBackToZero()
    {
        var s = new SeewoOpt.Services.AppSettings { FirstRunDate = DateTime.MinValue };
        int days = s.GuardedDays(new DateTime(2026, 10, 10, 12, 0, 0));
        Check("未记录起始日时回退为第 0 天", days == 0, $"得到 {days}");
    }

    /// <summary>
    /// 天数永不倒退。若注册表里的日期被手工改成未来，
    /// 直接相减会得到负数——界面上就成了"已守护 -5 天"。
    /// </summary>
    private static void TestGuardedDaysNeverGoesBackwards()
    {
        var s = new SeewoOpt.Services.AppSettings
        {
            FirstRunDate = new DateTime(2030, 1, 1, 0, 0, 0)   // 未来
        };

        int days = s.GuardedDays(new DateTime(2026, 10, 10, 12, 0, 0));
        Check("起始日在未来时不出现负天数", days == 0, $"得到 {days}");
    }

    /// <summary>
    /// 天数只与"日期"有关，与当天几点无关。
    /// 把起始日固定在 10 月 10 日，则 10 月 15 日不论何时打开都应是第 5 天。
    /// </summary>
    private static void TestGuardedDaysIgnoresTimeOfDay()
    {
        var s = new SeewoOpt.Services.AppSettings
        {
            FirstRunDate = new DateTime(2026, 10, 10, 6, 0, 0)
        };

        var probes = new[]
        {
            new DateTime(2026, 10, 15, 0, 0, 0),
            new DateTime(2026, 10, 15, 5, 59, 0),
            new DateTime(2026, 10, 15, 6, 0, 0),
            new DateTime(2026, 10, 15, 23, 59, 59),
        };

        bool allFive = true;
        string detail = null;
        foreach (var p in probes)
        {
            int d = s.GuardedDays(p);
            if (d != 5)
            {
                allFive = false;
                detail = $"{p:yyyy-MM-dd HH:mm:ss} 得到 {d} 天，应为 5";
                break;
            }
        }
        Check("10/10 起算，10/15 全天各时刻均为第 5 天", allFive, detail);
    }

    /// <summary>
    /// 守护起始日的可信度判定——**直接用实机那条错误记录回归**。
    ///
    /// 【实机缺陷】2026-10-11，用户机器系统时钟停在 2000-01-01，
    /// 程序在那一刻把"首次运行日期"记成了 2000-01-01；
    /// NTP 校准到真实日期后，界面显示 **9781 天**（用户截图）。
    /// 9781 = (2026-10-11 − 2000-01-01) + 1，正是旧公式算出来的值。
    ///
    /// 修法：同步成功之后才记录起始日，并且每次都校验旧值，
    /// 剔除"不可能成立"的（晚于今天 / 早于功能上线日）。
    /// </summary>
    private static void TestFirstRunDateTrust()
    {
        DateTime today = new DateTime(2026, 10, 11);

        // 判定的三个方向
        Check("空值判定为 Missing（尚未记录）",
            SeewoOpt.Services.GuardDayPolicy.Judge(DateTime.MinValue, today)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.Missing);

        Check("实机那条 2000-01-01 判定为不可信（早于功能上线日）",
            SeewoOpt.Services.GuardDayPolicy.Judge(new DateTime(2000, 1, 1), today)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.BeforeFeatureEpoch,
            "不剔除它，界面就会继续显示那个 9781 天的假数字");

        Check("落在未来的日期判定为不可信",
            SeewoOpt.Services.GuardDayPolicy.Judge(new DateTime(2030, 1, 1), today)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.InFuture);

        Check("功能上线当天判定为可信（边界不能过严）",
            SeewoOpt.Services.GuardDayPolicy.Judge(
                SeewoOpt.Services.GuardDayPolicy.FeatureEpoch, today)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.Trustworthy);

        Check("今天判定为可信",
            SeewoOpt.Services.GuardDayPolicy.Judge(today, today)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.Trustworthy);

        // 【为什么这里要换一个"今天"】
        // 功能 2026-10-10 才上线，用 2026-10-11 当"今天"时，
        // 不存在"跑了很久且仍然可信"的样本（上线前一天的值本就该判不可信）。
        // 把时间推到一年后再看：一个跑了 200 多天的起始日仍应被采信——
        // 判定只剔除逻辑上不可能的，不做"看起来太久"的主观判断。
        DateTime nextYear = new DateTime(2027, 6, 1);
        Check("跑了 200 多天的起始日仍判定为可信（不按'太久'来怀疑）",
            SeewoOpt.Services.GuardDayPolicy.Judge(new DateTime(2026, 10, 10), nextYear)
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.Trustworthy,
            "若这条被改成不可信，正常用了大半年的用户会被无故清零");

        // 解析结果：不可信一律重置为今天
        SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict v1;
        DateTime fixedUp = SeewoOpt.Services.GuardDayPolicy.ResolveTrustedFirstRunDate(
            new DateTime(2000, 1, 1), today, out v1);
        Check("2000-01-01 被重置为今天（从第 0 天重新开始）",
            fixedUp == today,
            $"得到 {fixedUp:yyyy-MM-dd}");

        SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict v2;
        DateTime kept = SeewoOpt.Services.GuardDayPolicy.ResolveTrustedFirstRunDate(
            new DateTime(2026, 10, 10), today, out v2);
        Check("可信的旧值原样沿用，不被重置",
            kept == new DateTime(2026, 10, 10) && v2
                == SeewoOpt.Services.GuardDayPolicy.FirstRunDateVerdict.Trustworthy,
            $"得到 {kept:yyyy-MM-dd}（{v2}）");

        // 量级留证：那条错误记录若照旧使用，会显示成多大的数字
        var bogus = new SeewoOpt.Services.AppSettings
        {
            FirstRunDate = new DateTime(2000, 1, 1)
        };
        int bogusDays = bogus.GuardedDays(today);
        Check($"留证：以 2000-01-01 为起点会算出 {bogusDays} 天（实机截图显示 9781）",
            bogusDays == 9780,
            $"得到 {bogusDays}｜旧公式 +1 正是截图里的 9781");
    }

    // ---------------------------------------------------------------
    // 升级迁移：从旧版本升级上来的用户，要补勾"程序启动时自动开始执行既定任务"
    //
    // 【为什么这一组必须存在】
    // 判定依赖"注册表设置键是否存在"，而这个事实只在**第一次启动**时成立：
    // 全新安装第一次启动后，程序自己会把键建起来（保存设置、记录守护起始日）。
    // 于是最容易写错的不是三种判定本身，而是"忘记给全新安装落标记"——
    // 那样第二次启动就会被判成"老用户"，替新用户把选项勾上。
    // 这个错误第一次测试永远是好的，只有隔一次启动才暴露，
    // 所以必须用测试把"落标记"这条义务钉死。
    // ---------------------------------------------------------------

    /// <summary>已处理过的安装：不再重复处理（幂等），也无需再写标记。</summary>
    private static void TestAutoStartUpgradeIdempotent()
    {
        SeewoOpt.Services.AutoStartTaskUpgradeVerdict verdict =
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.Judge(
                markerPresent: true, settingsKeyExists: true);

        Check("迁移标记已存在时不重复处理",
            verdict == SeewoOpt.Services.AutoStartTaskUpgradeVerdict.AlreadyApplied,
            $"得到 {verdict}");

        Check("已处理过时不再写标记（否则每次启动都写注册表）",
            !SeewoOpt.Services.AutoStartTaskUpgradePolicy.ShouldWriteMarker(verdict),
            "得到 true");
    }

    /// <summary>老用户升级：键存在且无标记 → 补勾；处理完成后不再重复。</summary>
    private static void TestAutoStartUpgradeForLegacyUser()
    {
        SeewoOpt.Services.AutoStartTaskUpgradeVerdict verdict =
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.Judge(
                markerPresent: false, settingsKeyExists: true);

        Check("注册表里已有设置且无迁移标记 → 判为老用户升级",
            verdict == SeewoOpt.Services.AutoStartTaskUpgradeVerdict.EnableForUpgrade,
            $"得到 {verdict}");

        Check("老用户升级时要补勾\"启动时自动执行\"",
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.ShouldEnableAutoStartTask(verdict),
            "得到 false（老用户升级后不会自动勾选）");

        Check("老用户升级同样要落标记（否则每次启动都重复补勾）",
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.ShouldWriteMarker(verdict),
            "得到 false");

        // 处理完成后的闭环：第二次启动必须是"已处理"
        SeewoOpt.Services.AutoStartTaskUpgradeVerdict second =
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.Judge(
                markerPresent: true, settingsKeyExists: true);
        Check("老用户升级处理完成后，下次启动不再重复",
            second == SeewoOpt.Services.AutoStartTaskUpgradeVerdict.AlreadyApplied,
            $"得到 {second}");
    }

    /// <summary>
    /// 全新安装：不替用户做决定（保持默认不勾），
    /// 但**必须**落标记——否则第二次启动会被误判成老用户。
    /// </summary>
    private static void TestAutoStartUpgradeForFreshInstall()
    {
        SeewoOpt.Services.AutoStartTaskUpgradeVerdict verdict =
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.Judge(
                markerPresent: false, settingsKeyExists: false);

        Check("注册表里没有设置 → 判为全新安装",
            verdict == SeewoOpt.Services.AutoStartTaskUpgradeVerdict.KeepDefaultForNewInstall,
            $"得到 {verdict}");

        Check("全新安装不补勾（保持默认不勾选）",
            !SeewoOpt.Services.AutoStartTaskUpgradePolicy.ShouldEnableAutoStartTask(verdict),
            "得到 true（替新用户做了主）");

        // 【本组最关键的一条】全新安装也必须落标记。
        // 漏了它，第二次启动就是"键存在 + 无标记" → 判成老用户 → 自动勾选。
        Check("全新安装也必须落标记（否则第二次启动会被误判成老用户）",
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.ShouldWriteMarker(verdict),
            "得到 false——这正是\"第二次启动被误判\"的成因");

        // 两次启动的完整推演：第一次（无键）→ 按判定落标记 → 第二次（有键）
        bool markerAfterFirstRun = SeewoOpt.Services.AutoStartTaskUpgradePolicy
            .ShouldWriteMarker(verdict);
        SeewoOpt.Services.AutoStartTaskUpgradeVerdict second =
            SeewoOpt.Services.AutoStartTaskUpgradePolicy.Judge(
                markerPresent: markerAfterFirstRun, settingsKeyExists: true);
        Check("推演：全新安装第二次启动仍判为\"已处理\"，不会补勾",
            second == SeewoOpt.Services.AutoStartTaskUpgradeVerdict.AlreadyApplied,
            $"得到 {second}（新用户被误判成老用户）");
    }

    /// <summary>
    /// 设置规整必须把超量关机规则截到上限。
    /// 注册表可以手工写入 20 条，界面拦不住——上限得在模型层收紧。
    /// </summary>
    private static void TestSettingsNormalizeClampsRuleCount()
    {
        var s = SeewoOpt.Services.AppSettings.CreateDefault();
        for (int i = 0; i < 20; i++)
        {
            s.ShutdownRules.Add(new SeewoTimeRuleProbe(i).Rule);
        }

        s.Normalize();

        Check($"Normalize 把规则数收紧到 {SeewoOpt.Services.AppSettings.MaxShutdownRules} 条",
            s.ShutdownRules.Count == SeewoOpt.Services.AppSettings.MaxShutdownRules,
            $"得到 {s.ShutdownRules.Count} 条");

        // null 列表也不得让 Normalize 抛异常
        var n = SeewoOpt.Services.AppSettings.CreateDefault();
        n.ShutdownRules = null;
        n.Normalize();
        Check("Normalize 把 null 规则列表补成空列表",
            n.ShutdownRules != null && n.ShutdownRules.Count == 0);

        // 音量边界
        var v = SeewoOpt.Services.AppSettings.CreateDefault();
        v.VolumeLevel = 250; v.Normalize();
        Check("Normalize 把音量上限收到 100", v.VolumeLevel == 100, $"得到 {v.VolumeLevel}");
        v.VolumeLevel = -30; v.Normalize();
        Check("Normalize 把音量下限收到 0", v.VolumeLevel == 0, $"得到 {v.VolumeLevel}");
    }

    /// <summary>
    /// 多公钥信任链：用第二把私钥签名的包，必须被“当前 + 下一把”接受。
    ///
    /// 【为什么必须有这条】换钥时旧客户端只认第一把；如果只支持单钥，
    /// 新钥签的包会被所有旧客户端拒绝，而且补丁本身也发不出去。
    /// 这条测试固定住“按顺序试每一把、命中第二把也算通过”的契约。
    /// </summary>
    private static void TestSignatureVerifierAcceptsSecondKey()
    {
        string file = null;
        string sig = null;
        try
        {
            using (RSA oldKey = RSA.Create(2048))
            using (RSA newKey = RSA.Create(2048))
            {
                byte[] payload = Encoding.UTF8.GetBytes("seewo-update-payload");
                file = Path.Combine(Path.GetTempPath(),
                    "seewo-sig-" + Guid.NewGuid().ToString("N") + ".bin");
                sig = file + ".sig";

                File.WriteAllBytes(file, payload);
                File.WriteAllBytes(sig, newKey.SignData(
                    payload, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1));

                string oldPub = Convert.ToBase64String(oldKey.ExportSubjectPublicKeyInfo());
                string newPub = Convert.ToBase64String(newKey.ExportSubjectPublicKeyInfo());

                SignatureCheckResult both = SignatureVerifier.VerifyWithKeys(
                    file, sig, new[] { oldPub, newPub });
                Check("第二把公钥签名的包应被接受",
                    both.IsValid && both.KeyIndex == 1,
                    both.Describe());

                SignatureCheckResult oldOnly = SignatureVerifier.VerifyWithKeys(
                    file, sig, new[] { oldPub });
                Check("只信任旧公钥时必须拒绝新公钥签名",
                    !oldOnly.IsValid,
                    oldOnly.Describe());

                File.WriteAllBytes(file, Encoding.UTF8.GetBytes("tampered"));
                SignatureCheckResult tampered = SignatureVerifier.VerifyWithKeys(
                    file, sig, new[] { oldPub, newPub });
                Check("文件被改动后两把公钥都应拒绝",
                    !tampered.IsValid,
                    tampered.Describe());
            }
        }
        finally
        {
            try { if (file != null && File.Exists(file)) File.Delete(file); } catch { }
            try { if (sig != null && File.Exists(sig)) File.Delete(sig); } catch { }
        }
    }

    /// <summary>构造规整用规则的辅助——只为凑数量，语义无关。</summary>
    private sealed class SeewoTimeRuleProbe
    {
        public ShutdownTimeRule Rule { get; }
        public SeewoTimeRuleProbe(int seed)
        {
            Rule = new ShutdownTimeRule
            {
                Hour = seed % 24,
                Minute = seed % 60,
                DayMask = ShutdownTimeRule.EveryDayMask
            };
        }
    }

    // ---------------------------------------------------------------

    /// <summary>便捷重载：不需要附加说明时省略 detail</summary>
    // ---------------------------------------------------------------
    // 多更新源选源（Gitee / GitHub）
    // ---------------------------------------------------------------

    /// <summary>
    /// 版本高的优先，与来源无关——Gitee 落后时不能让用户卡在旧版。
    /// </summary>
    private static void TestPickNewerPrefersHigherVersion()
    {
        UpdateCandidate gitee = MakeCandidate("Gitee", true, "26.10.16", "https://gitee.com/a.exe");
        UpdateCandidate github = MakeCandidate("GitHub", false, "26.10.17", "https://github.com/a.exe");

        UpdateCandidate best = UpdateSourcePolicy.PickNewer(gitee, github);

        Check("选源：GitHub 版本更高时选 GitHub", best == github, best == null ? "null" : best.SourceName);
    }

    /// <summary>版本相同时优先国内源——教室里下载更快。</summary>
    private static void TestPickNewerPrefersDomesticOnTie()
    {
        UpdateCandidate gitee = MakeCandidate("Gitee", true, "26.10.17", "https://gitee.com/a.exe");
        UpdateCandidate github = MakeCandidate("GitHub", false, "26.10.17", "https://github.com/a.exe");

        Check("选源：同版本时选 Gitee", UpdateSourcePolicy.PickNewer(gitee, github) == gitee);
        Check("选源：参数顺序不影响结论", UpdateSourcePolicy.PickNewer(github, gitee) == gitee);
    }

    /// <summary>任一端不可用（null）时用另一端；两端都不可用返回 null。</summary>
    private static void TestPickNewerToleratesMissingSource()
    {
        UpdateCandidate gitee = MakeCandidate("Gitee", true, "26.10.17", "https://gitee.com/a.exe");

        Check("选源：GitHub 不可用时用 Gitee", UpdateSourcePolicy.PickNewer(gitee, null) == gitee);
        Check("选源：Gitee 不可用时用 GitHub", UpdateSourcePolicy.PickNewer(null, gitee) == gitee);
        Check("选源：两个都不可用返回 null", UpdateSourcePolicy.PickNewer(null, null) == null);
    }

    /// <summary>
    /// 备用地址只收同版本的源：另一个源还没同步时，
    /// 拿它的地址只会下到旧版——而签名对旧版同样有效，只能靠版本校验兜住。
    /// </summary>
    private static void TestBuildDownloadUrlsOnlySameVersion()
    {
        UpdateCandidate gitee = MakeCandidate("Gitee", true, "26.10.17", "https://gitee.com/a.exe");
        UpdateCandidate githubSame = MakeCandidate("GitHub", false, "26.10.17", "https://github.com/a.exe");
        UpdateCandidate githubOlder = MakeCandidate("GitHub", false, "26.10.16", "https://github.com/old.exe");

        var same = UpdateSourcePolicy.BuildDownloadUrls(gitee, new[] { gitee, githubSame });
        Check("地址：同版本收两个地址", same.Count == 2, string.Join(", ", same));
        Check("地址：主地址排在最前", same[0] == "https://gitee.com/a.exe");
        Check("地址：重复地址去重",
            UpdateSourcePolicy.BuildDownloadUrls(gitee, new[] { gitee, gitee }).Count == 1);

        var older = UpdateSourcePolicy.BuildDownloadUrls(gitee, new[] { gitee, githubOlder });
        Check("地址：另一个源版本更旧时不收", older.Count == 1, string.Join(", ", older));

        Check("地址：没有主源时返回空列表",
            UpdateSourcePolicy.BuildDownloadUrls(null, new[] { gitee }).Count == 0);
    }

    /// <summary>
    /// 三段与四段版本号必须视为相等：
    /// AssemblyVersion 常是 26.10.17，FileVersion 可能被写成 26.10.17.0，
    /// 直接 CompareTo 会把一次正常更新误判成"版本不符"。
    /// </summary>
    private static void TestVersionMatchesIgnoresMissingRevision()
    {
        Check("版本：26.10.17 与 26.10.17.0 视为相同",
            UpdateSourcePolicy.VersionMatches(new Version("26.10.17.0"), new Version("26.10.17")));

        Check("版本：修订号不同不影响判定",
            UpdateSourcePolicy.VersionMatches(new Version("26.10.17.7"), new Version("26.10.17")));
    }

    /// <summary>低版本必须被拒绝——这是镜像返回旧版 exe 时唯一的拦截条件。</summary>
    private static void TestVersionMatchesRejectsOlderBuild()
    {
        Check("版本：26.10.16 不能冒充 26.10.17",
            !UpdateSourcePolicy.VersionMatches(new Version("26.10.16"), new Version("26.10.17")));

        Check("版本：空版本号一律拒绝",
            !UpdateSourcePolicy.VersionMatches(null, new Version("26.10.17")));
    }

    /// <summary>
    /// Gitee 的发行版 JSON 与 GitHub 有两处实际差异，都要能吃下：
    ///   1. assets 里**没有 size 字段**（反序列化后是 0，日志里不能当真）；
    ///   2. assets 里混着自动生成的源码包（/archive/refs/tags/...），
    ///      它不是可执行文件，必须被跳过。
    /// 样本取自真实响应的裁剪版，而不是照着 GitHub 的字段手写。
    /// </summary>
    private static void TestGiteeReleaseJsonParsesWithoutSize()
    {
        const string json =
            "{\"id\":784180,\"tag_name\":\"v26.10.17\",\"prerelease\":false," +
            "\"body\":\"修复关机补执行。\\n\\nSHA256: " +
            "F8CF4CB4CC73FB171731F7161437A79BDADB571FFBD4FC681D97B055C0CF288B\"," +
            "\"assets\":[" +
            "{\"browser_download_url\":\"https://gitee.com/x/y/archive/refs/tags/v26.10.17.zip\"," +
            "\"name\":\"v26.10.17.zip\"}," +
            "{\"browser_download_url\":\"https://gitee.com/x/y/releases/download/v26.10.17/SeewoOpt.exe\"," +
            "\"name\":\"SeewoOpt.exe\"}," +
            "{\"browser_download_url\":\"https://gitee.com/x/y/releases/download/v26.10.17/SeewoOpt.exe.sig\"," +
            "\"name\":\"SeewoOpt.exe.sig\"}" +
            "]}";

        bool parsed = ReleaseInfoParser.TryParse(json, out ReleaseInfo release, out string error);
        Check("Gitee：缺少 size 字段也能解析", parsed, error);
        if (!parsed) return;

        Version tagVersion = ReleaseInfoParser.ParseVersion(release.TagName);
        Check("Gitee：tag 解析为 26.10.17",
            tagVersion != null && tagVersion.ToString() == "26.10.17",
            release.TagName);

        ReleaseAsset asset = ReleaseInfoParser.SelectExecutableAsset(release.Assets);
        Check("Gitee：跳过源码包，选中 .exe",
            asset != null && asset.Name == "SeewoOpt.exe",
            asset == null ? "null" : asset.Name);
        Check("Gitee：不把 .sig 当成主程序",
            asset != null && !asset.Name.EndsWith(".sig", StringComparison.OrdinalIgnoreCase));

        Check("Gitee：仍能从发行说明里提取 SHA256",
            ReleaseInfoParser.TryExtractExpectedSha256(release.Body)
                == "F8CF4CB4CC73FB171731F7161437A79BDADB571FFBD4FC681D97B055C0CF288B");
    }

    /// <summary>构造候选版本的辅助。</summary>
    private static UpdateCandidate MakeCandidate(string name, bool domestic, string version, string url)
    {
        return new UpdateCandidate
        {
            SourceName = name,
            IsDomestic = domestic,
            Version = new Version(version),
            AssetName = "SeewoOpt.exe",
            DownloadUrl = url
        };
    }

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
