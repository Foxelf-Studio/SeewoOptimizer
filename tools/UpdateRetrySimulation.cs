using System;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// 复刻 Program.cs 中更新检查的并发状态机，验证"时钟错误导致检查失败后，
/// 时间同步能否成功触发重试并最终完成"。
///
/// 为什么必须这样测：静态检查只能证明代码写成了预期形状，
/// 证明不了 volatile + Interlocked 的实际可见性与时序。
/// 这个 bug 的本质就是时序问题——UI 读到了上一轮残留的完成标志。
/// </summary>
internal static class UpdateRetrySimulation
{
    // 与 Program.cs 同构的状态
    private static volatile bool UpdateCheckCompleted = false;
    private static volatile bool _updateCheckFailed = false;
    private static int _updateCheckRunning = 0;

    private static int _passCount = 0;
    private static int _failCount = 0;
    private static string _lastResult = "";

    /// <summary>模拟一次网络往返的耗时，保证并发窗口真实存在</summary>
    private const int NetworkDelayMs = 150;

    private static int Main()
    {
        int failures = 0;

        failures += RunCase("时钟错误：首次失败 → 同步 → 重试成功", clockWrongAtStart: true, expectRetry: true);
        failures += RunCase("时钟正常：首次成功 → 不应重试", clockWrongAtStart: false, expectRetry: false);
        failures += RunCase("重试期间防重入", clockWrongAtStart: true, expectRetry: true, simulateReentrancy: true);

        Console.WriteLine();
        Console.WriteLine(failures == 0 ? "全部通过" : $"失败 {failures} 项");
        return failures == 0 ? 0 : 1;
    }

    private static int RunCase(string name, bool clockWrongAtStart, bool expectRetry, bool simulateReentrancy = false)
    {
        Reset();
        Console.WriteLine("=== " + name + " ===");

        // 1) 启动时的检查（Main 里就跑了，此刻时钟还是错的）
        StartUpdateCheck("启动时", shouldFail: clockWrongAtStart);

        // 模拟真实耗时：网络请求不是瞬间返回的
        Thread.Sleep(400);

        bool completedAfterFirst = UpdateCheckCompleted;
        Console.WriteLine($"  首次检查后 UpdateCheckCompleted = {completedAfterFirst}");
        if (!completedAfterFirst)
        {
            Console.WriteLine("  [FAIL] finally 应无条件置 true");
            return 1;
        }

        if (!clockWrongAtStart)
        {
            // 时钟正常：不应重试
            RetryUpdateCheckAfterSync();
            int passDelta = _passCount - 1;
            if (passDelta != 0)
            {
                Console.WriteLine($"  [FAIL] 首次已成功却重试了 {passDelta} 次");
                return 1;
            }
            Console.WriteLine("  首次成功，未触发重试 [OK]");
            Console.WriteLine();
            return 0;
        }

        // 2) 时钟错误场景：同步成功后重试
        RetryUpdateCheckAfterSync();

        // 关键断言：重试必须重置完成标志，否则 UI 会读到上一轮的 true 直接退出
        if (UpdateCheckCompleted)
        {
            Console.WriteLine("  [FAIL] 重试前未重置 UpdateCheckCompleted，UI 会误判为已完成并退出");
            return 1;
        }
        Console.WriteLine("  重试已重置完成标志 [OK]");

        // 3) 模拟重试进行中——UI 此时绝不能看到 true
        bool sawPrematureComplete = UpdateCheckCompleted;
        if (sawPrematureComplete)
        {
            Console.WriteLine("  [FAIL] 重试进行中却读到完成标志");
            return 1;
        }
        Console.WriteLine("  重试进行中未误报完成 [OK]");

        // 4) 防重入：必须在检查"仍在进行中"时重复触发才测得到拦截效果。
        //    单独隔离并放在最后，避免受前面步骤残留任务干扰。
        if (simulateReentrancy)
        {
            // 必须等前面步骤遗留的线程池任务真正收尾。
            // 否则它们会在本用例期间继续跑并改动计数与标志位，
            // 让断言看到与本用例无关的干扰——这正是前两轮假失败的原因。
            Thread.Sleep(NetworkDelayMs * 4);

            Reset();
            int guard = Interlocked.CompareExchange(ref _updateCheckRunning, 1, 0);
            if (guard != 0)
            {
                Console.WriteLine("  [FAIL] 测试前置状态异常：标志位应为空闲");
                return 1;
            }

            int before = _passCount + _failCount;
            StartUpdateCheck("并发触发A", shouldFail: false);
            StartUpdateCheck("并发触发B", shouldFail: false);
            StartUpdateCheck("并发触发C", shouldFail: false);
            Thread.Sleep(400);

            int after = _passCount + _failCount;
            if (after != 0)
            {
                Console.WriteLine($"  [FAIL] 防重入失效，检查进行中仍发起了 {after} 次（预期 0）");
                Interlocked.Exchange(ref _updateCheckRunning, 0);
                return 1;
            }
            Console.WriteLine($"  检查进行中连续 3 次触发全部被拦截（实际发起 {after} 次）[OK]");

            Interlocked.Exchange(ref _updateCheckRunning, 0);
            StartUpdateCheck("释放后触发", shouldFail: false);
            Thread.Sleep(400);
            if (_passCount != 1)
            {
                Console.WriteLine($"  [FAIL] 释放后应能正常发起，实际成功 {_passCount} 次");
                return 1;
            }
            Console.WriteLine("  标志释放后可正常发起 [OK]");

            Console.WriteLine();
            return 0;   // 防重入是最后一项，测完即止
        }

        // 5) 重试完成
        Thread.Sleep(400);
        if (!UpdateCheckCompleted)
        {
            Console.WriteLine("  [FAIL] 重试结束后未标记完成，UI 会卡在托盘");
            return 1;
        }
        Console.WriteLine($"  重试完成，_lastResult = {_lastResult} [OK]");

        Console.WriteLine();
        return 0;
    }

    private static void Reset()
    {
        UpdateCheckCompleted = false;
        _updateCheckFailed = false;
        _updateCheckRunning = 0;
        _passCount = 0;
        _failCount = 0;
        _lastResult = "";
    }

    // 与 Program.StartUpdateCheck 同构
    private static void StartUpdateCheck(string reason, bool shouldFail)
    {
        if (Interlocked.CompareExchange(ref _updateCheckRunning, 1, 0) != 0)
            return;

        _ = Task.Run(() => CheckForUpdates(shouldFail));
    }

    // 与 Program.RetryUpdateCheckAfterSync 同构
    private static void RetryUpdateCheckAfterSync()
    {
        if (!_updateCheckFailed)
            return;

        _updateCheckFailed = false;
        UpdateCheckCompleted = false;
        StartUpdateCheck("时间同步后重试", shouldFail: false);
    }

    // 与 Program.CheckForUpdatesAsync 的成功/失败骨架同构。
    // 加了可配置延时，用来制造真实的并发重叠窗口——
    // 若检查瞬间返回，"防重入"就永远测不出东西。
    private static void CheckForUpdates(bool shouldFail)
    {
        try
        {
            Thread.Sleep(NetworkDelayMs);

            if (shouldFail)
                throw new InvalidOperationException("remote certificate is invalid");

            Interlocked.Increment(ref _passCount);
            _lastResult = "取得新版本信息";
        }
        catch
        {
            Interlocked.Increment(ref _failCount);
            _updateCheckFailed = true;
        }
        finally
        {
            Interlocked.Exchange(ref _updateCheckRunning, 0);
            UpdateCheckCompleted = true;
        }
    }
}
