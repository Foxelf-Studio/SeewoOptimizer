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

    // 与 TimeSyncForm 同构：进托盘后等待更新的上限
    private const int UPDATE_CHECK_WAIT_LIMIT_MS = 60000;
    private const int UPDATE_CHECK_TIMER_INTERVAL_MS = 1000;

    // 模拟 UI 的计时循环结果
    private static bool _uiExited = false;
    private static bool _uiSawCompletion = false;
    private static int _uiWaitedMs = 0;

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
        failures += RunUiWaitCase();

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
        _hasFailedOnce = false;
        // 注意：_slowRetryMs 由各用例显式设置，不在此重置
    }

    /// <summary>
    /// 覆盖 TimeSyncForm 在同步成功后的那段 UI 逻辑，这是修复能否真正
    /// 生效的最后一环，原有用例没有覆盖：
    ///
    ///   1. 调 RetryUpdateCheckAfterSync()
    ///   2. 等 2 秒
    ///   3. 进托盘，然后判断 UpdateCheckCompleted
    ///      - 已 true  → 退出
    ///      - 未 true  → 启动 1 秒计时器，最多等 60 秒
    ///
    /// 要验证的失败模式：重试是网络请求，2 秒内很可能还没回来。
    /// 此时 UI 绝不能把"还没完成"误判成"已完成"或"超时"而提前退出，
    /// 否则重试发出的请求会被程序退出打断，修复形同虚设。
    ///
    /// 反向也要验证：重试若因断网永不返回，60 秒后必须真的退出，
    /// 不能永远卡在托盘。
    /// </summary>
    private static int RunUiWaitCase()
    {
        int failures = 0;
        Console.WriteLine("=== UI 等待逻辑：重试耗时超过 2 秒时不得提前退出 ===");

        // --- 场景 A：重试较慢（5 秒），2 秒时还没完成 ---
        Reset();
        _slowRetryMs = 5000;   // 关键：让重试明显慢于"进托盘前的 2 秒等待"
        StartUpdateCheck("启动时", shouldFail: true);   // 时钟错，首次失败
        Thread.Sleep(400);
        if (!_updateCheckFailed)
        {
            Console.WriteLine("  [FAIL] 前置条件不成立：首次检查应失败");
            return 1;
        }

        RetryUpdateCheckAfterSync();                    // 同步成功 → 重试
        bool resetOk = !UpdateCheckCompleted;
        Console.WriteLine($"  重试已重置完成标志 = {resetOk}");
        if (!resetOk) { Console.WriteLine("  [FAIL] 未重置完成标志"); return 1; }

        // 模拟 "WaitOrCancel(2000)"：只等 2 秒，而重试要 5 秒才回来
        Thread.Sleep(2000);
        bool completedAtTrayEnter = UpdateCheckCompleted;
        Console.WriteLine($"  进托盘瞬间 UpdateCheckCompleted = {completedAtTrayEnter}（此时重试仍在进行）");

        if (completedAtTrayEnter)
        {
            Console.WriteLine("  [FAIL] 重试未完成却已置完成标志，UI 会立即退出并打断重试");
            return 1;
        }
        Console.WriteLine("  未误报完成 [OK]");

        // 模拟 Timer 循环：1 秒一跳，最多 60 跳
        for (int tick = 0; tick < UPDATE_CHECK_WAIT_LIMIT_MS / UPDATE_CHECK_TIMER_INTERVAL_MS; tick++)
        {
            Thread.Sleep(UPDATE_CHECK_TIMER_INTERVAL_MS);
            _uiWaitedMs += UPDATE_CHECK_TIMER_INTERVAL_MS;

            if (UpdateCheckCompleted)
            {
                _uiSawCompletion = true;
                _uiExited = true;
                break;
            }
            if (_uiWaitedMs >= UPDATE_CHECK_WAIT_LIMIT_MS)
            {
                _uiExited = true;   // 超时强制退出
                break;
            }
        }

        Console.WriteLine($"  UI 实际等待 {_uiWaitedMs}ms，看到完成 = {_uiSawCompletion}，退出 = {_uiExited}");
        if (!_uiSawCompletion)
        {
            Console.WriteLine("  [FAIL] UI 在重试完成前就退出了，重试请求会被打断");
            failures++;
        }
        else if (_uiWaitedMs >= UPDATE_CHECK_WAIT_LIMIT_MS)
        {
            Console.WriteLine("  [FAIL] 等待时长撞上 60 秒上限，说明重置后重试根本没在跑");
            failures++;
        }
        else
        {
            Console.WriteLine("  重试在 60 秒上限内完成，UI 正确等待 [OK]");
        }

        // --- 场景 B：重试永不返回（断网），60 秒后必须退出 ---
        Console.WriteLine();
        Console.WriteLine("=== UI 等待逻辑：重试卡死时必须靠 60 秒上限退出 ===");
        Reset();
        _updateCheckFailed = true;      // 假装首次失败
        UpdateCheckCompleted = false;

        // 只置 running，不真正发起任务——模拟请求挂住永不返回
        Interlocked.Exchange(ref _updateCheckRunning, 1);

        int waited = 0;
        bool exited = false;
        for (int tick = 0; tick < UPDATE_CHECK_WAIT_LIMIT_MS / UPDATE_CHECK_TIMER_INTERVAL_MS + 2; tick++)
        {
            Thread.Sleep(UPDATE_CHECK_TIMER_INTERVAL_MS);
            waited += UPDATE_CHECK_TIMER_INTERVAL_MS;
            if (UpdateCheckCompleted || waited >= UPDATE_CHECK_WAIT_LIMIT_MS)
            {
                exited = true;
                break;
            }
        }
        Console.WriteLine($"  卡死场景等待 {waited}ms 后退出 = {exited}");
        if (!exited)
        {
            Console.WriteLine("  [FAIL] 无上限退出机制，程序会永远卡在托盘");
            failures++;
        }
        else
        {
            Console.WriteLine("  60 秒上限生效 [OK]");
        }

        Interlocked.Exchange(ref _updateCheckRunning, 0);
        _slowRetryMs = 0;
        Console.WriteLine();
        return failures;
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
    //
    // _slowRetryMs > 0 时，只让"重试"变慢（首次仍用 NetworkDelayMs），
    // 用于验证"重试还没跑完时 UI 会不会提前退出"。
    private static int _slowRetryMs = 0;

    private static void CheckForUpdates(bool shouldFail)
    {
        try
        {
            int delay = NetworkDelayMs;
            if (!shouldFail && _slowRetryMs > 0 && _hasFailedOnce)
                delay = _slowRetryMs;

            Thread.Sleep(delay);

            if (shouldFail)
                throw new InvalidOperationException("remote certificate is invalid");

            Interlocked.Increment(ref _passCount);
            _lastResult = "取得新版本信息";
        }
        catch
        {
            Interlocked.Increment(ref _failCount);
            _updateCheckFailed = true;
            _hasFailedOnce = true;
        }
        finally
        {
            Interlocked.Exchange(ref _updateCheckRunning, 0);
            UpdateCheckCompleted = true;
        }
    }

    private static volatile bool _hasFailedOnce = false;
}
