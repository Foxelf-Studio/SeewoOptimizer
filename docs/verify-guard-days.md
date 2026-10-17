# 守护天数手工验证

“已守护 x 天”只在**跨过午夜**时才 +1，当天无法自然验证；而且天数**只在时间同步成功后**才刷新（时钟不可信时算出的天数是错的，见 `Services/GuardDayPolicy.cs`）。所以按下面的方法来测。

> 下例以 **2026-10-17** 为“今天”，实际执行时请换成当天日期。

天数起点存在注册表 `HKCU\Software\TimeSyncTool` 的 `FirstRunDate`（REG_SZ，格式 `yyyy-MM-dd`）。改它、再点一次 **重新同步**，即可立刻看到结果——无需重启、无需改系统时间。

> `HKCU` 是命令行里的缩写，注册表编辑器（regedit）里它的全名是 **HKEY_CURRENT_USER**。想用图形界面就展开 `HKEY_CURRENT_USER → Software → TimeSyncTool`；但下面的命令行方式更快、不会点错（`HKCU` 在 cmd 里直接可用）。键不存在时 `reg add` 会自动创建。

**开始前先看一眼当前状态**：

```cmd
reg query "HKCU\Software\TimeSyncTool"
```

## ① 验证正常递增

把起始日设为**昨天**（下例假设今天是 2026-10-17）：

```cmd
reg add "HKCU\Software\TimeSyncTool" /v FirstRunDate /t REG_SZ /d 2026-10-16 /f
```

点 **重新同步** → 同步成功后界面应显示 **第 1 天**，日志出现 `守护天数已刷新：第 1 天`。

## ② 验证错误起点会被自动修正（复现历史上的“9781 天”缺陷）

```cmd
reg add "HKCU\Software\TimeSyncTool" /v FirstRunDate /t REG_SZ /d 2000-01-01 /f
reg delete "HKCU\Software\TimeSyncTool" /v LastKnownGuardedDays /f
```

重启程序 → 界面显示占位文案「正在确认守护天数…」，**不会出现 9781**（时钟不准时不再现算）。点 **重新同步** → 应显示 **第 0 天**，日志出现：

```text
守护起始日期 2000-01-01 早于本功能上线日（2026-10-10），不可能成立，已重置为 2026-10-17
守护天数已刷新：第 0 天
```

把日期改成**明天**可测另一条分支，日志会写“晚于今天…已重置”。

## ③ 验证时钟不准时不会算出假数字

把系统时间改到 2000-01-01，重启程序、**先不要同步**：界面应显示上次确认过的值（全新安装时是占位文案），而不是按 2000 年算出来的天文数字。此时再同步，应显示第 0 天。

> **注意**：起始日早于 `2026-10-10`（守护天数功能上线日）一律会被判为不可信并重置。所以手工填一个很早的日期走的是 ② 那条重置路径，想看“较大的天数”只能靠真实时间累积。

## 恢复

把 `FirstRunDate` 改回首次运行的日期，或删掉该值重新开始：

```cmd
reg delete "HKCU\Software\TimeSyncTool" /v FirstRunDate /f
reg delete "HKCU\Software\TimeSyncTool" /v LastKnownGuardedDays /f
```
