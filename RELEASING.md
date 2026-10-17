# 发布流程

这份清单的存在理由只有一个：**让 Release 里的 exe 能被它的标签复现**。
v26.10.11 就踩过一次——标签停在 r20，用户下载到的却是 r23，按标签构建出来的版本自更新必坏。

## 三条铁律

1. **标签指向真正发布的那次构建**。先提交、再打标签、再从标签构建；不允许“先打标签，后面继续提交”。
2. **一个版本号只对应一个二进制**。测试版本用预发布（Pre-release）或临时版本号，测完删除，不占用正式标签。
3. **上传顺序：先 `SeewoOpt.exe`，再 `SeewoOpt.exe.sig`**。旧客户端会取 Release 里的第一个下载地址，顺序反了会去下载签名文件。

## 步骤

1. 改 `WindowsFormsApp1/Properties/AssemblyInfo.cs`：
   - `AssemblyVersion` / `AssemblyFileVersion` 改成新版本，例如 `26.10.17`。
2. 改 `WindowsFormsApp1/Services/BuildInfo.cs` 的 `BuildId`：
   - 格式 `yyyy.MM.dd-rN`，同一个版本号下每次构建都要递增，绝不重复。
3. 提交并推送 `master`。
4. 打标签（必须与 AssemblyVersion 完全一致，只多一个 `v` 前缀）：
   ```bash
   git tag v26.10.17
   git push origin v26.10.17
   ```
5. 从标签构建，不要从工作区构建：
   ```bash
   git checkout v26.10.17
   tools\build.cmd Release
   ```
6. 签名并自检：
   ```bash
   tools\publish.cmd Release
   ```
7. 确认产物：
   - 文件属性里的 `FileVersion` == 标签去掉 `v`；
   - 启动日志里的 `BuildId` 是本次的新值；
   - 记下 `SHA256`。
8. 创建 **普通 Release**（不要勾 Pre-release），上传 exe，再上传 `.sig`。
9. Release 正文末尾写一行：
   ```
   SHA256: <64 位十六进制>
   ```
   更新器会从正文里提取它做辅助校验；签名才是硬门槛。

## 为什么测试版要发成 Pre-release

更新检查走的是 `/releases/latest`，它按 **发布时间** 而不是版本号挑“最新”。
一个测试版哪怕版本号更低，只要发得晚，也会把用户导向它。
发成 Pre-release 后 `/releases/latest` 不会返回它；测完直接删掉，不要留着占位。

## 标签校验

CI 在推送 `v*` 标签时会：

1. 干净检出该标签；
2. 编译 net48 产物；
3. 读产物自己的 `FileVersion`，与标签名比对，不一致直接失败。

因此“标签写 26.10.17、程序集还是 26.10.11”这种错误会在 CI 就暴露，不会等到用户更新时才发现。

## 一次性清理：v26.11.0 / v26.12.0

这两个标签是自更新测试用的，已确认弃用；`v26.10.11` 需要移到实际发布产物对应的提交。

```bash
# 1) 把 v26.10.11 移到实际发布的那次提交
git tag -f v26.10.11 8f04c82

# 2) 删除两个测试标签
git tag -d v26.11.0 v26.12.0

# 3) 推到远端（移动标签必须 --force；删除标签用空 refspec）
git push origin refs/tags/v26.10.11 --force
git push origin :refs/tags/v26.11.0 :refs/tags/v26.12.0
```

> 已经拉取过旧标签的机器需要 `git fetch --tags --force` 才会更新本地引用。
> 如果 GitHub Release 页面的源码归档仍指向旧提交，删掉该 Release 的归档缓存或重新发布一次即可。
