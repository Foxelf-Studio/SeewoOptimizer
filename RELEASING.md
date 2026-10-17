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

## 同步发布到 Gitee（国内更新源）

教室网络里 GitHub 经常整段不可达：`api.github.com` 被 DNS 污染、
`objects.githubusercontent.com` 在 IP 层被拦。客户端因此同时查询两个源，
**版本高的优先，版本相同时优先用 Gitee 下载**。

Gitee 那一路要能生效，每次发版都必须同步一次：

1. 在 Gitee 建一个**公开**仓库，路径与 `Program.cs` 里的 `GITEE_API` 常量一致
   （当前是 `Foxelf-Studio/SeewoOptimizer`）。路径不同就改常量，两边必须对上。
   Gitee 发布发行版要求账号已实名认证。
2. 把同一个标签推上去（只推标签也够，客户端只读发行版）：
   ```bash
   git push gitee v26.10.17
   ```
3. 在 Gitee 创建发行版。**版本号必须与 GitHub 的标签完全一致**（`v26.10.17`）：
   - 客户端用 tag 解析版本号；
   - 同版本才会把 Gitee 的地址当作首选下载地址；
   - 版本号写错（比如多写个 `.1`）会让客户端认为有更新，然后被版本校验拦下。
4. 上传附件，**文件名必须与 GitHub 上完全一致**：
   - 先 `SeewoOpt.exe`，再 `SeewoOpt.exe.sig`；
   - `.sig` 的地址是 exe 地址后面直接拼 `.sig` 得来的，改名就对不上；
   - 只传 exe 不传 sig，更新会在验签阶段失败，日志写“未找到更新文件的签名”。
5. 发行说明里同样写一行 `SHA256: <64 位十六进制>`（客户端会记录，签名仍是硬门槛）。

验证附件可用（应当 `302 → 302 → 200`，不需要 Cookie、不需要登录）：

```bash
curl -sIL "https://gitee.com/Foxelf-Studio/SeewoOptimizer/releases/download/v26.10.17/SeewoOpt.exe" ^
  | findstr /I "HTTP/ location content-length"
```

已知限制：

- 单附件上限 100 MB，仓库附件总量 1 GB——650 KB 的 exe 够发上千次。
- Gitee 的发行版 API **不返回 `size` 字段**，客户端不依赖它。
- Gitee 落后于 GitHub 时不会让用户卡住：两边都查，取版本高的那个；
  但如果 GitHub 同时不可达，用户就停在旧版，所以**发版时两边都要传**。

## 轮换签名密钥

客户端只信任编译进 exe 的公钥（`Services/SignatureVerifier.cs` 的
`TrustedPublicKeysBase64` 数组）。换钥必须按下面顺序，不能跳步：

1. 生成新密钥对（RSA-3072）：
   ```bash
   openssl genrsa -out private_key_next.pem 3072
   openssl rsa -in private_key_next.pem -pubout -outform DER | base64 -w 0
   ```
2. 把新公钥的 Base64 **追加**到 `TrustedPublicKeysBase64`，旧钥保留。
3. **用旧私钥**签名并发布这一版。用户升级后，客户端同时信任旧钥和新钥。
4. 确认目标机器都已升级（看日志里的版本号/ BuildId），再改用**新私钥**签名。
   此时旧客户端只认旧钥会拒绝新包——所以第 3 步必须先完成。
5. 过渡期结束后，从数组里移除旧钥。移除旧钥的这一版必须用新钥签名，
   且要确认没有停留在旧版本的机器，否则它们再也收不到更新。

`tools/publish.cmd` 默认用 `signing\public_key.pem` 自检；换钥签名时，
把 `PUBKEY` 指向与本次私钥配套的公钥，或把新公钥覆盖到该路径。

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
