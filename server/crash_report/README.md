# 失败上报接收端

客户端把崩溃/失败上报 POST 到这里，本服务落成 JSONL，并提供一个网页看板。

## 先回答"我怎么接收"

三种方式，按省事程度排：

| 方式 | 怎么看到 | 代价 |
|---|---|---|
| **网页看板**（本服务自带） | 浏览器打开 `http://<地址>:8080/?k=<口令>` | 需要一台有公网地址的机器 |
| **JSONL 文件** | 直接看 `reports-YYYY-MM.jsonl`，或拉回本地用脚本统计 | 同上 |
| **邮件/群机器人** | 在接收端加十几行转发（下面有说明） | 同上 + 一个 SMTP 或 webhook |

省事的做法：先只跑本服务，每天扫一眼网页；量大了再考虑转发到邮箱。

**注意：你必须有一个公网可达的地址。** 没有的话，见下面的三种部署选型。

## 部署选型

### A. 自己的 VPS / 轻量服务器（推荐，最直接）

```bat
set SEEWO_TOKEN=换成长随机串
set SEEWO_DIR=D:\seewo-reports
python app.py
```

前面挂 Nginx/Caddy 上 HTTPS，然后把 `https://你的域名/report` 填进客户端。

### B. 腾讯云函数 SCF（无服务器，国内可达，免费额度足够）

新建 Python 函数，把 `app.py` 里的 `Handler.do_POST` 逻辑改成 SCF 的 `main_handler(event, context)`：
解析 `event["body"]`、校验 `event["headers"]["x-report-token"]`，把 JSON 追加到 COS
（或直接 `print` 到日志服务 CLS，用 CLS 的检索看）。HTTP 触发器给的就是公网地址。

选它的理由：教室网络访问国内云比访问 GitHub/Cloudflare 稳得多。

### C. Cloudflare Worker（不推荐）

部署最快，但国内访问 `workers.dev` 与 GitHub 一样时通时不通——而这个功能恰恰要靠
那些网络不好的机器上报。除非你有备案域名 + 国内 CDN，否则别用。

## 客户端怎么接上

1. 在 `WindowsFormsApp1/Program.cs` 里填两个常量：
   ```csharp
   private const string CRASH_REPORT_URL = "https://你的地址/report";
   private const string CRASH_REPORT_TOKEN = "与 SEEWO_TOKEN 一致";
   ```
2. 重新构建发版。**留空则完全不上报**，客户端不会发起任何请求。
3. 口令两端必须一致；不一致时服务端返回 403，客户端静默丢弃（不会重试）。

## 上报了什么，没上报什么

会发（约 1–3 KB）：

- 事件名、时间、版本、BuildId
- 系统版本字符串、是否 64 位系统、进程是否 64 位
- 物理内存 MB、CPU 型号
- 匿名设备标识（随机 GUID，与硬件无关，重装即变）
- 日志尾部最多 12 KB、异常堆栈

发送前自动替换：用户目录 → `%USERPROFILE%`、计算机名 → `<pc>`、用户名 → `<user>`、
IPv4 → `<ip>`。

不会发：姓名、账号、文件内容、屏幕截图、键鼠行为、使用时长。不做日活统计。

## 频率与节流

- 只在失败时发；同一类事件 10 分钟内最多一次，单次运行最多 8 条。
- 超时 8 秒，失败就丢，不重试、不排队、不影响主流程。

## 本地自测

```bat
set SEEWO_TOKEN=testtoken
set SEEWO_DIR=.\reports
python app.py
```

另开一个窗口：

```bat
curl -X POST -H "X-Report-Token: testtoken" -H "Content-Type: application/json" ^
  -d "{\"event\":\"test\",\"version\":\"26.10.19\"}" http://127.0.0.1:8080/report
```

返回 `204` 即成功；浏览器打开 `http://127.0.0.1:8080/?k=testtoken` 应看到那条记录。

## 这个服务不做什么

- 不做身份认证（口令只是过滤扫描器），不要把它当成安全边界；
- 默认不加密，正式部署请走 HTTPS 反代；
- 不做多用户、不做权限分级——它是给一个人看日志用的。
