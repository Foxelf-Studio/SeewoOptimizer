# 部署到腾讯云轻量应用服务器

目标：一台 2C2G 的轻量服务器，跑起 `app.py`，用 `http://<公网IP>/report` 接收上报。

## 为什么选轻量而不是云函数

- 只有这条路能用现成代码，我已经端到端测过；
- 数据落在自己磁盘上，随时能导出、迁移、备份，不依赖任何厂商 SDK；
- 有现成网页看板，不用去日志服务里翻检索语句。

云函数确实更省（按量、几乎 0 元），但存储要接 COS 或日志服务，查看体验差很多。
一年一两百块的差距，不值得为此多维护一套没测过的上传逻辑。

## 0. 买机器

轻量应用服务器 → 新建：

- 地域：离你学校/常用地近的（广州、上海、北京）
- 镜像：**Ubuntu 22.04 LTS**
- 配置：2C2G 足够（这个服务几乎不吃 CPU）
- 带宽：3–5 Mbps 够（每次上报 1–3 KB）
- 防火墙：放行 **TCP 80**（控制台 → 防火墙 → 添加规则）

**不需要域名，也不需要备案。** 直接用公网 IP + HTTP。
（想要 HTTPS 就必须有域名，境内服务器还要备案，周期 2–4 周。上报内容已脱敏，
先用 HTTP 可以接受；介意的话再走这条路。）

## 1. 把文件传上去

在本地仓库根目录执行（把 `IP` 换成你的公网 IP）：

```bash
scp server/crash_report/app.py              root@IP:/tmp/app.py
scp server/crash_report/seewo-report.service root@IP:/tmp/seewo-report.service
```

## 2. 在服务器上装起来

```bash
ssh root@IP
```

然后逐段执行：

```bash
# 专用用户与目录（不要用 root 跑服务）
useradd --system --no-create-home --shell /usr/sbin/nologin seewo
mkdir -p /opt/seewo-report /var/lib/seewo-report
install -o seewo -g seewo -m 640 /tmp/app.py /opt/seewo-report/app.py
chown seewo:seewo /var/lib/seewo-report

# 口令：换成你自己的长随机串，客户端要填同一个
printf 'SEEWO_TOKEN=%s\n' "$(head -c 24 /dev/urandom | base64 | tr -d '/+=')" > /etc/seewo-report.env
chmod 600 /etc/seewo-report.env
cat /etc/seewo-report.env          # 把这一行抄下来，等下填客户端

# systemd
install -m 644 /tmp/seewo-report.service /etc/systemd/system/seewo-report.service
systemctl daemon-reload
systemctl enable --now seewo-report
systemctl status seewo-report --no-pager
```

## 3. 自测

```bash
source /etc/seewo-report.env
curl -s -o /dev/null -w '%{http_code}\n' -X POST \
  -H "X-Report-Token: $SEEWO_TOKEN" -H 'Content-Type: application/json' \
  -d '{"event":"deploy_test","version":"26.10.19"}' http://127.0.0.1/report
# 应输出 204

curl -s "http://127.0.0.1/?k=$SEEWO_TOKEN" | head -5
# 应看到网页，里面有一条 deploy_test
```

再从**外网**确认一次（本地机器上执行，不经过服务器）：

```bash
curl -s -o /dev/null -w '%{http_code}\n' -X POST \
  -H "X-Report-Token: 你的口令" -H 'Content-Type: application/json' \
  -d '{"event":"deploy_test","version":"26.10.19"}' http://IP/report
```

204 就说明公网通了。打不开先查轻量控制台的防火墙有没有放行 80。

## 4. 填进客户端并发版

`WindowsFormsApp1/Program.cs`：

```csharp
private const string CRASH_REPORT_URL = "http://你的IP/report";
private const string CRASH_REPORT_TOKEN = "与 SEEWO_TOKEN 完全一致";
```

然后走正常发布流程。**留空则完全不上报**，两端口令不一致会被服务端 403 拒掉。

## 日常维护

- 看数据：浏览器开 `http://IP/?k=口令`；原始数据在 `/var/lib/seewo-report/reports-YYYY-MM.jsonl`
- 备份：`rsync -av root@IP:/var/lib/seewo-report/ ./backup/`，或定期拉回本地
- 日志：`journalctl -u seewo-report -n 50`
- 升级代码：重新 scp `app.py` 到 `/opt/seewo-report/`，`systemctl restart seewo-report`

## 安全边界（务必看清）

- 这是 **HTTP 明文**。内容已脱敏（无姓名/IP/用户名），但链路上能被看到；
- 口令只挡扫描器，不是身份认证。谁拿到 IP + 口令都能往里灌数据；
- 服务不存访问者 IP，也不写访问日志（代码里刻意关掉了 `log_message`）；
- 真要被灌了：`reports-*.jsonl` 里按 `event` 过滤即可，删掉对应行就行。
