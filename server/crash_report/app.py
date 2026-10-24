#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""失败上报接收端（单文件，只用标准库）。

用途：接住 SeewoOpt 发来的崩溃/失败上报，落成 JSONL，并提供一个
      能直接看的网页。适合放在一台有公网地址的机器上（VPS / 轻量服务器 /
      家里电脑 + 内网穿透）。

启动：
    set SEEWO_TOKEN=换成你自己的口令
    set SEEWO_DIR=D:\\seewo-reports
    python app.py

然后：
    POST http://<你的地址>:8080/report   ← 客户端往这里发
    GET  http://<你的地址>:8080/?k=<口令> ← 你自己看

安全边界说明（别指望它挡住真正的攻击）：
  · 口令只用来过滤扫描器和随手乱发，不是身份认证；
  · 服务本身不做 TLS，正式用请放在 Nginx/Caddy 后面上 HTTPS，
    否则运营商会看到明文（虽然内容已脱敏）。
"""

import json
import os
import re
import threading
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

TOKEN = os.environ.get("SEEWO_TOKEN", "change-me")
DATA_DIR = os.environ.get("SEEWO_DIR", os.path.join(os.getcwd(), "reports"))
PORT = int(os.environ.get("PORT", "8080"))

MAX_BODY = 256 * 1024
_lock = threading.Lock()


def month_file(now=None):
    now = now or datetime.now(timezone.utc)
    return os.path.join(DATA_DIR, "reports-%s.jsonl" % now.strftime("%Y-%m"))


def load_recent(limit=80):
    """读最近 N 条。按月文件倒序读，够看就行。"""
    rows = []
    if not os.path.isdir(DATA_DIR):
        return rows

    files = sorted(
        (f for f in os.listdir(DATA_DIR) if re.match(r"reports-\d{4}-\d{2}\.jsonl$", f)),
        reverse=True,
    )
    for name in files:
        path = os.path.join(DATA_DIR, name)
        try:
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        rows.append(json.loads(line))
                    except ValueError:
                        continue
        except OSError:
            continue
        if len(rows) >= limit:
            break

    return rows[-limit:][::-1]


def summarize(rows):
    by_event = {}
    by_version = {}
    devices = set()
    for r in rows:
        by_event[r.get("event", "?")] = by_event.get(r.get("event", "?"), 0) + 1
        v = r.get("version", "?")
        by_version[v] = by_version.get(v, 0) + 1
        if r.get("device"):
            devices.add(r["device"])
    return by_event, by_version, len(devices)


def esc(text):
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


class Handler(BaseHTTPRequestHandler):
    server_version = "SeewoReport/1"

    def log_message(self, fmt, *args):
        # 默认日志会打每个请求，安静一点，只留错误
        pass

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path != "/":
            self.send_error(404)
            return

        query = parse_qs(parsed.query)
        if query.get("k", [""])[0] != TOKEN:
            self.send_response(401)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("需要口令：/?k=<SEEWO_TOKEN>".encode("utf-8"))
            return

        rows = load_recent()
        by_event, by_version, device_count = summarize(rows)
        self._render(rows, by_event, by_version, device_count)

    def do_POST(self):
        if urlparse(self.path).path != "/report":
            self.send_error(404)
            return

        if self.headers.get("X-Report-Token", "") != TOKEN:
            self.send_error(403)
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_error(400)
            return

        if length <= 0 or length > MAX_BODY:
            self.send_error(413)
            return

        raw = self.rfile.read(length)
        try:
            report = json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            self.send_error(400)
            return

        if not isinstance(report, dict) or not report.get("event"):
            self.send_error(400)
            return

        report["received_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with _lock, open(month_file(), "a", encoding="utf-8") as f:
                f.write(json.dumps(report, ensure_ascii=False) + "\n")
        except OSError:
            self.send_error(500)
            return

        # 控制台留一行，方便盯着
        print("[report] %s %s v%s dev=%s" % (
            report.get("received_at"),
            report.get("event"),
            report.get("version"),
            (report.get("device") or "")[:8],
        ))

        self.send_response(204)
        self.end_headers()

    def _render(self, rows, by_event, by_version, device_count):
        def table(d):
            if not d:
                return "<p class=muted>暂无</p>"
            items = "".join(
                "<tr><td>%s</td><td>%d</td></tr>" % (esc(k), v)
                for k, v in sorted(d.items(), key=lambda kv: -kv[1])
            )
            return "<table><tr><th>项</th><th>次数</th></tr>%s</table>" % items

        body = []
        for r in rows:
            detail = (r.get("detail") or "").strip()
            log = (r.get("log") or "").strip()
            body.append(
                "<details><summary><b>%s</b> · %s · v%s · %s · %dMB%s</summary>"
                "<p class=muted>%s · %s · 64位系统=%s · 进程64位=%s</p>"
                "<pre>%s</pre><pre class=log>%s</pre></details>" % (
                    esc(r.get("received_at", "")),
                    esc(r.get("event", "")),
                    esc(r.get("version", "")),
                    esc((r.get("cpu") or "")[:60]),
                    int(r.get("ram_mb") or 0),
                    " · 设备 " + esc((r.get("device") or "")[:8]) if r.get("device") else "",
                    esc(r.get("os", "")),
                    esc(r.get("build", "")),
                    esc(r.get("os64", "")),
                    esc(r.get("proc64", "")),
                    esc(detail[:4000]),
                    esc(log[-8000:]),
                )
            )

        html = """<!doctype html><html lang=zh><meta charset=utf-8>
<title>失败上报</title>
<style>
 body{font:14px/1.6 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif;
      max-width:980px;margin:24px auto;padding:0 16px;color:#222}
 h1{font-size:18px;margin:0 0 12px}
 .cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}
 .card{flex:1;min-width:200px;border:1px solid #ddd;border-radius:8px;padding:12px}
 .card b{display:block;font-size:22px}
 .muted{color:#888}
 table{border-collapse:collapse;margin:8px 0}
 td,th{border:1px solid #e3e3e3;padding:4px 10px;font-size:13px;text-align:left}
 details{border-bottom:1px solid #eee;padding:8px 0}
 summary{cursor:pointer}
 pre{background:#f7f7f7;padding:8px;border-radius:6px;overflow:auto;white-space:pre-wrap;
     font-size:12px;max-height:320px}
 pre.log{color:#444}
</style>
<h1>失败上报</h1>
<div class=cards>
 <div class=card><span class=muted>最近记录</span><b>%d</b></div>
 <div class=card><span class=muted>涉及设备</span><b>%d</b></div>
 <div class=card><span class=muted>数据目录</span><div class=muted>%s</div></div>
</div>
<div class=cards>
 <div class=card><span class=muted>按事件</span>%s</div>
 <div class=card><span class=muted>按版本</span>%s</div>
</div>
%s
</html>""" % (
            len(rows), device_count, esc(DATA_DIR),
            table(by_event), table(by_version),
            "".join(body) if body else "<p class=muted>还没有收到任何上报。</p>",
        )

        data = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    print("监听 0.0.0.0:%d，数据目录 %s" % (PORT, DATA_DIR))
    print("上报地址: POST /report   查看: GET /?k=<口令>")
    if TOKEN == "change-me":
        print("警告：仍在用默认口令，务必设置 SEEWO_TOKEN")
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()


if __name__ == "__main__":
    main()
