#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把一版产物同步发布到 Gitee，作为国内更新源。

为什么需要它：教室网络里 api.github.com 与 objects.githubusercontent.com
经常整段不可达，客户端会同时查 Gitee 与 GitHub。Gitee 这一路要生效，
每次发版都必须把同一个 tag、同样文件名的附件传上去。

用法（令牌建议放环境变量，避免进命令历史）：

    set GITEE_TOKEN=xxxxxxxx
    python tools\\publish_gitee.py --tag v26.10.17 ^
        --notes ..\\v26.10.17_更新日志.md ^
        --exe ..\\dist\\SeewoOpt.exe --sig ..\\dist\\SeewoOpt.exe.sig

只建发行版、不推代码：加 --no-push。
只看会做什么、不实际写：加 --dry-run。
"""

import argparse
import io
import json
import mimetypes
import os
import subprocess
import sys
import urllib.error
import urllib.request

API = "https://gitee.com/api/v5"
DEFAULT_REPO = "Foxelf-Studio/SeewoOptimizer"


def mask(text, token):
    if token:
        return text.replace(token, "***")
    return text


def request(method, path, token, payload=None, body=None, content_type=None):
    """发一个 API 请求，返回 (status, 解析后的 JSON 或原始文本)。"""
    url = API + path
    data = None
    headers = {"User-Agent": "SeewoOptimizer-Publisher"}

    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    elif body is not None:
        data = body
        headers["Content-Type"] = content_type

    token_param = ("&" if "?" in path else "?") + "access_token=" + token
    req = urllib.request.Request(url + token_param, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8", "replace")
            return resp.status, parse_json(raw)
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        return e.code, parse_json(raw)
    except urllib.error.URLError as e:
        return 0, {"message": "网络错误: %s" % e.reason}


def parse_json(raw):
    try:
        return json.loads(raw)
    except Exception:
        return {"raw": raw[:500]}


def encode_multipart(fields, file_field, file_path):
    boundary = "----SeewoOptBoundary" + os.urandom(8).hex()
    buf = io.BytesIO()

    for key, value in fields.items():
        buf.write(("--%s\r\n" % boundary).encode())
        buf.write(('Content-Disposition: form-data; name="%s"\r\n\r\n' % key).encode())
        buf.write(str(value).encode("utf-8"))
        buf.write(b"\r\n")

    filename = os.path.basename(file_path)
    ctype = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    buf.write(("--%s\r\n" % boundary).encode())
    buf.write(
        ('Content-Disposition: form-data; name="%s"; filename="%s"\r\n'
         % (file_field, filename)).encode()
    )
    buf.write(("Content-Type: %s\r\n\r\n" % ctype).encode())
    with open(file_path, "rb") as f:
        buf.write(f.read())
    buf.write(b"\r\n")
    buf.write(("--%s--\r\n" % boundary).encode())

    return buf.getvalue(), "multipart/form-data; boundary=%s" % boundary


def find_release(repo, tag, token):
    status, data = request("GET", "/repos/%s/releases/tags/%s" % (repo, tag), token)
    if status == 200 and isinstance(data, dict) and data.get("tag_name"):
        return data
    return None


def main():
    parser = argparse.ArgumentParser(description="把一版产物同步发布到 Gitee")
    parser.add_argument("--token", default=os.environ.get("GITEE_TOKEN"),
                        help="Gitee 私人令牌（默认读环境变量 GITEE_TOKEN）")
    parser.add_argument("--repo", default=DEFAULT_REPO, help="仓库路径 owner/repo")
    parser.add_argument("--tag", required=True, help="版本标签，必须与 GitHub 完全一致")
    parser.add_argument("--notes", help="发行说明 Markdown 文件")
    parser.add_argument("--exe", required=True, help="SeewoOpt.exe")
    parser.add_argument("--sig", required=True, help="SeewoOpt.exe.sig")
    parser.add_argument("--no-push", action="store_true", help="不推送代码与标签")
    parser.add_argument("--dry-run", action="store_true", help="只打印计划，不执行")
    args = parser.parse_args()

    if not args.token:
        print("缺少令牌：请设置环境变量 GITEE_TOKEN，或传 --token")
        return 2

    token = args.token

    # 先验一次令牌。否则 401 会被 find_release 当成"发行版不存在"，
    # 接着去创建，最后报一个看不出根因的错。
    if not args.dry_run:
        status, me = request("GET", "/user", token)
        if status != 200:
            print("令牌校验失败（%s）：%s" % (status, me.get("message", me)))
            return 2
        print("令牌有效：%s" % me.get("login"))

    # 附件名必须与 GitHub 上完全一致：客户端拼 .sig 地址时是
    # 在 exe 地址后面直接加后缀，改名就对不上。
    exe_name = os.path.basename(args.exe)
    sig_name = os.path.basename(args.sig)
    if sig_name != exe_name + ".sig":
        print("附件名不合法：%s 必须是 %s.sig" % (sig_name, exe_name))
        return 2

    for path in (args.exe, args.sig):
        if not os.path.isfile(path):
            print("找不到文件：%s" % path)
            return 2

    if args.notes and not os.path.isfile(args.notes):
        print("找不到发行说明：%s" % args.notes)
        return 2

    print("仓库: %s" % args.repo)
    print("标签: %s" % args.tag)
    print("附件: %s / %s" % (exe_name, sig_name))

    # 1) 推代码与标签。客户端只读发行版，但没有 tag 就建不了发行版。
    if not args.no_push:
        remote = "https://oauth2:%s@gitee.com/%s.git" % (token, args.repo)
        for ref in ("refs/heads/master", "refs/tags/%s" % args.tag):
            cmd = ["git", "push", remote, ref]
            print("\n[git] push %s" % ref)
            if args.dry_run:
                print("  (dry-run) 跳过")
                continue
            proc = subprocess.run(cmd, capture_output=True, text=True)
            output = mask((proc.stdout or "") + (proc.stderr or ""), token).strip()
            if output:
                print("  " + output.replace("\n", "\n  "))
            if proc.returncode != 0:
                print("  推送失败")
                return 1

    # 2) 建发行版（已存在就复用，方便重传附件）
    release = find_release(args.repo, args.tag, token)
    if release:
        print("\n[release] 已存在 id=%s，复用" % release.get("id"))
    else:
        body = ""
        if args.notes:
            with open(args.notes, encoding="utf-8") as f:
                body = f.read()

        payload = {
            "tag_name": args.tag,
            "name": args.tag,
            "body": body,
            "target_commitish": "master",
            "prerelease": False,
        }
        print("\n[release] 创建 %s" % args.tag)
        if args.dry_run:
            print("  (dry-run) 跳过")
            release = {"id": 0, "assets": []}
        else:
            status, data = request("POST", "/repos/%s/releases" % args.repo, token, payload=payload)
            if status not in (200, 201):
                print("  创建失败 %s: %s" % (status, data))
                return 1
            release = data
            print("  创建成功 id=%s" % release.get("id"))

    existing = {a.get("name") for a in (release.get("assets") or [])}
    release_id = release.get("id")

    # 3) 依次上传 exe、sig。顺序与 GitHub 一致：先 exe。
    for path in (args.exe, args.sig):
        name = os.path.basename(path)
        if name in existing:
            print("\n[upload] %s 已存在，跳过" % name)
            continue

        print("\n[upload] %s (%d 字节)" % (name, os.path.getsize(path)))
        if args.dry_run:
            print("  (dry-run) 跳过")
            continue

        body, ctype = encode_multipart({"access_token": token}, "file", path)
        status, data = request(
            "POST",
            "/repos/%s/releases/%s/attach_files" % (args.repo, release_id),
            token,
            body=body,
            content_type=ctype,
        )
        if status in (200, 201):
            print("  上传成功")
        else:
            print("  上传失败 %s: %s" % (status, data))
            return 1

    # 4) 复核：列出附件，确认两个都在
    if not args.dry_run:
        verify = find_release(args.repo, args.tag, token)
        names = [a.get("name") for a in ((verify or {}).get("assets") or [])]
        print("\n[verify] %s 当前附件: %s" % (args.tag, ", ".join(names) or "无"))

        missing = [n for n in (exe_name, sig_name) if n not in names]
        if missing:
            print("缺少附件: %s" % ", ".join(missing))
            return 1

    print("\n完成。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
