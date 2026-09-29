#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把本地文件推到 GitHub 仓库（走 REST API，不需要装 git）

用法：python3 gh_push.py <owner/repo> <文件1> <文件2> ...
"""
import base64
import json
import os
import sys
import urllib.parse
import urllib.request

TOKEN_FILE = "/root/.gh_token"
BASE = "/root/ai-finance-ledger"


def api(method, url, data=None, retries=5):
    import time
    body = json.dumps(data).encode("utf-8") if data is not None else None
    last = None
    for _ in range(retries):
        req = urllib.request.Request(url, method=method, data=body)
        req.add_header("Authorization", "Bearer " + open(TOKEN_FILE).read().strip())
        req.add_header("Accept", "application/vnd.github+json")
        req.add_header("User-Agent", "rin-ledger")
        if body:
            req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            return {"_error": e.code, "_body": e.read().decode("utf-8")[:200]}
        except Exception as e:
            last = e
            time.sleep(2)
    return {"_error": "network", "_body": str(last)[:200]}


def repo_path(local):
    p = os.path.abspath(local)
    if p.startswith(BASE):
        p = p[len(BASE):].lstrip("/")
    return p


def main():
    owner_repo = sys.argv[1]
    files = sys.argv[2:]
    for f in files:
        path = repo_path(f)
        if not os.path.exists(f):
            print("SKIP 不存在:", f)
            continue
        content = base64.b64encode(open(f, "rb").read()).decode()
        quoted = urllib.parse.quote(path)
        old = api("GET", "https://api.github.com/repos/%s/contents/%s" % (owner_repo, quoted))
        body = {"message": "update %s" % path, "content": content}
        if isinstance(old, dict) and old.get("sha"):
            body["sha"] = old["sha"]
        res = api("PUT", "https://api.github.com/repos/%s/contents/%s" % (owner_repo, quoted), body)
        if isinstance(res, dict) and res.get("commit"):
            print("OK  ", path, res["commit"]["sha"][:7])
        else:
            print("FAIL", path, res)


if __name__ == "__main__":
    main()