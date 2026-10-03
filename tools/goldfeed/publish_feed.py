# -*- coding: utf-8 -*-
"""把 goldfund.json 发布到 GitHub 免费云端（release 固定 tag），客户端经 api.github.com 拉取。

身份/坐标解析：
- 仓库：环境 GITHUB_REPOSITORY（云端 Action）否则默认 titan-ai-7721/titan-goldfund-data；
- Token：环境 GITHUB_TOKEN / TITAN_GH_TOKEN，否则读 %LOCALAPPDATA%\\titan_gh_token.txt。
会自动：建公开仓库 -> 空仓库初始化 README -> 维护 release tag 'data' ->
删除旧 goldfund.json 资产 -> 上传新文件。
"""
from __future__ import annotations
import json, os, sys, time
from pathlib import Path

import requests

OWNER = "titan-ai-7721"
REPO = "titan-goldfund-data"
DATA_TAG = "data"
ASSET_NAME = "goldfund.json"
API = "https://api.github.com"


def _coord():
    gr = os.environ.get("GITHUB_REPOSITORY", "").strip()
    if "/" in gr:
        o, n = gr.split("/", 1)
    else:
        o, n = OWNER, REPO
    return o, n


def _token():
    for k in ("GITHUB_TOKEN", "TITAN_GH_TOKEN"):
        v = os.environ.get(k, "").strip()
        if v:
            return v
    p = Path(os.environ.get("LOCALAPPDATA", "")) / "titan_gh_token.txt"
    if p.exists():
        return p.read_text(encoding="utf-8").strip()
    raise SystemExit("缺少 GitHub Token")


class GH:
    def __init__(self, token, owner, repo):
        self.s = requests.Session()
        self.s.headers.update({"Authorization": f"token {token}",
                               "Accept": "application/vnd.github+json",
                               "User-Agent": "TITAN-GoldFeed"})
        self.o, self.n = owner, repo

    def _call(self, method, url, **kw):
        for i in range(3):
            r = self.s.request(method, url, **kw)
            if r.status_code in (502, 503, 504):
                time.sleep(1.5 * (i + 1)); continue
            return r
        return r

    def ensure_repo(self):
        r = self._call("GET", f"{API}/repos/{self.o}/{self.n}")
        if r.status_code == 404:
            r = self._call("POST", f"{API}/user/repos",
                           json={"name": self.n, "private": False,
                                 "has_issues": False, "has_wiki": False,
                                 "has_projects": False})
            if r.status_code not in (200, 201):
                raise SystemExit("建仓库失败: " + r.text[:200])
        c = self._call("GET", f"{API}/repos/{self.o}/{self.n}/contents/README.md")
        if c.status_code == 404:
            self._call("PUT", f"{API}/repos/{self.o}/{self.n}/contents/README.md",
                       json={"message": "init",
                             "content": "VElUQU4gZ29sZCBmdW5kYW1lbnRhbHMgZmVlZA=="})

    def put_file(self, path, content_bytes, msg):
        url = f"{API}/repos/{self.o}/{self.n}/contents/{path}"
        import base64
        r = self._call("GET", url)
        data = {"message": msg, "content":
                base64.b64encode(content_bytes).decode()}
        if r.status_code == 200:
            data["sha"] = r.json()["sha"]
        r = self._call("PUT", url, json=data)
        return r.status_code in (200, 201)

    def _release_id(self):
        r = self._call("GET", f"{API}/repos/{self.o}/{self.n}/releases/tags/{DATA_TAG}")
        if r.status_code == 200:
            return r.json()["id"]
        r = self._call("POST", f"{API}/repos/{self.o}/{self.n}/releases",
                       json={"tag_name": DATA_TAG, "name": "Gold Fund Data",
                             "prerelease": False})
        if r.status_code not in (200, 201):
            raise SystemExit("建 release 失败: " + r.text[:200])
        return r.json()["id"]

    def _list_assets(self, rid):
        r = self._call("GET", f"{API}/repos/{self.o}/{self.n}/releases/{rid}/assets")
        return r.json() if r.status_code == 200 else []

    def _purge_asset(self, rid, name, tries=8):
        """删除同名资产并轮询直到真正消失（删除为最终一致，且旧 URL 易错）。"""
        for _ in range(tries):
            assets = self._list_assets(rid)
            targets = [a for a in assets if a["name"] == name]
            if not targets:
                return True
            for a in targets:
                self._call("DELETE",
                           f"{API}/repos/{self.o}/{self.n}/releases/assets/{a['id']}")
            time.sleep(1.2)
        return not any(a["name"] == name for a in self._list_assets(rid))

    def publish_asset(self, raw: bytes):
        rid = self._release_id()
        if not self._purge_asset(rid, ASSET_NAME):
            raise SystemExit("旧资产未能删除，已取消")
        url = (f"https://uploads.github.com/repos/{self.o}/{self.n}/releases/"
               f"{rid}/assets?name={ASSET_NAME}")
        for i in range(4):
            r = self.s.post(url, data=raw,
                            headers={"Content-Type": "application/json"})
            if r.status_code in (200, 201):
                return r.json()
            if r.status_code == 422 and i < 3:  # already_exists，再清再传
                self._purge_asset(rid, ASSET_NAME); time.sleep(1.2); continue
            raise SystemExit("上传资产失败: " + r.text[:200])
        raise SystemExit("上传资产多次失败")


def main():
    o, n = _coord()
    feed = Path(__file__).with_name(ASSET_NAME)
    if not feed.exists():
        raise SystemExit("未找到 goldfund.json，先运行 build_feed.py")
    raw = feed.read_bytes()
    gh = GH(_token(), o, n)
    gh.ensure_repo()
    a = gh.publish_asset(raw)
    print(f"已发布 {o}/{n} release:{DATA_TAG} {ASSET_NAME} "
          f"size={a['size']} download={a['browser_download_url']}")


if __name__ == "__main__":
    main()
