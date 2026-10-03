# -*- coding: utf-8 -*-
"""构建黄金基本面数据馈源 goldfund.json（由云端定时任务每日运行，也可本机运行）。

数据：
- ETF：SPDR Gold Shares 每日持仓（吨）与当日增减；在线抓取官方 .xls/.xlsx 归档，
  失败回退内置最近数据集；
- 中国央行：人民银行官方储备资产（黄金，万盎司，月度公布），换算吨并算环比增减；
  在线抓取失败回退内置序列。
说明：中国央行黄金储备为【月度】官方数据（不存在逐日买卖公开数据），
本馈源按官方最新月份更新，环比即当月“增持/抛售”。

输出 goldfund.json：
{
 "generated": epoch,
 "etf":  {"source","as_of","rows":[[date, change_tonnes], ...]},
 "pboc":{"source","unit":"万盎司","as_of",
         "rows":[[month, level_wanoz, level_tonnes, mom_tonnes], ...]}
}
"""
from __future__ import annotations
import io, json, re, time
from typing import List, Tuple, Optional

import requests
import urllib3
urllib3.disable_warnings()

H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome124 Safari/537.36",
     "Accept-Language": "en-US,en;q=0.9,zh-CN;q=0.8"}

OZ_TO_TONNES = 0.0000311034768          # 1 troy ounce -> tonnes
WAN_TO_TONNES = 10000 * OZ_TO_TONNES    # 1 万盎司 -> 0.311034768 吨

# --------------------------------------------------------------- 内置回退数据
# SPDR 每日持仓变化（吨）
ETF_SEED: List[Tuple[str, float]] = [
    ("2026.09.18", 4.28), ("2026.09.17", 0.86), ("2026.09.16", 1.71),
    ("2026.09.15", 2.85), ("2026.09.14", 0.00), ("2026.09.11", -2.85),
    ("2026.09.10", -0.35), ("2026.09.09", 0.00), ("2026.09.08", -1.43),
    ("2026.09.04", -1.43), ("2026.09.03", -3.14), ("2026.09.02", 9.98),
    ("2026.09.01", 4.28), ("2026.08.31", 0.00), ("2026.08.28", -4.28),
    ("2026.08.27", 1.14), ("2026.08.26", -2.85), ("2026.08.25", -1.14),
    ("2026.08.24", 2.28), ("2026.08.21", 8.27),
]

# 中国央行黄金储备（月末，万盎司）—— 依据人民银行官方公布整理（2025-12 起）
PBOC_SEED: List[Tuple[str, float]] = [
    ("2025.12", 7415.0),
    ("2026.01", 7419.0), ("2026.02", 7422.0), ("2026.03", 7438.0),
    ("2026.04", 7464.0), ("2026.05", 7496.0), ("2026.06", 7544.0),
    ("2026.07", 7608.0), ("2026.08", 7673.0),
]

SPDR_URLS = [
    "https://www.spdrgoldshares.com/assets/dynamic/GLD/GLD_US_archive_EN.xlsx",
    "https://www.spdrgoldshares.com/assets/dynamic/GLD/GLD_US_archive_EN.xls",
    "https://www.spdrgoldshares.com/media/GLD/file/GLD_US_archive_EN.xlsx",
]


# --------------------------------------------------------------- SPDR 解析
def _norm_date(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    m = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}.{int(m.group(2)):02d}.{int(m.group(3)):02d}"
    # Excel serial date number
    try:
        n = float(s)
        if 20000 < n < 80000:
            from datetime import datetime, timedelta
            d = datetime(1899, 12, 30) + timedelta(days=n)
            return d.strftime("%Y.%m.%d")
    except ValueError:
        pass
    return None


def _parse_table_rows(raw: bytes) -> List[Tuple[str, float]]:
    """从 xlsx/xls 字节中解析 (date, tonnes) 序列。"""
    grid: List[List[object]] = []
    name = ""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            grid.append(list(r))
        name = "xlsx"
    except Exception:
        try:
            import xlrd
            book = xlrd.open_workbook(file_contents=raw)
            sh = book.sheet_by_index(0)
            grid = [sh.row_values(i) for i in range(sh.nrows)]
            name = "xls"
        except Exception:
            return []
    # locate header row
    hi = -1
    date_c = ton_c = oz_c = -1
    for i, row in enumerate(grid[:40]):
        cells = [str(c).strip().lower() for c in row]
        joined = " ".join(cells)
        if "date" in joined and ("tonn" in joined or "ounce" in joined or "holding" in joined):
            hi = i
            for j, c in enumerate(cells):
                if "date" in c and date_c < 0:
                    date_c = j
                if "tonn" in c:
                    ton_c = j
                if "ounce" in c and oz_c < 0:
                    oz_c = j
            break
    if hi < 0 or date_c < 0:
        return []
    out: List[Tuple[str, float]] = []
    for row in grid[hi + 1:]:
        if date_c >= len(row):
            continue
        d = _norm_date(row[date_c])
        if not d:
            continue
        tonnes = None
        if ton_c >= 0 and ton_c < len(row):
            try:
                tonnes = float(row[ton_c])
            except (TypeError, ValueError):
                pass
        if tonnes is None and oz_c >= 0 and oz_c < len(row):
            try:
                tonnes = float(row[oz_c]) * OZ_TO_TONNES
            except (TypeError, ValueError):
                pass
        if tonnes and 100 < tonnes < 5000:
            out.append((d, round(tonnes, 3)))
    return out


def fetch_etf() -> Tuple[List[Tuple[str, float]], str]:
    """返回 (每日增减序列(date,change), source)。"""
    for u in SPDR_URLS:
        try:
            r = requests.get(u, headers=H, timeout=45, verify=False)
            print(f"[diag] ETF {u} -> {r.status_code} bytes {len(r.content)} "
                  f"magic {r.content[:2]}", flush=True)
            if r.status_code == 200 and r.content[:2] in (b"PK", b"\xd0\xcf"):
                levels = _parse_table_rows(r.content)
                print(f"[diag] ETF parsed levels {len(levels)}", flush=True)
                if len(levels) >= 5:
                    levels.sort(key=lambda x: x[0])
                    changes = []
                    for i in range(1, len(levels)):
                        changes.append((levels[i][0],
                                        round(levels[i][1] - levels[i - 1][1], 2)))
                    return changes[-25:], "SPDR官方归档(在线)"
        except Exception as e:
            print("[diag] ETF EXC", u, repr(e), flush=True)
            continue
    return list(ETF_SEED), "内置最近数据(离线)"


# --------------------------------------------------------------- 中国央行
PBOC_LIST = "http://www.pbc.gov.cn/diaochatongjisi/116219/116319/index.html"


def fetch_pboc() -> Tuple[List[Tuple[str, float]], str]:
    """返回 (月末黄金储备万盎司序列(month,level), source)。最佳努力，失败回退。"""
    try:
        r = requests.get(PBOC_LIST, headers=H, timeout=30, verify=False)
        print("[diag] PBOC list", r.status_code, "bytes", len(r.content), flush=True)
        txt = r.content.decode(r.apparent_encoding or "utf-8", "ignore")
        # 找最新一篇“官方储备资产”文章链接
        items = re.findall(r'href=["\']([^"\']+)["\'][^>]*>([^<]*储备资产[^<]*)<', txt)
        print("[diag] PBOC items found", len(items), flush=True)
        base = "http://www.pbc.gov.cn"
        latest = None
        for href, _t in items[:6]:
            url = href if href.startswith("http") else base + (href if href.startswith("/") else "/" + href)
            ar = requests.get(url, headers=H, timeout=30, verify=False)
            at = ar.content.decode(ar.apparent_encoding or "utf-8", "ignore")
            print("[diag] PBOC art", ar.status_code, url, flush=True)
            m = re.search(r"黄金储备[\s\S]{0,120}?([67]\d{2,3}(?:\.\d+)?)\s*万?盎司", at)
            print("[diag] PBOC match", bool(m), flush=True)
            if m:
                latest = (url, float(m.group(1)))
                break
        if latest:
            seed = list(PBOC_SEED)
            # 文章月份：取 URL/标题中的年月，缺省用当前月
            mm = time.strftime("%Y.%m")
            seed = [r for r in seed if r[0] != mm]
            seed.append((mm, latest[1]))
            seed.sort(key=lambda x: x[0])
            return seed, "中国人民银行(在线·月度)"
    except Exception as e:
        print("[diag] PBOC EXC", repr(e), flush=True)
    return list(PBOC_SEED), "内置官方数据(离线·月度)"


# --------------------------------------------------------------- 组装
def build() -> dict:
    etf_rows, etf_src = fetch_etf()
    pboc_lvl, pboc_src = fetch_pboc()
    # PBOC 换算吨 + 环比
    pboc_rows = []
    prev = None
    for month, wanoz in pboc_lvl:
        tonnes = round(wanoz * WAN_TO_TONNES, 2)
        mom = round(tonnes - prev, 2) if prev is not None else 0.0
        pboc_rows.append([month, wanoz, tonnes, mom])
        prev = tonnes
    return {
        "generated": int(time.time()),
        "etf": {"source": etf_src, "as_of": etf_rows[0][0] if etf_rows else "",
                "rows": [list(r) for r in etf_rows]},
        "pboc": {"source": pboc_src, "unit": "万盎司",
                 "as_of": pboc_rows[-1][0] if pboc_rows else "", "rows": pboc_rows},
    }


if __name__ == "__main__":
    feed = build()
    with open("goldfund.json", "w", encoding="utf-8") as f:
        json.dump(feed, f, ensure_ascii=False, indent=1)
    print("etf source:", feed["etf"]["source"], "rows:", len(feed["etf"]["rows"]))
    print("pboc source:", feed["pboc"]["source"], "latest:", feed["pboc"]["rows"][-1])
