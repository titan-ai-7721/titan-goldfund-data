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
import io, json, os, re, sys, time
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

SPDR_ARCHIVE_API = ("https://api.spdrgoldshares.com/api/v1/historical-archive"
                     "?product=gld&exchange=NYSE&lang=en")
SPDR_LANDING = "https://www.spdrgoldshares.com/usa/gld/"


# --------------------------------------------------------------- SPDR 解析
def _norm_date(v) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    # dd-Mon-yyyy（SPDR 归档格式，如 18-Nov-2004）
    mn = re.search(r"(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{4})", s)
    if mn:
        mon = _MONTHS.get(mn.group(2)[:3].lower())
        if mon:
            return f"{int(mn.group(3)):04d}.{mon:02d}.{int(mn.group(1)):02d}"
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


_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
           "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def _extract_grid(grid: List[List[object]]) -> List[Tuple[str, float]]:
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
        if tonnes and 50 < tonnes < 5000:
            out.append((d, round(tonnes, 3)))
    return out


def _parse_table_rows(raw: bytes) -> List[Tuple[str, float]]:
    """从 xlsx/xls 字节中解析 (date, tonnes) 序列（遍历所有工作表定位归档表）。"""
    books: List[List[List[object]]] = []
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        for ws in wb.worksheets:
            books.append([list(r) for r in ws.iter_rows(values_only=True)])
    except Exception:
        try:
            import xlrd
            book = xlrd.open_workbook(file_contents=raw)
            sh = book.sheet_by_index(0)
            books = [[sh.row_values(i) for i in range(sh.nrows)]]
        except Exception:
            return []
    for grid in books:
        out = _extract_grid(grid)
        if out:
            return out
    return []


def _spdr_endpoint() -> str:
    """优先已知 API；若官网改版，则从落地页动态提取历史归档端点。"""
    try:
        r = requests.get(SPDR_LANDING, headers=H, timeout=40, verify=False)
        ms = re.findall(r'https://api\.spdrgoldshares\.com/[^"\']+historical-archive[^"\']*',
                        r.text)
        if ms:
            return ms[0].replace("&amp;", "&")
    except Exception as e:
        print("[diag] landing EXC", repr(e), flush=True)
    return SPDR_ARCHIVE_API


def fetch_etf() -> Tuple[List[Tuple[str, float]], str]:
    """返回 (每日增减序列(date,change), source)。"""
    u = _spdr_endpoint()
    try:
        r = requests.get(u, headers={**H, "Accept": "*/*"}, timeout=60,
                         verify=False, allow_redirects=True)
        print("[diag] ETF api", r.status_code, r.headers.get("Content-Type"),
              "bytes", len(r.content), "magic", r.content[:4], flush=True)
        raw = r.content if r.content[:2] in (b"PK", b"\xd0\xcf") else None
        if raw is None:  # 可能返回 JSON，内含下载地址
            try:
                j = r.json()
                print("[diag] ETF json top keys", list(j)[:10], flush=True)
                cand = []

                def walk(x):
                    if isinstance(x, dict):
                        for v in x.values():
                            walk(v)
                    elif isinstance(x, list):
                        for v in x:
                            walk(v)
                    elif isinstance(x, str) and re.search(r'https?://.+\.(xlsx|xls)', x, re.I):
                        cand.append(x)
                walk(j)
                if cand:
                    rr = requests.get(cand[0], headers=H, timeout=60, verify=False)
                    print("[diag] ETF json url", rr.status_code, rr.content[:2], flush=True)
                    if rr.content[:2] in (b"PK", b"\xd0\xcf"):
                        raw = rr.content
            except Exception as e:
                print("[diag] ETF body not file/json", repr(e), flush=True)
        if raw:
            levels = _parse_table_rows(raw)
            print("[diag] ETF parsed levels", len(levels), flush=True)
            if len(levels) >= 5:
                levels.sort(key=lambda x: x[0])
                changes = [(levels[i][0],
                            round(levels[i][1] - levels[i - 1][1], 2))
                           for i in range(1, len(levels))]
                return changes[-25:], "SPDR官方归档(在线)"
    except Exception as e:
        print("[diag] ETF EXC", repr(e), flush=True)
    return list(ETF_SEED), "内置最近数据(离线)"


# --------------------------------------------------------------- 中国央行
PBOC_HBTJGL = ("http://www.pbc.gov.cn/diaochatongjisi/116219/116319/"
               "2026ntjsj/hbtjgl/index.html")
PBOC_BASE = "http://www.pbc.gov.cn"


def _parse_pboc_workbook(raw: bytes) -> Tuple[Optional[str], Optional[float]]:
    """从月度『官方储备资产』工作簿提取 (月份YYYY.MM, 黄金储备万盎司)。

    黄金量取『以盎司计算的纯金数量（百万盎司）』行（如 76.73 百万盎司=7673 万盎司），
    而非黄金市值行（亿美元/亿SDR）。
    """
    sheets = None
    if raw[:2] == b"PK":
        try:
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            sheets = [(ws.title, list(ws.iter_rows(values_only=True))) for ws in wb.worksheets]
        except Exception:
            sheets = None
    if sheets is None:
        try:
            import xlrd
            bk = xlrd.open_workbook(file_contents=raw)
            sheets = [(sh.name, [sh.row_values(i) for i in range(sh.nrows)])
                      for sh in bk.sheets()]
        except Exception:
            return None, None
    month: Optional[str] = None
    mloz: Optional[float] = None
    for _name, rows in sheets:
        for r in rows:
            cells = ["" if c is None else str(c) for c in r]
            line = " ".join(cells)
            if month is None:
                mm = re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月", line)
                if mm:
                    month = f"{int(mm.group(1)):04d}.{int(mm.group(2)):02d}"
            if mloz is None and ("纯金数量" in line or "百万盎司" in line
                                 or "volume in millions" in line.lower()):
                for c in cells:
                    for tok in re.findall(r"\d+\.?\d*", c):
                        try:
                            v = float(tok)
                        except ValueError:
                            continue
                        if 20 <= v <= 200:
                            mloz = v
                            break
                    if mloz is not None:
                        break
    if month is None or mloz is None:
        return month, None
    return month, round(mloz * 100.0, 1)  # 百万盎司 -> 万盎司


def fetch_pboc() -> Tuple[List[Tuple[str, float]], str]:
    """返回 (月末黄金储备万盎司序列(month,level), source)。最佳努力，失败回退。"""
    try:
        r = requests.get(PBOC_HBTJGL, headers=H, timeout=40, verify=False)
        print("[diag] PBOC hbtjgl", r.status_code, "bytes", len(r.content), flush=True)
        r.encoding = r.apparent_encoding
        hrefs = re.findall(r'href=["\']([^"\']+\.(?:xlsx|xls))["\']', r.text)
        cands = []
        for h in hrefs[-14:]:  # 月度概览块在页面末尾
            cands.append(h if h.startswith("http")
                         else PBOC_BASE + (h if h.startswith("/") else "/" + h))
        cands.sort(key=lambda u: 0 if u.lower().endswith("xlsx") else 1)  # xlsx 优先
        online = {}
        for url in cands:
            try:
                ar = requests.get(url, headers=H, timeout=40, verify=False)
                if ar.content[:2] not in (b"PK", b"\xd0\xcf"):
                    continue
                month, level = _parse_pboc_workbook(ar.content)
                print("[diag] PBOC wb", url.split("/")[-1], month, level, flush=True)
                if level and month:
                    online.setdefault(month, level)
            except Exception as e:
                print("[diag] PBOC wb EXC", repr(e), flush=True)
        if online:
            merged = {m: v for m, v in PBOC_SEED}
            merged.update(online)
            return sorted(merged.items()), "中国人民银行(在线·月度)"
    except Exception as e:
        print("[diag] PBOC EXC", repr(e), flush=True)
    return list(PBOC_SEED), "内置官方数据(离线·月度)"


# --------------------------------------------------------------- 组装
def _get(u: str):
    return requests.get(u, headers=H, timeout=40, verify=False)


def discover():
    """结构深检：打印 SPDR 归档与 PBOC 月度工作簿的真实表结构，供修正解析。"""
    print("==== SPDR xlsx structure ====", flush=True)
    u = _spdr_endpoint()
    r = requests.get(u, headers={**H, "Accept": "*/*"}, timeout=60, verify=False)
    print("download", r.status_code, r.content[:4], flush=True)
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(r.content), read_only=True, data_only=True)
    print("sheets:", wb.sheetnames, flush=True)
    for ws in wb.worksheets[:2]:
        print("--- sheet:", ws.title, "dims", ws.max_row, ws.max_column, flush=True)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 14:
                break
            print(i, [("" if c is None else str(c)[:16]) for c in row][:12], flush=True)

    print("==== PBOC latest workbooks ====", flush=True)
    base = "http://www.pbc.gov.cn/diaochatongjisi/attachDir/2026/09/"
    for fn in ["2026093016040859408.xls", "2026093016043439203.xls"]:
        ar = requests.get(base + fn, headers=H, timeout=40, verify=False)
        raw = ar.content
        print("FILE", fn, ar.status_code, raw[:4], flush=True)
        sheets = []
        if raw[:2] == b"PK":
            wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            sheets = [(s.title, list(s.iter_rows(values_only=True))) for s in wb.worksheets]
        elif raw[:2] == b"\xd0\xcf":
            import xlrd
            bk = xlrd.open_workbook(file_contents=raw)
            sheets = [(sh.name, [sh.row_values(i) for i in range(sh.nrows)])
                      for sh in bk.sheets()]
        print(" sheets:", [s[0] for s in sheets], flush=True)
        for name, rows in sheets:
            for i, vals in enumerate(rows):
                joined = " ".join(str(v) for v in vals)
                if "黄金" in joined:
                    print("  [" + name + "] row", i,
                          [str(v)[:18] for v in vals][:10], flush=True)


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
    if os.environ.get("DISCOVER", "").lower() in ("1", "true", "yes"):
        discover()
        sys.exit(0)
    feed = build()
    with open("goldfund.json", "w", encoding="utf-8") as f:
        json.dump(feed, f, ensure_ascii=False, indent=1)
    print("etf source:", feed["etf"]["source"], "rows:", len(feed["etf"]["rows"]))
    print("pboc source:", feed["pboc"]["source"], "latest:", feed["pboc"]["rows"][-1])
