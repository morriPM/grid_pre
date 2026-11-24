"""
core.py — 电网导出数据清洗核心逻辑

与 app.py 分离: 这里不 import streamlit, 便于单独测试。

处理链路 (每一步都可调):
    读入 → 定位时间列块 → 选行(筛电表/数据类型) → 差分 → ×倍率 → 输出

三个必须显式声明的语义 (每个文件的导出约定可能不同, 不能跨文件沿用):

1) 时间步长 step
   从表头时间标签自动判定 (96列@15min / 24列@1h / ...), 步长决定一格代表多长区间。

2) 跨日差放哪一格 anchor
   left : 每行最后一格放跨日差(本行末->次日首), 表首格 = 第一个减完的数
   right: 每行第一格放跨日差(上行末->本行首), 整表首格 = 0
   两种约定下同一个物理区间落在不同列, 表头必须跟着自己那一格走。

3) 输出量纲
   Δ读数 × 综合倍率 = 该区间的电量(kWh)
   要平均功率(kW) 还要再乘 60/step (15min->x4, 1h->x1)
"""
from __future__ import annotations

import re

BIFF_MAGIC = b"\xd0\xcf\x11\xe0"
OOXML_MAGIC = b"PK\x03\x04"

TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")


# ================= 容器嗅探 =================
def sniff_container(path: str) -> str:
    """按魔数判断真实格式, 不信扩展名。返回 'xls' / 'xlsx' / 'csv'"""
    with open(path, "rb") as f:
        head = f.read(8)
    if head.startswith(BIFF_MAGIC):
        return "xls"
    if head.startswith(OOXML_MAGIC):
        return "xlsx"
    return "csv"


# ================= 表头识别 =================
def is_time_label(v) -> bool:
    return isinstance(v, str) and bool(TIME_RE.match(v.strip()))


def find_time_block(header: list) -> tuple[int, int] | None:
    """在表头里找连续的时间标签块, 返回 (起, 止) 0基闭区间。取最长的连续段。"""
    best = None
    i = 0
    while i < len(header):
        if is_time_label(header[i]):
            j = i
            while j + 1 < len(header) and is_time_label(header[j + 1]):
                j += 1
            if best is None or (j - i) > (best[1] - best[0]):
                best = (i, j)
            i = j + 1
        else:
            i += 1
    return best


def parse_step(labels: list[str]) -> int:
    """从时间标签算步长(分钟)。取相邻差的中位数, 最稳。"""
    mins = []
    for s in labels:
        try:
            h, m = str(s).strip().split(":")
            mins.append(int(h) * 60 + int(m))
        except (ValueError, AttributeError):
            continue
    diffs = sorted(d for d in
                   (mins[i + 1] - mins[i] for i in range(len(mins) - 1)) if d > 0)
    return diffs[len(diffs) // 2] if diffs else 15


def find_col(header: list, *names) -> int | None:
    """按表头名找列。返回 0基列号, 找不到 None。列位置随站点变, 不能写死。"""
    for n in names:
        for i, h in enumerate(header):
            if isinstance(h, str) and h.strip() == n:
                return i
    return None


# ================= 读入 =================
def list_sheets(path: str) -> list[str]:
    kind = sniff_container(path)
    if kind == "xls":
        import xlrd

        return xlrd.open_workbook(path).sheet_names()
    if kind == "xlsx":
        import openpyxl

        return openpyxl.load_workbook(path, read_only=True).sheetnames
    return ["csv"]


def read_table(path: str, sheet: str | None = None):
    """读成 (header, body_rows)。body_rows 是 list[list], 保持原类型。"""
    kind = sniff_container(path)
    if kind == "xls":
        import xlrd

        book = xlrd.open_workbook(path)
        sh = book.sheet_by_name(sheet) if sheet else book.sheet_by_index(0)
        header = [sh.cell_value(0, c) for c in range(sh.ncols)]
        body = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(1, sh.nrows)]
        return header, body
    if kind == "xlsx":
        import openpyxl

        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return [], []
        return list(rows[0]), [list(r) for r in rows[1:]]
    import csv

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    return rows[0], rows[1:]


def to_float(v):
    """原文件里数字常是文本型, float() 自带 try parse。"""
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ================= 探测 =================
def probe(path: str, sheet: str | None = None) -> dict:
    """探出这份表的结构, 给界面用。"""
    kind = sniff_container(path)
    header, body = read_table(path, sheet)
    blk = find_time_block(header)
    t0, t1 = blk if blk else (None, None)
    labels = [str(header[c]) for c in range(t0, t1 + 1)] if blk else []
    step = parse_step(labels) if labels else None

    meter_col = find_col(header, "电能表资产编号", "电能表编号")
    dtype_col = find_col(header, "数据类型")
    date_col = find_col(header, "数据日期", "日期")
    ratio_col = find_col(header, "综合倍率")

    return {
        "container": kind,
        "sheets": list_sheets(path),
        "ncols": len(header),
        "nrows": len(body),
        "header": header,
        "body": body,
        "time_block": blk,
        "time_labels": labels,
        "step": step,
        "n_intervals": (t1 - t0 + 1) if blk else 0,
        "meter_col": meter_col,
        "dtype_col": dtype_col,
        "date_col": date_col,
        "ratio_col": ratio_col,
        "meters": distinct_values(body, meter_col),
        "dtypes": distinct_values(body, dtype_col),
        "date_first": body[0][date_col] if body and date_col is not None else None,
        "date_last": body[-1][date_col] if body and date_col is not None else None,
    }


# ================= 选行 =================
def _same(a, b) -> bool:
    return str(a).strip() == str(b).strip()


def distinct_values(body, col) -> list[str]:
    if col is None:
        return []
    seen = []
    for r in body:
        v = str(r[col]).strip()
        if v and v not in seen:
            seen.append(v)
    return seen


def select_rows(body, meter_col=None, meter_id=None, dtype_col=None, dtype=None):
    """按电表编号 / 数据类型筛选。返回 (选中的行, 统计)。

    多表文件必须先筛电表再差分: 表在文件里按块交错, 未筛时相邻行属于不同表,
    跨行差会算错量级 (实测能出现 -10 亿这种值)。所以界面不给"全部"选项。
    """
    rows = body
    stats = {"total": len(body), "after_meter": len(body), "after_dtype": len(body)}

    if meter_col is not None and meter_id not in (None, ""):
        rows = [r for r in rows if _same(r[meter_col], meter_id)]
    stats["after_meter"] = len(rows)

    if dtype_col is not None and dtype not in (None, "", "全部"):
        rows = [r for r in rows if _same(r[dtype_col], dtype)]
    stats["after_dtype"] = len(rows)
    return rows, stats


# ================= 倍率 =================
def resolve_multipliers(rows, ratio_col, ratio_const=None):
    """倍率来源: 优先逐行读综合倍率列, 否则用常数。返回 (每行倍率, 实际用的来源)。

    只有这一个倍率。60/step 那个换算系数不在这里, 它是量纲换算, 不是站点系数。
    """
    if ratio_col is not None:
        mults = []
        for r in rows:
            v = to_float(r[ratio_col])
            mults.append(None if v is None else v)
        return mults, "逐行读『综合倍率』列"
    if ratio_const:
        return [float(ratio_const) for _ in rows], "固定常数"
    return [None for _ in rows], "无倍率"


# ================= 差分 + 倍率 =================
def build_intervals(raw_rows, mults, anchor="left", power_factor=1.0):
    """后面减前面, 再 × 倍率 × power_factor。

    anchor='left' : 每行最后一格放跨日差(本行末->次日首); 表首格 = 第一个减完的数
    anchor='right': 每行第一格放跨日差(上行末->本行首); 整表首格 = 0

    只处理累计读数 (必须差分)。瞬时功率那种文件不走这条管线。
    """
    if not raw_rows:
        return []
    n = len(raw_rows[0])
    out = []
    for i, vals in enumerate(raw_rows):
        row = []
        for j in range(n):
            m = mults[i] if mults else 1.0
            if anchor == "right":
                if j == 0:
                    if i == 0:
                        row.append(0.0)              # 整表首格没有前驱
                        continue
                    a, b = raw_rows[i - 1][n - 1], vals[0]
                else:
                    a, b = vals[j - 1], vals[j]
            else:                                  # left
                a = vals[j]
                if j + 1 < n:
                    b = vals[j + 1]
                elif i + 1 < len(raw_rows):
                    b = raw_rows[i + 1][0]          # 次日00:00
                else:
                    b = None                       # 末行没有次日
            if a is None or b is None or m is None:
                row.append(None)
            else:
                row.append(round((b - a) * m * power_factor, 2))
        out.append(row)
    return out


def build_labels(n: int, step_min: int, anchor: str) -> list[str]:
    """区间标签: 必须命名自己这一格装的区间。

    left : 第j格装 [j, j+1] 区间, 末格装 23:xx-24:00
    right: 第j格装 [j-1, j] 区间, 首格装 23:xx-00:00
    """
    def fmt(m: int) -> str:
        return "%02d:00" % (m // 60) if step_min >= 60 else "%02d:%02d" % (m // 60, m % 60)

    labels = []
    for j in range(n):
        if anchor == "right":
            if j == 0:
                labels.append("%s-00:00" % fmt((n - 1) * step_min))
            else:
                labels.append("%s-%s" % (fmt((j - 1) * step_min), fmt(j * step_min)))
        else:
            if j + 1 < n:
                labels.append("%s-%s" % (fmt(j * step_min), fmt((j + 1) * step_min)))
            else:
                labels.append("%s-24:00" % fmt((n - 1) * step_min))
    return labels


# ================= 选代表日 =================
def mode_curve(out_rows):
    """逐时段取众数, 得到一条"最常出现"的负荷曲线。

    对每个时段, 统计所有天里这个值出现多少次, 取最多的那个。众数比平均值更
    贴近真实工况: 平均值是各天混合出来的虚构曲线, 众数通常就是某个真实出现
    过的运行状态 (比如某个固定档位)。

    返回 (众数列, 覆盖率列): 众数在该时段的有效值里占的比例 (0~1)。
    没有有效值的时段, 众数为 None。

    并列时 (多个值出现次数一样多) 取这些值的中位数 —— 任意挑一个会引入
    不该有的偏向, 中位数至少是对称的。
    """
    from collections import Counter

    modes = []
    shares = []
    for j in range(len(out_rows[0]) if out_rows else 0):
        vs = [r[j] for r in out_rows if r[j] is not None]
        if not vs:
            modes.append(None)
            shares.append(0.0)
            continue
        counts = Counter(vs)
        top_n = max(counts.values())
        tied = sorted(v for v, n in counts.items() if n == top_n)
        if len(tied) == 1:
            modes.append(tied[0])
        else:
            mid = len(tied) // 2
            modes.append(tied[mid] if len(tied) % 2 else
                         (tied[mid - 1] + tied[mid]) / 2)
        shares.append(top_n / len(vs))
    return modes, shares


# ================= 汇总 =================
def summarize(out_rows, labels, step_min=15, values_are_power=True):
    """返回 (每日汇总, 缺失数, 负值列表)。

    每日汇总 = (日电量kWh, 日均功率kW, 缺失格数); 有缺失则前两项为 None, 不做部分和。

    values_are_power: 每格是平均功率(kW) 还是区间电量(kWh)。
        kW 时 电量 = Σ格 × (step/60);  kWh 时 电量 = Σ格。
    """
    interval_hours = step_min / 60.0
    daily, miss_info, negs = [], [], []
    for i, row in enumerate(out_rows):
        present = [v for v in row if v is not None]
        miss = len(row) - len(present)
        miss_info.append(miss)
        if miss:
            daily.append((None, None, miss))
        else:
            kwh = sum(present) * (interval_hours if values_are_power else 1.0)
            daily.append((round(kwh, 2), round(kwh / 24, 2), 0))
        for j, v in enumerate(row):
            if v is not None and v < 0:
                negs.append((i, j, labels[j], v))
    return daily, miss_info, negs


# ================= 导出 =================
def build_workbook(job: dict) -> bytes:
    """拼成 xlsx: Sheet1 区间值, Sheet2 原始读数。

    区间值 sheet 右边带三列汇总: 日电量(kWh) / 日均功率(kW) / 缺失区间(个),
    不再单开一个"每日汇总" sheet。
    """
    import io

    import openpyxl

    info = job["info"]
    hdr = info["header"]
    rows = job["rows"]
    raw = job["raw"]
    out = job["out"]
    labels = job["labels"]
    daily = job["daily"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "区间值"
    for c in range(len(hdr)):
        ws.cell(1, c + 1, labels[c - job["t0"]] if job["t0"] <= c <= job["t1"] else hdr[c])
    ws.cell(1, len(hdr) + 1, "日电量(kWh)")
    ws.cell(1, len(hdr) + 2, "日均功率(kW)")
    ws.cell(1, len(hdr) + 3, "缺失区间(个)")
    for i, r in enumerate(rows):
        rr = i + 2
        for c in range(len(hdr)):
            if job["t0"] <= c <= job["t1"]:
                v = out[i][c - job["t0"]]
                ws.cell(rr, c + 1, "--" if v is None else v)
            else:
                ws.cell(rr, c + 1, r[c])
        kwh, kw, miss = daily[i]
        ws.cell(rr, len(hdr) + 1, "--" if kwh is None else kwh)
        ws.cell(rr, len(hdr) + 2, "--" if kw is None else kw)
        ws.cell(rr, len(hdr) + 3, miss)
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("原始读数")
    for c in range(len(hdr)):
        ws2.cell(1, c + 1, hdr[c])
    for i, r in enumerate(rows):
        rr = i + 2
        for c in range(len(hdr)):
            if job["t0"] <= c <= job["t1"]:
                v = raw[i][c - job["t0"]]
                ws2.cell(rr, c + 1, "--" if v is None else v)
            else:
                ws2.cell(rr, c + 1, r[c])
    ws2.freeze_panes = "A2"

    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()
