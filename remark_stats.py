"""
衣杆报表统计 v3
==================================================
运行方式：python3 remark_stats.py
  - 弹出文件选择窗口 → 选择 xlsx / csv 源表
  - 逐条解析"商家备注"，每条备注 → 结果表一行
  - 输出列（严格对齐目标表 10 列）：
      A:款式   B:颜色   C:米数总计   D:底座总数   E:全包围底座总数
      F:顶装半通总数   G:顶装全通总数   H:顶装转角总数
      I:中托总数   J:螺丝总数
      + 核对列：原始备注
  - 输出文件：源文件名（统计明细表）.xlsx（存在源文件同目录）
"""

import os, re, sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════
# 1. 文件选择
# ══════════════════════════════════════════════

def pick_file():
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="选择源表文件（xlsx / csv）",
        initialdir=os.path.expanduser("~/Desktop"),
        filetypes=[("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv"),
                   ("所有支持格式", "*.xlsx *.csv")],
    )
    root.destroy()
    return path

def show_error(msg):
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    messagebox.showerror("错误", msg); root.destroy()

def show_info(msg):
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    messagebox.showinfo("完成", msg); root.destroy()

SOURCE = pick_file()
if not SOURCE:
    print("未选择文件，退出。"); sys.exit(0)

ext = os.path.splitext(SOURCE)[1].lower()
if ext not in (".xlsx", ".csv"):
    show_error(f"不支持的格式：{ext}\n请选择 .xlsx 或 .csv 文件。"); sys.exit(1)

source_dir  = os.path.dirname(SOURCE)
source_base = os.path.splitext(os.path.basename(SOURCE))[0]
OUTPUT      = os.path.join(source_dir, f"{source_base}（统计明细表）.xlsx")
print(f"📂 源文件：{SOURCE}")
print(f"📝 输出  ：{OUTPUT}")


# ══════════════════════════════════════════════
# 2. 读取源表
# ══════════════════════════════════════════════

df_src = pd.read_csv(SOURCE, encoding="utf-8-sig") if ext == ".csv" \
         else pd.read_excel(SOURCE)
df_src.columns = df_src.columns.str.strip()
if "商家备注" not in df_src.columns:
    show_error(f"源表中未找到[商家备注]列\n当前列：{list(df_src.columns)}"); sys.exit(1)

remarks = df_src["商家备注"].dropna()
remarks = remarks[remarks.astype(str).str.strip() != ""]
print(f"✅ 非空备注：{len(remarks)} 条")


# ══════════════════════════════════════════════
# 3. 解析逻辑
# ══════════════════════════════════════════════

# ── 款式关键词 ──
def extract_style(text):
    # 先做容错匹配（typo 变体），再统一映射回标准名
    for kw in ["X3150", "M335", "6号威法", "6号威发", "2mm"]:
        if kw in text:
            return kw
    # X315（漏了0）→ X3150
    if re.search(r"X315[^0]", text):
        return "X3150"
    return ""


# ── 颜色映射 ──
COLOR_RULES = [
    (r"瓷白|白色|白", "白"),
    (r"暮云灰|玄铁灰|灰色|灰", "灰"),
    (r"雅黑|路虎黑|黑色|黑", "黑"),
    (r"香槟|金色|金", "香槟/金"),
]
def extract_color(text):
    for pat, label in COLOR_RULES:
        if re.search(pat, text):
            return label
    return "其他"


# ── 辅助解析函数 ──

SKIP_RE = [r"^共\d+根$", r"^只要杆子", r"^纯杆子", r"^补发", r"^\d+$",
           r"^深度\d+$", r"^-\d+", r"^加点", r"^升级"]

def should_skip(tok):
    for pat in SKIP_RE:
        if re.match(pat, tok): return True
    if re.match(r"^[\u4e00-\u9fff]+$", tok): return True
    return False

def is_size_item(tok):
    return bool(re.match(r"^\d{2,4}\s*[*×xX]\s*\d+\s*根?$", tok))

def parse_qty(tok):
    """返回 (名称, 数量)，支持 底座*2 / 底座2个 / 螺丝135 三种格式"""
    # 格式1：名称*数量（可含可选单位）
    m = re.search(r"^(.+?)\s*[*×xX]\s*(\d+)\s*[个根]?$", tok)
    if m: return m.group(1).strip(), int(m.group(2))
    # 格式2：名称+数字+单位（如 底座12个、半通6个）
    m = re.search(r"^(.+?)(\d+)\s*[个根]?$", tok)
    if m: return m.group(1).strip(), int(m.group(2))
    return tok, 0

def split_tokens(text):
    text = text.replace("，", ",").replace("、", ",").replace("；", ",")
    return [t.strip() for t in re.split(r"[,\s]+", text) if t.strip()]

def normalize_size(tok):
    m = re.match(r"^(\d{2,4})\s*[*×xX]\s*(\d+)\s*根?$", tok)
    if m: return int(m.group(1)), int(m.group(2))
    return None, 0


def parse_remark(raw):
    raw = str(raw).strip()

    # ── 拆分款式与明细 ──────────────────────────
    if "：" in raw:
        spec_part, detail = raw.split("：", 1)
    elif ":" in raw:
        spec_part, detail = raw.split(":", 1)
    elif "，" in raw and is_size_item(raw.split("，", 1)[0]) is False:
        spec_part, detail = raw.split("，", 1)
    else:
        spec_part, detail = "", raw

    spec_part = spec_part.strip()

    # ── 折扣前缀（如 "-5："）→ 从 detail 再拆一次 ──
    if re.match(r"^[-\d]+$", spec_part):
        for sep in ["：", ":", "，"]:
            if sep in detail:
                spec_part, detail = detail.split(sep, 1)
                break
        spec_part = spec_part.strip()

    # ── 处理「补发」──
    # 纯补发单：整条备注以"补发"开头
    if spec_part.startswith("补发"):
        spec_part = spec_part.replace("补发", "").strip("：: ")
        # 如果 spec_part 被清空，从 detail 重新拆
        if not spec_part:
            for sep in ["：", ":", "，"]:
                if sep in detail:
                    spec_part, detail = detail.split(sep, 1)
                    break
            spec_part = spec_part.strip()
    else:
        # 嵌套补发段落在 detail 中 → 截去（避免重复统计）
        if "补发" in detail:
            detail = re.split(r"补发", detail)[0]

    # ── 识别款式、颜色 ──────────────────────────
    style = extract_style(spec_part)
    if not style:
        style = extract_style(raw)
    if style == "6号威发":
        style = "6号威法"
    color = extract_color(spec_part.replace(style, "") if style else spec_part)
    if color == "其他" and style:
        color = extract_color(raw.replace(style, "", 1) if style in raw else raw)

    tokens = split_tokens(detail)

    # ── 统计变量 ──
    total_mm   = 0.0    # 总毫米数（之后 /1000 → 米数）
    base_total = 0
    quanbao    = 0
    bantong    = 0
    quantong   = 0
    zhuanjiao  = 0
    zhongtuo   = 0
    luosi      = 0

    for tok in tokens:
        if should_skip(tok):
            continue

        # 尺寸项 → 累加米数
        if is_size_item(tok):
            length, qty = normalize_size(tok)
            total_mm += length * qty
            continue

        # ── 底座项（优先处理）──
        if "底座" in tok:
            name, qty = parse_qty(tok)
            base_total += qty
            if "全包围" in tok:
                quanbao += qty
            continue

        # ── 配件项（不含"底座"）──
        name, qty = parse_qty(tok)
        if qty == 0:
            continue   # 无法解析数量则跳过

        n = name
        if "螺丝" in n:
            luosi += qty
        else:
            # 顶装配件：半通/全通/转角 可能共存于一个 token
            if "半通" in n:
                bantong += qty
            if "全通" in n:
                quantong += qty
            if "转角" in n:
                zhuanjiao += qty
            if "中托" in n:
                zhongtuo += qty

    return {
        "款式":            style,
        "颜色":            color,
        "米数总计":        round(total_mm / 1000, 2) if total_mm > 0 else "",
        "底座总数":        base_total if base_total > 0 else "",
        "全包围底座总数":   quanbao if quanbao > 0 else "",
        "顶装半通总数":     bantong if bantong > 0 else "",
        "顶装全通总数":     quantong if quantong > 0 else "",
        "顶装转角总数":     zhuanjiao if zhuanjiao > 0 else "",
        "中托总数":        zhongtuo if zhongtuo > 0 else "",
        "螺丝总数":        luosi if luosi > 0 else "",
        "_原始备注":        raw,
    }


# ══════════════════════════════════════════════
# 4. 批量解析
# ══════════════════════════════════════════════

OUT_COLS = [
    "款式", "颜色", "米数总计",
    "底座总数", "全包围底座总数",
    "顶装半通总数", "顶装全通总数", "顶装转角总数",
    "中托总数", "螺丝总数",
    "_原始备注",
]

rows = [parse_remark(r) for r in remarks]
df_out = pd.DataFrame(rows, columns=OUT_COLS)
print(f"✅ 解析完成，共 {len(df_out)} 行")


# ══════════════════════════════════════════════
# 5. 写 Excel
# ══════════════════════════════════════════════

DISPLAY_COLS = [
    "款式", "颜色", "米数总计",
    "底座总数", "全包围底座总数",
    "顶装半通总数", "顶装全通总数", "顶装转角总数",
    "中托总数", "螺丝总数",
    "原始备注（核对用）",
]

HDR_BG = "1F4E79"; SUB_BG = "2E75B6"
ODD_BG = "EBF3FB"; EVEN_BG = "FFFFFF"
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN   = Alignment(horizontal="left",   vertical="center")
RIGHT_ALIGN  = Alignment(horizontal="right",  vertical="center")
def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

wb = Workbook()
ws = wb.active
ws.title = "统计明细"
ws.freeze_panes = "A3"

ncols = len(DISPLAY_COLS)
col_letter = get_column_letter(ncols)

# 标题行
ws.merge_cells(f"A1:{col_letter}1")
tc = ws.cell(1, 1, f"{source_base} — 商家备注统计明细")
tc.font      = Font(bold=True, size=14, color="FFFFFF")
tc.fill      = PatternFill("solid", start_color=HDR_BG)
tc.alignment = CENTER_ALIGN
ws.row_dimensions[1].height = 30

# 表头行
for ci, col in enumerate(DISPLAY_COLS, 1):
    c = ws.cell(2, ci, col)
    c.font      = Font(bold=True, size=10, color="FFFFFF")
    c.fill      = PatternFill("solid", start_color=SUB_BG)
    c.alignment = CENTER_ALIGN
    c.border    = thin_border()
ws.row_dimensions[2].height = 26

# 数据行
NUM_COLS = {"米数总计", "底座总数", "全包围底座总数",
            "顶装半通总数", "顶装全通总数", "顶装转角总数",
            "中托总数", "螺丝总数"}
CENTER_COLS = {"款式", "颜色"}

for ri, row in df_out.iterrows():
    er = ri + 3
    bg = ODD_BG if ri % 2 == 0 else EVEN_BG
    for ci, key in enumerate(OUT_COLS, 1):
        val = row[key]
        if pd.isna(val): val = ""
        c = ws.cell(er, ci, val)
        c.font   = Font(size=10)
        c.fill   = PatternFill("solid", start_color=bg)
        c.border = thin_border()
        display_key = DISPLAY_COLS[ci - 1]
        if display_key in NUM_COLS:
            c.alignment = RIGHT_ALIGN
        elif display_key in CENTER_COLS:
            c.alignment = CENTER_ALIGN
        else:
            c.alignment = LEFT_ALIGN

# 列宽
COL_WIDTHS = {
    "款式": 10, "颜色": 8, "米数总计": 10,
    "底座总数": 10, "全包围底座总数": 14,
    "顶装半通总数": 12, "顶装全通总数": 12, "顶装转角总数": 12,
    "中托总数": 10, "螺丝总数": 10,
    "原始备注（核对用）": 55,
}
for ci, key in enumerate(DISPLAY_COLS, 1):
    ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(key, 14)

wb.save(OUTPUT)
print(f"\n🎉 输出：{OUTPUT}")
show_info(f"统计完成！\n共 {len(df_out)} 行明细\n\n输出文件：\n{OUTPUT}")
