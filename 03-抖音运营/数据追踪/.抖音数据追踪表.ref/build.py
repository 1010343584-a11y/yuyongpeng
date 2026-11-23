"""
抖音数据追踪表.xlsx 生成脚本
用途：抖音真人出镜账号（心灵成长赛道，20-40岁女性受众）的每日运营数据记录
结构：每日数据（明细） + 周复盘（周汇总） + 月度汇总（月汇总）
"""
import sys

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


# ===== 配色常量（商务蓝）=====
XL_PRIMARY = xl_color("#4472C4")
XL_PRIMARY_LIGHT = xl_color("#D9E2F3")
XL_BORDER_COLOR = xl_color("#BFBFBF")
XL_DARK = xl_color("#2F5597")
XL_GREEN = xl_color("#70AD47")
XL_GOLD = xl_color("#FFC700")
XL_SECTION = xl_color("#E8EDF5")

# ===== 字体 =====
FONT_HEADER = Font(name="微软雅黑", size=10.5, bold=True, color="FFFFFFFF")
FONT_BODY = Font(name="微软雅黑", size=10, color="FF000000")
FONT_TITLE = Font(name="微软雅黑", size=13, bold=True, color=XL_DARK)
FONT_SECTION = Font(name="微软雅黑", size=9.5, color=xl_color("#5A5A5A"))
FONT_TOTAL = Font(name="微软雅黑", size=10, bold=True, color=XL_DARK)

# ===== 填充 / 边框 =====
FILL_HEADER = PatternFill(start_color=XL_PRIMARY, end_color=XL_PRIMARY, fill_type="solid")
FILL_INPUT = PatternFill(start_color=XL_PRIMARY_LIGHT, end_color=XL_PRIMARY_LIGHT, fill_type="solid")
FILL_SECTION = PatternFill(start_color=XL_SECTION, end_color=XL_SECTION, fill_type="solid")
FILL_TOTAL = PatternFill(start_color=XL_PRIMARY_LIGHT, end_color=XL_PRIMARY_LIGHT, fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style="thin", color=XL_BORDER_COLOR),
    right=Side(style="thin", color=XL_BORDER_COLOR),
    top=Side(style="thin", color=XL_BORDER_COLOR),
    bottom=Side(style="thin", color=XL_BORDER_COLOR),
)
BORDER_TOP_DOUBLE = Border(top=Side(style="double", color=XL_DARK))

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")

FMT_INT = "#,##0"
FMT_PCT = "0%"
FMT_DATE = "YYYY-MM-DD"


def put_header(ws, row, headers):
    """headers: list of (名称, 列宽)"""
    for idx, (name, width) in enumerate(headers, start=1):
        c = ws.cell(row=row, column=idx, value=name)
        c.font = FONT_HEADER
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_THIN
        ws.column_dimensions[get_column_letter(idx)].width = width


def put_section(ws, row, text, ncols):
    """整行说明条（跨列合并、浅底）"""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = FONT_SECTION
    c.fill = FILL_SECTION
    c.alignment = ALIGN_LEFT
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = FILL_SECTION


def put_cell(ws, row, col, value, *, fill=FILL_INPUT, align=ALIGN_CENTER,
             num_fmt=None, font=FONT_BODY):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.fill = fill
    c.border = BORDER_THIN
    c.alignment = align
    if num_fmt:
        c.number_format = num_fmt
    return c


wb = openpyxl.Workbook()
wb.properties.title = "抖音数据追踪表"

# =====================================================================
# Sheet 1：每日数据
# =====================================================================
ws1 = wb.active
ws1.title = "每日数据"
ws1.sheet_properties.tabColor = XL_PRIMARY

put_section(ws1, 1, "📊  每日数据记录 ｜ 抖音真人出镜 · 心灵成长赛道（20-40岁女性）", 12)

put_header(ws1, 3, [
    ("序号", 6), ("日期", 13), ("视频标题 / 主题", 38), ("内容形式", 13),
    ("播放量", 12), ("点赞", 10), ("收藏", 10), ("评论", 10),
    ("转发", 10), ("净增粉丝", 11), ("完播率", 10), ("备注", 32),
])

daily_samples = [
    (1, "2026-11-21", "厉娜体情绪收尾：成年人都需要一个情绪出口", "口播",
     12500, 1480, 356, 127, 34, 72, 0.42, "示例行，可删除"),
    (2, "2026-11-22", "王阳明心学：别人随口一句话，为什么能难受一整天", "口播",
     18300, 2150, 620, 198, 56, 98, 0.48, "示例行，可删除"),
    (3, "2026-11-23", "累了就出去玩，不是贪玩是好好爱自己", "边走边说",
     9600, 1050, 280, 89, 23, 45, 0.38, "示例行，可删除"),
]

row_start = 4
for offset, (idx, date_v, title, form, plays, likes, col, cmt, share, fans, rate, note) in enumerate(daily_samples):
    r = row_start + offset
    put_cell(ws1, r, 1, idx, num_fmt=FMT_INT)
    put_cell(ws1, r, 2, date_v, num_fmt=FMT_DATE)
    put_cell(ws1, r, 3, title, align=ALIGN_LEFT_WRAP)
    put_cell(ws1, r, 4, form)
    put_cell(ws1, r, 5, plays, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 6, likes, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 7, col, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 8, cmt, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 9, share, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 10, fans, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws1, r, 11, rate, num_fmt=FMT_PCT)
    put_cell(ws1, r, 12, note, align=ALIGN_LEFT_WRAP)

last_daily = row_start + len(daily_samples) - 1
put_section(ws1, last_daily + 2,
            "⚠️ 以上为示例数据，正式使用时可整体替换；每日发完视频后当天填写，数据越连续越能看出规律。", 12)

ws1.freeze_panes = "A4"
ws1.auto_filter.ref = f"A3:L{last_daily}"

# 内容形式下拉
dv_form = DataValidation(type="list", formula1='"口播,边走边说,图文轮播,对话式"', allow_blank=True)
ws1.add_data_validation(dv_form)
dv_form.add(f"D{row_start}:D{last_daily + 60}")

ws1.print_options.horizontalCentered = True

# =====================================================================
# Sheet 2：周复盘
# =====================================================================
ws2 = wb.create_sheet("周复盘")
ws2.sheet_properties.tabColor = XL_GREEN

put_section(ws2, 1, "📈  每周复盘 ｜ 追踪规律 · 迭代优化（每周末用 15 分钟填一次）", 11)

put_header(ws2, 3, [
    ("序号", 6), ("周次", 8), ("起止日期", 21), ("发布条数", 10),
    ("累计播放", 12), ("累计点赞", 11), ("累计收藏", 11), ("累计评论", 11),
    ("净增粉丝", 11), ("本周最佳视频", 30), ("本周问题 / 下周改进方向", 36),
])

weekly_samples = [
    (1, "W1", "2026-11-17 ~ 11-23", 4, 68400, 8180, 1926, 645, 312,
     "「王阳明：别人一句话为什么能难受一整天」", "周三发布时段偏早，流量未跑起来；下周改到 20:30 后发布。"),
    (2, "W2", "2026-11-24 ~ 11-30", None, None, None, None, None, None, "", ""),
    (3, "W3", "2026-12-01 ~ 12-07", None, None, None, None, None, None, "", ""),
    (4, "W4", "2026-12-08 ~ 12-14", None, None, None, None, None, None, "", ""),
]

row_start = 4
for offset, (idx, week, span, count, plays, likes, col, cmt, fans, best, issue) in enumerate(weekly_samples):
    r = row_start + offset
    put_cell(ws2, r, 1, idx, num_fmt=FMT_INT)
    put_cell(ws2, r, 2, week)
    put_cell(ws2, r, 3, span, align=ALIGN_LEFT)
    put_cell(ws2, r, 4, count if count is not None else None, num_fmt=FMT_INT)
    put_cell(ws2, r, 5, plays if plays is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws2, r, 6, likes if likes is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws2, r, 7, col if col is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws2, r, 8, cmt if cmt is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws2, r, 9, fans if fans is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws2, r, 10, best if best else None, align=ALIGN_LEFT_WRAP)
    put_cell(ws2, r, 11, issue if issue else None, align=ALIGN_LEFT_WRAP)

last_weekly = row_start + len(weekly_samples) - 1
put_section(ws2, last_weekly + 2,
            "💡 复盘三问：本周哪条最好、为什么好？哪条最差、问题在哪？下周只改一个变量。", 11)

ws2.freeze_panes = "A4"
ws2.auto_filter.ref = f"A3:K{last_weekly}"

# =====================================================================
# Sheet 3：月度汇总
# =====================================================================
ws3 = wb.create_sheet("月度汇总")
ws3.sheet_properties.tabColor = XL_GOLD

put_section(ws3, 1, "📅  月度汇总 ｜ 一眼看趋势（每月月底填写一次）", 8)

put_header(ws3, 3, [
    ("序号", 6), ("月份", 12), ("发布条数", 11), ("累计播放", 14),
    ("累计点赞", 12), ("累计收藏", 12), ("净增粉丝", 12), ("备注", 32),
])

monthly_samples = [
    (1, "2026-11", 16, 245800, 28600, 7200, 1450, "首月，测试不同选题类型"),
    (2, "2026-12", None, None, None, None, None, ""),
    (3, "2027-01", None, None, None, None, None, ""),
    (4, "2027-02", None, None, None, None, None, ""),
]

row_start = 4
for offset, (idx, month, count, plays, likes, col, fans, note) in enumerate(monthly_samples):
    r = row_start + offset
    put_cell(ws3, r, 1, idx, num_fmt=FMT_INT)
    put_cell(ws3, r, 2, month)
    put_cell(ws3, r, 3, count if count is not None else None, num_fmt=FMT_INT)
    put_cell(ws3, r, 4, plays if plays is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws3, r, 5, likes if likes is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws3, r, 6, col if col is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws3, r, 7, fans if fans is not None else None, align=ALIGN_RIGHT, num_fmt=FMT_INT)
    put_cell(ws3, r, 8, note if note else None, align=ALIGN_LEFT_WRAP)

last_monthly = row_start + len(monthly_samples) - 1
total_row = last_monthly + 1

# 合计行：引用上方数据区
put_cell(ws3, total_row, 1, None, fill=FILL_TOTAL)
put_cell(ws3, total_row, 2, "合计", fill=FILL_TOTAL, align=ALIGN_CENTER, font=FONT_TOTAL)
for col_letter in ["C", "D", "E", "F", "G"]:
    c = ws3[f"{col_letter}{total_row}"]
    c.value = f"=SUM({col_letter}{row_start}:{col_letter}{last_monthly})"
    c.font = FONT_TOTAL
    c.fill = FILL_TOTAL
    c.alignment = ALIGN_RIGHT
    c.number_format = FMT_INT
    c.border = BORDER_THIN
put_cell(ws3, total_row, 8, None, fill=FILL_TOTAL)

for col in range(1, 9):
    ws3.cell(row=total_row, column=col).border = Border(
        left=Side(style="thin", color=XL_BORDER_COLOR),
        right=Side(style="thin", color=XL_BORDER_COLOR),
        top=Side(style="double", color=XL_DARK),
        bottom=Side(style="thin", color=XL_BORDER_COLOR),
    )

put_section(ws3, total_row + 2,
            "📌 累计数据可从「每日数据」按月份手工汇总，也可直接填入本表。", 8)

ws3.freeze_panes = "A4"
ws3.auto_filter.ref = f"A3:H{last_monthly}"

# =====================================================================
output_path = "/workspace/知识库/03-抖音运营/数据追踪/抖音数据追踪表.xlsx"
wb.save(output_path)
print(f"OK saved: {output_path}")
print(f"sheets: {wb.sheetnames}")
