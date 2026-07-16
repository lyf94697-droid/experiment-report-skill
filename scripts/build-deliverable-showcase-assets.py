from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Sequence

from PIL import Image, ImageDraw, ImageFont


WIDTH = 1600
HEIGHT = 900
WHITE = "#FFFFFF"
INK = "#18212F"
MUTED = "#667085"
BLUE = "#245C7A"
LIGHT_BLUE = "#EAF2F7"
TEAL = "#0F766E"
LIGHT_TEAL = "#E7F4F2"
GREEN = "#2E7D32"
LIGHT_GREEN = "#EAF5EA"
ORANGE = "#C65D21"
LIGHT_ORANGE = "#FFF1E8"
RED = "#B42318"
LIGHT_RED = "#FDECEC"
GRID = "#D0D5DD"
SOFT = "#F7F8FA"
TERMINAL = "#111827"
TERMINAL_TEXT = "#E5E7EB"

FONT_REGULAR_PATH = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_BOLD_PATH = Path(r"C:\Windows\Fonts\msyhbd.ttc")
FONT_MONO_PATH = Path(r"C:\Windows\Fonts\consola.ttf")
FONT_MONO_BOLD_PATH = Path(r"C:\Windows\Fonts\consolab.ttf")


def font(size: int, *, bold: bool = False, mono: bool = False) -> ImageFont.FreeTypeFont:
    if mono:
        path = FONT_MONO_BOLD_PATH if bold else FONT_MONO_PATH
    else:
        path = FONT_BOLD_PATH if bold else FONT_REGULAR_PATH
    return ImageFont.truetype(str(path), size)


def rounded(draw: ImageDraw.ImageDraw, box, *, fill, outline=None, width=2, radius=18):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def text_size(draw: ImageDraw.ImageDraw, text: str, text_font) -> tuple[int, int]:
    bbox = draw.textbbox((0, 0), text, font=text_font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def centered_text(draw: ImageDraw.ImageDraw, box, text: str, text_font, fill=INK):
    x1, y1, x2, y2 = box
    tw, th = text_size(draw, text, text_font)
    draw.text(((x1 + x2 - tw) / 2, (y1 + y2 - th) / 2 - 2), text, font=text_font, fill=fill)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, text_font, max_width: int) -> list[str]:
    lines: list[str] = []
    for source_line in text.splitlines() or [""]:
        if not source_line:
            lines.append("")
            continue
        current = ""
        for character in source_line:
            candidate = current + character
            if current and text_size(draw, candidate, text_font)[0] > max_width:
                lines.append(current.rstrip())
                current = character
            else:
                current = candidate
        if current:
            lines.append(current.rstrip())
    return lines


def multiline(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int],
    text: str,
    text_font,
    *,
    fill=INK,
    max_width: int,
    line_gap: int = 12,
) -> int:
    x, y = xy
    line_height = text_size(draw, "示例Ag", text_font)[1]
    for line in wrap_text(draw, text, text_font, max_width):
        draw.text((x, y), line, font=text_font, fill=fill)
        y += line_height + line_gap
    return y


def arrow(draw: ImageDraw.ImageDraw, start, end, *, fill=BLUE, width=5):
    draw.line([start, end], fill=fill, width=width)
    x1, y1 = start
    x2, y2 = end
    if abs(x2 - x1) >= abs(y2 - y1):
        direction = 1 if x2 > x1 else -1
        points = [(x2, y2), (x2 - 18 * direction, y2 - 10), (x2 - 18 * direction, y2 + 10)]
    else:
        direction = 1 if y2 > y1 else -1
        points = [(x2, y2), (x2 - 10, y2 - 18 * direction), (x2 + 10, y2 - 18 * direction)]
    draw.polygon(points, fill=fill)


def new_canvas(title: str, subtitle: str, *, accent=BLUE) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (WIDTH, HEIGHT), WHITE)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, WIDTH, 96), fill=accent)
    draw.text((70, 25), title, font=font(40, bold=True), fill=WHITE)
    draw.text((70, 117), subtitle, font=font(24), fill=MUTED)
    draw.line((70, 164, WIDTH - 70, 164), fill=GRID, width=2)
    return image, draw


def save(image: Image.Image, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, format="PNG", optimize=True)


def terminal_image(title: str, command_lines: Sequence[tuple[str, str]], path: Path):
    image = Image.new("RGB", (WIDTH, HEIGHT), "#E8EDF3")
    draw = ImageDraw.Draw(image)
    rounded(draw, (70, 60, WIDTH - 70, HEIGHT - 60), fill=TERMINAL, outline="#344054", width=2, radius=22)
    draw.rectangle((70, 60, WIDTH - 70, 125), fill="#273142")
    for index, color in enumerate(("#FF5F57", "#FEBB2E", "#28C840")):
        draw.ellipse((100 + index * 36, 82, 118 + index * 36, 100), fill=color)
    title_font = font(26, bold=True)
    title_width, _ = text_size(draw, title, title_font)
    draw.text(((WIDTH - title_width) / 2, 75), title, font=title_font, fill="#F9FAFB")

    y = 158
    mono = font(27, mono=True)
    mono_bold = font(27, bold=True, mono=True)
    for text, color in command_lines:
        current_font = mono_bold if color in ("#7DD3FC", "#86EFAC") else mono
        draw.text((105, y), text, font=current_font, fill=color)
        y += 43
    save(image, path)


def draw_table(
    draw: ImageDraw.ImageDraw,
    box,
    headers: Sequence[str],
    rows: Sequence[Sequence[str]],
    *,
    widths: Sequence[float] | None = None,
    header_fill=LIGHT_BLUE,
    row_fill=WHITE,
    header_color=INK,
    text_color=INK,
    font_size=23,
):
    x1, y1, x2, y2 = box
    columns = len(headers)
    if widths is None:
        widths = [1 / columns] * columns
    total = sum(widths)
    normalized = [value / total for value in widths]
    xs = [x1]
    for value in normalized:
        xs.append(xs[-1] + (x2 - x1) * value)
    row_count = len(rows) + 1
    row_height = (y2 - y1) / row_count
    for row_index in range(row_count):
        top = y1 + row_index * row_height
        bottom = top + row_height
        fill = header_fill if row_index == 0 else (row_fill if row_index % 2 else SOFT)
        draw.rectangle((x1, top, x2, bottom), fill=fill)
        for column in range(columns):
            cell = (xs[column], top, xs[column + 1], bottom)
            draw.rectangle(cell, outline=GRID, width=2)
            value = headers[column] if row_index == 0 else str(rows[row_index - 1][column])
            value_font = font(font_size, bold=row_index == 0)
            centered_text(draw, cell, value, value_font, fill=header_color if row_index == 0 else text_color)


def network_topology(path: Path):
    image, draw = new_canvas("局域网实验拓扑与地址规划", "两台虚拟主机位于同一 /24 网段，通过仅主机虚拟交换机互联")
    host_y = 300
    host_w, host_h = 360, 280
    left = (100, host_y, 100 + host_w, host_y + host_h)
    right = (WIDTH - 100 - host_w, host_y, WIDTH - 100, host_y + host_h)
    switch = (620, 360, 980, 525)
    for box, title, ip, mac in (
        (left, "主机 A / Windows Server", "192.168.10.11 /24", "00-0C-29-11-0A-01"),
        (right, "主机 B / Windows Server", "192.168.10.12 /24", "00-0C-29-12-0B-02"),
    ):
        rounded(draw, box, fill=SOFT, outline=BLUE, width=4)
        draw.rectangle((box[0], box[1], box[2], box[1] + 64), fill=LIGHT_BLUE)
        centered_text(draw, (box[0], box[1], box[2], box[1] + 64), title, font(25, bold=True), fill=BLUE)
        draw.text((box[0] + 34, box[1] + 102), f"IPv4：{ip}", font=font(24), fill=INK)
        draw.text((box[0] + 34, box[1] + 154), "掩码：255.255.255.0", font=font(24), fill=INK)
        draw.text((box[0] + 34, box[1] + 206), f"MAC：{mac}", font=font(20, mono=True), fill=MUTED)
    rounded(draw, switch, fill=LIGHT_TEAL, outline=TEAL, width=4)
    centered_text(draw, (620, 360, 980, 440), "VMnet2 仅主机交换机", font(28, bold=True), fill=TEAL)
    centered_text(draw, (620, 440, 980, 515), "网段 192.168.10.0 /24", font(24), fill=INK)
    arrow(draw, (460, 440), (620, 440), fill=TEAL)
    arrow(draw, (1140, 440), (980, 440), fill=TEAL)
    rounded(draw, (270, 660, 1330, 795), fill=WHITE, outline=GRID, width=2)
    draw.text((315, 690), "验证链路：", font=font(26, bold=True), fill=INK)
    draw.text((480, 690), "ipconfig → ping 192.168.10.12 → arp -a", font=font(26, mono=True), fill=BLUE)
    draw.text((315, 748), "判定标准：4 次 ICMP 回复、0% 丢包，并在 ARP 缓存中出现对端动态条目。", font=font(23), fill=MUTED)
    save(image, path)


def network_diagnosis(path: Path):
    image, draw = new_canvas("连通性故障定位流程", "按“本机协议栈 → 地址配置 → 二层解析 → 安全策略”的顺序排查", accent=TEAL)
    steps = [
        ("1", "检查协议栈", "ping 127.0.0.1", LIGHT_BLUE, BLUE),
        ("2", "检查本机地址", "ipconfig /all", LIGHT_TEAL, TEAL),
        ("3", "核对同一网段", "IP 与掩码组合", LIGHT_ORANGE, ORANGE),
        ("4", "检查邻居解析", "arp -a / arp -d *", LIGHT_GREEN, GREEN),
        ("5", "检查防火墙", "允许 ICMP Echo", LIGHT_RED, RED),
    ]
    x = 85
    for index, (number, title, command, fill, accent) in enumerate(steps):
        box = (x, 285, x + 250, 545)
        rounded(draw, box, fill=fill, outline=accent, width=3)
        draw.ellipse((x + 88, 315, x + 162, 389), fill=accent)
        centered_text(draw, (x + 88, 315, x + 162, 389), number, font(30, bold=True), fill=WHITE)
        centered_text(draw, (x + 20, 410, x + 230, 458), title, font(24, bold=True), fill=accent)
        centered_text(draw, (x + 15, 472, x + 235, 526), command, font(20, mono=True), fill=INK)
        if index < len(steps) - 1:
            arrow(draw, (x + 250, 415), (x + 290, 415), fill=MUTED, width=4)
        x += 290
    rounded(draw, (180, 645, 1420, 790), fill=SOFT, outline=GRID, width=2)
    draw.text((220, 675), "本次异常复现：", font=font(25, bold=True), fill=INK)
    draw.text((420, 675), "将主机 B 掩码误设为 255.255.0.0 后，地址规划不一致；恢复 /24 并清理 ARP 缓存后通信恢复。", font=font(23), fill=MUTED)
    draw.text((220, 730), "排查原则：", font=font(25, bold=True), fill=INK)
    draw.text((350, 730), "每一步都记录命令、现象和结论，避免只写“重启后正常”。", font=font(23), fill=MUTED)
    save(image, path)


def scheduling_flow(path: Path):
    image, draw = new_canvas("进程调度模拟程序流程", "同一组进程数据分别进入 FCFS、SJF 和 RR 调度器，并统一统计指标")
    boxes = [
        ((100, 310, 340, 475), "读取进程集合", "PID / 到达时间 / 服务时间"),
        ((470, 310, 710, 475), "维护就绪队列", "处理空闲时段与新到达进程"),
        ((840, 245, 1110, 390), "选择调度策略", "FCFS · SJF · RR(q=2)"),
        ((840, 470, 1110, 615), "推进系统时钟", "记录开始、完成与剩余时间"),
        ((1240, 345, 1500, 520), "计算评价指标", "等待 / 周转 / 响应 / 切换次数"),
    ]
    for box, title, subtitle in boxes:
        rounded(draw, box, fill=SOFT, outline=BLUE, width=3)
        centered_text(draw, (box[0] + 10, box[1] + 20, box[2] - 10, box[1] + 78), title, font(25, bold=True), fill=BLUE)
        multiline(draw, (box[0] + 24, box[1] + 88), subtitle, font(20), fill=MUTED, max_width=box[2] - box[0] - 48, line_gap=8)
    arrow(draw, (340, 392), (470, 392), fill=BLUE)
    arrow(draw, (710, 392), (840, 320), fill=BLUE)
    arrow(draw, (975, 390), (975, 470), fill=BLUE)
    arrow(draw, (1110, 542), (1240, 432), fill=BLUE)
    arrow(draw, (840, 542), (710, 475), fill=MUTED)
    draw.text((570, 650), "循环条件：未完成进程数 > 0", font=font(26, bold=True), fill=INK)
    draw.text((570, 706), "关键校验：就绪队列为空时，将时钟推进到下一到达时刻。", font=font(23), fill=MUTED)
    save(image, path)


def scheduling_gantt(path: Path):
    image, draw = new_canvas("三种调度算法甘特图", "测试集：P1(0,8)、P2(1,4)、P3(2,9)、P4(3,5)、P5(4,2)")
    schedules = {
        "FCFS": [(0, 8, "P1"), (8, 12, "P2"), (12, 21, "P3"), (21, 26, "P4"), (26, 28, "P5")],
        "SJF": [(0, 8, "P1"), (8, 10, "P5"), (10, 14, "P2"), (14, 19, "P4"), (19, 28, "P3")],
        "RR(q=2)": [
            (0, 2, "P1"), (2, 4, "P2"), (4, 6, "P3"), (6, 8, "P1"), (8, 10, "P4"),
            (10, 12, "P5"), (12, 14, "P2"), (14, 16, "P3"), (16, 18, "P1"),
            (18, 20, "P4"), (20, 22, "P3"), (22, 24, "P1"), (24, 25, "P4"),
            (25, 27, "P3"), (27, 28, "P3"),
        ],
    }
    colors = {"P1": "#2F6B8A", "P2": "#4E9F3D", "P3": "#B7791F", "P4": "#7C3AED", "P5": "#C2415D"}
    chart_left, chart_right = 250, 1510
    scale = (chart_right - chart_left) / 28
    for row_index, (algorithm, blocks) in enumerate(schedules.items()):
        y = 280 + row_index * 180
        draw.text((75, y + 25), algorithm, font=font(27, bold=True), fill=INK)
        draw.line((chart_left, y + 70, chart_right, y + 70), fill=GRID, width=2)
        for start, end, process in blocks:
            x1 = chart_left + start * scale
            x2 = chart_left + end * scale
            draw.rectangle((x1, y, x2, y + 70), fill=colors[process], outline=WHITE, width=2)
            if x2 - x1 >= 38:
                centered_text(draw, (x1, y, x2, y + 70), process, font(20, bold=True), fill=WHITE)
            draw.text((x1 - 5, y + 80), str(start), font=font(16, mono=True), fill=MUTED)
        draw.text((chart_right - 8, y + 80), "28", font=font(16, mono=True), fill=MUTED)
    draw.text((75, 792), "说明：SJF 为非抢占式；RR 时间片为 2，矩形宽度表示 CPU 占用时长。", font=font(23), fill=MUTED)
    save(image, path)


def scheduling_metrics(path: Path):
    image, draw = new_canvas("调度性能指标对比", "所有指标均由同一测试集计算，单位为时间片", accent=TEAL)
    algorithms = ["FCFS", "SJF", "RR(q=2)"]
    metrics = {
        "平均等待时间": [11.4, 8.2, 13.0],
        "平均周转时间": [17.0, 13.8, 18.6],
        "平均首次响应": [11.4, 8.2, 2.8],
    }
    colors = [BLUE, TEAL, ORANGE]
    chart = (160, 250, 1110, 760)
    x1, y1, x2, y2 = chart
    max_value = 20
    for tick in range(0, max_value + 1, 4):
        y = y2 - (tick / max_value) * (y2 - y1)
        draw.line((x1, y, x2, y), fill=GRID, width=1)
        draw.text((105, y - 12), str(tick), font=font(18, mono=True), fill=MUTED)
    group_width = (x2 - x1) / len(algorithms)
    bar_width = 70
    for algorithm_index, algorithm in enumerate(algorithms):
        base_x = x1 + algorithm_index * group_width + 50
        for metric_index, (_, values) in enumerate(metrics.items()):
            value = values[algorithm_index]
            left = base_x + metric_index * 82
            top = y2 - (value / max_value) * (y2 - y1)
            draw.rectangle((left, top, left + bar_width, y2), fill=colors[metric_index])
            centered_text(draw, (left - 8, top - 36, left + bar_width + 8, top), f"{value:.1f}", font(18, bold=True), fill=INK)
        centered_text(draw, (base_x, y2 + 15, base_x + 230, y2 + 60), algorithm, font(22, bold=True), fill=INK)
    legend_y = 270
    for index, metric_name in enumerate(metrics):
        draw.rectangle((1190, legend_y, 1222, legend_y + 32), fill=colors[index])
        draw.text((1240, legend_y - 2), metric_name, font=font(21), fill=INK)
        legend_y += 62
    rounded(draw, (1170, 500, 1510, 735), fill=SOFT, outline=GRID, width=2)
    multiline(
        draw,
        (1200, 530),
        "结论\n• SJF 的平均等待和周转时间最低。\n• RR 的首次响应最快，但切换次数最多。\n• FCFS 实现简单，短作业可能受长作业阻塞。",
        font(21),
        fill=MUTED,
        max_width=280,
        line_gap=10,
    )
    save(image, path)


def scheduling_test_matrix(path: Path):
    image, draw = new_canvas("边界条件与回归测试", "除平均性能外，还验证空闲时段、同时到达和极小时间片", accent=ORANGE)
    headers = ["编号", "测试条件", "预期行为", "执行结果"]
    rows = [
        ("T01", "首进程到达时间为 3", "时钟直接推进到 3，不产生负等待", "通过"),
        ("T02", "3 个进程同时到达", "FCFS 保持输入顺序，SJF 按服务时间排序", "通过"),
        ("T03", "服务时间均为 1", "三种算法完成时间一致", "通过"),
        ("T04", "RR 时间片 q=1", "每次仅执行 1 个时间片，剩余时间正确递减", "通过"),
        ("T05", "进程集合为空", "输出空结果并提示无可调度进程", "通过"),
        ("T06", "非法服务时间 0", "输入校验拒绝执行", "通过"),
    ]
    draw_table(draw, (90, 235, 1510, 760), headers, rows, widths=[0.10, 0.28, 0.47, 0.15], header_fill=LIGHT_ORANGE)
    draw.text((90, 802), "回归结论：6/6 用例通过；算法主流程和异常分支均有可复现证据。", font=font(23, bold=True), fill=GREEN)
    save(image, path)


def scheduling_console(path: Path):
    lines = [
        (r"> python scheduler.py --dataset processes.csv --quantum 2", "#7DD3FC"),
        ("Loaded 5 processes: P1 P2 P3 P4 P5", TERMINAL_TEXT),
        ("", TERMINAL_TEXT),
        ("Algorithm   AvgWait   AvgTurnaround   AvgResponse   Switches", "#86EFAC"),
        ("FCFS          11.40          17.00          11.40          4", TERMINAL_TEXT),
        ("SJF            8.20          13.80           8.20          4", TERMINAL_TEXT),
        ("RR(q=2)       13.00          18.60           2.80         14", TERMINAL_TEXT),
        ("", TERMINAL_TEXT),
        ("Assertions: 31 passed, 0 failed", "#86EFAC"),
        ("Output written to results/scheduling-summary.json", "#A5B4FC"),
    ]
    terminal_image("scheduler.py — Windows Terminal", lines, path)


def architecture(path: Path):
    image, draw = new_canvas("学生成绩管理系统总体架构", "采用界面层、业务层、数据访问层和 SQLite 数据库的分层结构")
    layers = [
        ("界面交互层", "学生维护 · 课程维护 · 成绩录入 · 查询统计", LIGHT_BLUE, BLUE),
        ("业务服务层", "输入校验 · 成绩规则 · 统计计算 · 事务协调", LIGHT_TEAL, TEAL),
        ("数据访问层", "StudentDAO · CourseDAO · ScoreDAO · ReportDAO", LIGHT_ORANGE, ORANGE),
        ("数据持久层", "SQLite：student / course / score / user / operation_log", LIGHT_GREEN, GREEN),
    ]
    y = 235
    for index, (title, detail, fill, accent) in enumerate(layers):
        box = (220, y, 1380, y + 115)
        rounded(draw, box, fill=fill, outline=accent, width=3)
        draw.text((265, y + 25), title, font=font(28, bold=True), fill=accent)
        draw.text((530, y + 29), detail, font=font(24), fill=INK)
        if index < len(layers) - 1:
            arrow(draw, (800, y + 115), (800, y + 145), fill=MUTED, width=4)
        y += 150
    draw.text((235, 825), "设计原则：界面不直接执行 SQL；业务规则集中在 Service；所有写操作记录审计日志。", font=font(22), fill=MUTED)
    save(image, path)


def er_model(path: Path):
    image, draw = new_canvas("核心数据模型与关系", "学生、课程与成绩记录构成多对多关系，成绩表保存关联与业务数据", accent=TEAL)
    entities = {
        "Student": ((90, 270, 450, 610), ["student_id PK", "name", "class_name", "major"]),
        "Score": ((620, 235, 980, 645), ["score_id PK", "student_id FK", "course_id FK", "usual_score", "final_score", "total_score"]),
        "Course": ((1150, 270, 1510, 610), ["course_id PK", "course_name", "credit", "teacher"]),
    }
    for name, (box, fields) in entities.items():
        rounded(draw, box, fill=SOFT, outline=TEAL, width=3)
        draw.rectangle((box[0], box[1], box[2], box[1] + 70), fill=LIGHT_TEAL)
        centered_text(draw, (box[0], box[1], box[2], box[1] + 70), name, font(28, bold=True), fill=TEAL)
        y = box[1] + 95
        for field_name in fields:
            draw.text((box[0] + 35, y), field_name, font=font(22, mono=True), fill=INK)
            y += 48
    draw.line((450, 430, 620, 430), fill=TEAL, width=5)
    draw.line((980, 430, 1150, 430), fill=TEAL, width=5)
    draw.text((485, 390), "1", font=font(26, bold=True), fill=TEAL)
    draw.text((575, 390), "N", font=font(26, bold=True), fill=TEAL)
    draw.text((1020, 390), "N", font=font(26, bold=True), fill=TEAL)
    draw.text((1110, 390), "1", font=font(26, bold=True), fill=TEAL)
    rounded(draw, (365, 710, 1235, 810), fill=WHITE, outline=GRID, width=2)
    centered_text(draw, (385, 725, 1215, 795), "唯一约束：(student_id, course_id)；删除学生或课程前必须检查关联成绩。", font(23), fill=MUTED)
    save(image, path)


def desktop_shell(draw: ImageDraw.ImageDraw, title: str):
    rounded(draw, (65, 65, WIDTH - 65, HEIGHT - 65), fill=WHITE, outline="#98A2B3", width=3, radius=20)
    draw.rectangle((65, 65, WIDTH - 65, 135), fill="#243447")
    draw.text((105, 86), title, font=font(25, bold=True), fill=WHITE)
    for index, color in enumerate(("#28C840", "#FEBB2E", "#FF5F57")):
        draw.ellipse((1440 + index * 34, 90, 1458 + index * 34, 108), fill=color)
    draw.rectangle((65, 135, 305, HEIGHT - 65), fill="#F2F4F7")
    menus = ["工作台", "学生信息", "课程信息", "成绩录入", "查询统计", "系统设置"]
    y = 185
    for menu in menus:
        fill = BLUE if menu == "学生信息" else MUTED
        if menu == "学生信息":
            rounded(draw, (85, y - 10, 285, y + 42), fill=LIGHT_BLUE, radius=10)
        draw.text((115, y), menu, font=font(21, bold=menu == "学生信息"), fill=fill)
        y += 72


def student_ui(path: Path):
    image = Image.new("RGB", (WIDTH, HEIGHT), "#E9EEF4")
    draw = ImageDraw.Draw(image)
    desktop_shell(draw, "学生成绩管理系统  v1.0")
    draw.text((350, 175), "学生信息管理", font=font(32, bold=True), fill=INK)
    rounded(draw, (350, 230, 1490, 320), fill=SOFT, outline=GRID, width=2)
    draw.text((385, 260), "学号 / 姓名 / 班级", font=font(21), fill=MUTED)
    rounded(draw, (1220, 248, 1335, 302), fill=BLUE, radius=10)
    centered_text(draw, (1220, 248, 1335, 302), "查询", font(21, bold=True), fill=WHITE)
    rounded(draw, (1350, 248, 1465, 302), fill=TEAL, radius=10)
    centered_text(draw, (1350, 248, 1465, 302), "新增", font(21, bold=True), fill=WHITE)
    headers = ["学号", "姓名", "班级", "专业", "状态"]
    rows = [
        ("20260011", "林晨", "计科 2201", "计算机科学与技术", "正常"),
        ("20260012", "周雨", "计科 2201", "计算机科学与技术", "正常"),
        ("20260013", "陈帆", "软件 2202", "软件工程", "正常"),
        ("20260014", "许宁", "软件 2202", "软件工程", "正常"),
        ("20260015", "方远", "计科 2203", "计算机科学与技术", "正常"),
    ]
    draw_table(draw, (350, 355, 1490, 735), headers, rows, widths=[0.18, 0.14, 0.18, 0.34, 0.16])
    draw.text((350, 775), "共 5 条记录  |  最近保存：2026-04-30 16:42", font=font(20), fill=MUTED)
    save(image, path)


def score_entry_ui(path: Path):
    image = Image.new("RGB", (WIDTH, HEIGHT), "#E9EEF4")
    draw = ImageDraw.Draw(image)
    desktop_shell(draw, "学生成绩管理系统  v1.0")
    draw.text((350, 175), "成绩录入", font=font(32, bold=True), fill=INK)
    fields = [
        ("学生", "20260011  林晨"),
        ("课程", "CS204  数据结构"),
        ("平时成绩", "86"),
        ("期末成绩", "92"),
        ("总评成绩", "90.2"),
    ]
    y = 255
    for label, value in fields:
        draw.text((395, y + 12), label, font=font(23, bold=True), fill=INK)
        rounded(draw, (590, y, 1240, y + 58), fill=WHITE if label != "总评成绩" else LIGHT_TEAL, outline=GRID, width=2, radius=9)
        draw.text((625, y + 12), value, font=font(23), fill=TEAL if label == "总评成绩" else INK)
        y += 88
    rounded(draw, (590, 710, 760, 770), fill=BLUE, radius=10)
    centered_text(draw, (590, 710, 760, 770), "保存成绩", font(22, bold=True), fill=WHITE)
    rounded(draw, (785, 710, 935, 770), fill=SOFT, outline=GRID, radius=10)
    centered_text(draw, (785, 710, 935, 770), "重置", font(22, bold=True), fill=INK)
    rounded(draw, (1000, 685, 1450, 790), fill=LIGHT_GREEN, outline=GREEN, width=2, radius=12)
    draw.text((1035, 712), "✓ 保存成功", font=font(24, bold=True), fill=GREEN)
    draw.text((1035, 753), "成绩记录已写入数据库并生成审计日志。", font=font(19), fill=INK)
    save(image, path)


def statistics_dashboard(path: Path):
    image = Image.new("RGB", (WIDTH, HEIGHT), "#F2F4F7")
    draw = ImageDraw.Draw(image)
    rounded(draw, (55, 55, WIDTH - 55, HEIGHT - 55), fill=WHITE, outline=GRID, width=2, radius=20)
    draw.text((95, 90), "课程成绩统计 · 数据结构", font=font(34, bold=True), fill=INK)
    cards = [
        ("学生人数", "36", BLUE, LIGHT_BLUE),
        ("平均分", "82.6", TEAL, LIGHT_TEAL),
        ("最高分", "96.0", GREEN, LIGHT_GREEN),
        ("及格率", "91.7%", ORANGE, LIGHT_ORANGE),
    ]
    x = 95
    for title, value, accent, fill in cards:
        rounded(draw, (x, 165, x + 325, 290), fill=fill, outline=accent, width=2)
        draw.text((x + 28, 190), title, font=font(21), fill=MUTED)
        draw.text((x + 28, 228), value, font=font(38, bold=True), fill=accent)
        x += 365
    chart_box = (95, 350, 1015, 775)
    draw.rectangle(chart_box, fill=SOFT, outline=GRID, width=2)
    categories = ["<60", "60–69", "70–79", "80–89", "90–100"]
    values = [3, 4, 7, 13, 9]
    chart_left, chart_bottom = 170, 720
    max_value = 15
    for tick in range(0, 16, 3):
        y = chart_bottom - tick / max_value * 300
        draw.line((chart_left, y, 950, y), fill=GRID, width=1)
        draw.text((130, y - 10), str(tick), font=font(17, mono=True), fill=MUTED)
    for index, (category, value) in enumerate(zip(categories, values)):
        left = chart_left + index * 150 + 35
        top = chart_bottom - value / max_value * 300
        draw.rectangle((left, top, left + 80, chart_bottom), fill=TEAL)
        centered_text(draw, (left, top - 35, left + 80, top), str(value), font(19, bold=True), fill=INK)
        centered_text(draw, (left - 25, chart_bottom + 10, left + 105, chart_bottom + 50), category, font(18), fill=INK)
    rounded(draw, (1060, 350, 1505, 775), fill=SOFT, outline=GRID, width=2)
    draw.text((1100, 390), "质量检查", font=font(27, bold=True), fill=INK)
    checks = [
        ("成绩范围 0–100", "通过"),
        ("重复成绩记录", "0 条"),
        ("缺失学生关联", "0 条"),
        ("统计结果复核", "通过"),
        ("导出记录数", "36 条"),
    ]
    y = 455
    for label, result in checks:
        draw.text((1100, y), label, font=font(21), fill=MUTED)
        draw.text((1390, y), result, font=font(21, bold=True), fill=GREEN)
        y += 62
    save(image, path)


def course_test_matrix(path: Path):
    image, draw = new_canvas("系统功能测试记录", "覆盖正常流程、边界值、重复数据与关联完整性", accent=ORANGE)
    headers = ["用例", "操作与输入", "预期结果", "实际结果"]
    rows = [
        ("TC01", "新增学生 20260011", "保存成功，列表出现新记录", "通过"),
        ("TC02", "重复录入学号 20260011", "拒绝保存并提示学号重复", "通过"),
        ("TC03", "录入总评成绩 101", "拒绝保存并提示范围 0–100", "通过"),
        ("TC04", "按课程统计 36 条成绩", "平均分 82.6，最高分 96.0", "通过"),
        ("TC05", "删除存在成绩的学生", "阻止删除并提示先处理关联记录", "通过"),
        ("TC06", "导出课程成绩", "生成 36 条记录的 CSV 文件", "通过"),
    ]
    draw_table(draw, (80, 230, 1520, 765), headers, rows, widths=[0.10, 0.33, 0.42, 0.15], header_fill=LIGHT_ORANGE)
    draw.text((80, 808), "测试结论：关键功能 6/6 通过，数据校验和关联约束符合设计要求。", font=font(23, bold=True), fill=GREEN)
    save(image, path)


def build_all(repo_root: Path) -> list[Path]:
    cases = repo_root / "examples" / "cases"
    outputs: list[Path] = []

    network = cases / "network-dos" / "assets"
    network_topology(network / "network-topology.png")
    terminal_image(
        "Host A — ipconfig /all",
        [
            (r"C:\> ipconfig /all", "#7DD3FC"),
            ("Windows IP Configuration", TERMINAL_TEXT),
            ("", TERMINAL_TEXT),
            ("Ethernet adapter Ethernet0:", "#86EFAC"),
            ("   Physical Address . . . . . : 00-0C-29-11-0A-01", TERMINAL_TEXT),
            ("   DHCP Enabled. . . . . . . : No", TERMINAL_TEXT),
            ("   IPv4 Address. . . . . . . : 192.168.10.11 (Preferred)", TERMINAL_TEXT),
            ("   Subnet Mask . . . . . . . : 255.255.255.0", TERMINAL_TEXT),
            ("   Default Gateway . . . . . :", TERMINAL_TEXT),
        ],
        network / "host-a-ipconfig.png",
    )
    terminal_image(
        "Host B — netsh interface ipv4 show config",
        [
            (r'C:\> netsh interface ipv4 show config name="Ethernet0"', "#7DD3FC"),
            ('Configuration for interface "Ethernet0"', "#86EFAC"),
            ("   DHCP enabled:                         No", TERMINAL_TEXT),
            ("   IP Address:                           192.168.10.12", TERMINAL_TEXT),
            ("   Subnet Prefix:                        192.168.10.0/24", TERMINAL_TEXT),
            ("   InterfaceMetric:                      25", TERMINAL_TEXT),
            ("", TERMINAL_TEXT),
            (r"C:\> ping 192.168.10.11 -n 1", "#7DD3FC"),
            ("Reply from 192.168.10.11: bytes=32 time<1ms TTL=128", "#86EFAC"),
        ],
        network / "host-b-config.png",
    )
    terminal_image(
        "Host A — ping 与 ARP 验证",
        [
            (r"C:\> ping 192.168.10.12 -n 4", "#7DD3FC"),
            ("Reply from 192.168.10.12: bytes=32 time<1ms TTL=128", TERMINAL_TEXT),
            ("Reply from 192.168.10.12: bytes=32 time<1ms TTL=128", TERMINAL_TEXT),
            ("Reply from 192.168.10.12: bytes=32 time<1ms TTL=128", TERMINAL_TEXT),
            ("Reply from 192.168.10.12: bytes=32 time<1ms TTL=128", TERMINAL_TEXT),
            ("Packets: Sent = 4, Received = 4, Lost = 0 (0% loss)", "#86EFAC"),
            ("", TERMINAL_TEXT),
            (r"C:\> arp -a", "#7DD3FC"),
            ("192.168.10.12   00-0c-29-12-0b-02   dynamic", "#86EFAC"),
        ],
        network / "connectivity-evidence.png",
    )
    network_diagnosis(network / "failure-diagnosis.png")
    outputs.extend(sorted(network.glob("*.png")))

    scheduling = cases / "os-process-scheduling" / "assets"
    scheduling_flow(scheduling / "scheduling-flow.png")
    scheduling_gantt(scheduling / "scheduling-gantt.png")
    scheduling_metrics(scheduling / "scheduling-metrics.png")
    scheduling_console(scheduling / "scheduling-console.png")
    scheduling_test_matrix(scheduling / "scheduling-tests.png")
    outputs.extend(sorted(scheduling.glob("*.png")))

    course = cases / "course-design-student-management" / "assets"
    architecture(course / "system-architecture.png")
    er_model(course / "data-model.png")
    student_ui(course / "student-management-ui.png")
    score_entry_ui(course / "score-entry-ui.png")
    statistics_dashboard(course / "statistics-dashboard.png")
    course_test_matrix(course / "system-tests.png")
    outputs.extend(sorted(course.glob("*.png")))
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description="Build deterministic, school-neutral showcase assets.")
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
        help="Repository root. Defaults to the parent of scripts/.",
    )
    args = parser.parse_args()
    for output in build_all(args.repo_root.resolve()):
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
