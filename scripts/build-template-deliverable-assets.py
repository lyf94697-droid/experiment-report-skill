from __future__ import annotations

import argparse
import importlib.util
from pathlib import Path

from PIL import Image, ImageDraw


def load_showcase_helpers(repo_root: Path):
    module_path = repo_root / "scripts" / "build-deliverable-showcase-assets.py"
    spec = importlib.util.spec_from_file_location("showcase_assets", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load helper module: {module_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def database_plan(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "复合索引建立前后的执行计划",
        "目标查询：按 user_id 与 status 过滤，按 created_at 倒序返回最近 20 条记录",
        accent=s.BLUE,
    )
    headers = ["测试阶段", "访问方式", "实际读取行数", "额外排序", "耗时中位数"]
    rows = [
        ("仅主键索引", "ALL", "200000", "是", "78.4 ms"),
        ("复合索引", "range", "23", "否", "2.1 ms"),
    ]
    s.draw_table(
        draw,
        (90, 245, 1510, 515),
        headers,
        rows,
        widths=[0.22, 0.16, 0.22, 0.16, 0.24],
        header_fill=s.LIGHT_BLUE,
        font_size=24,
    )
    s.rounded(draw, (130, 605, 1470, 790), fill=s.SOFT, outline=s.GRID, width=2)
    draw.text((175, 638), "索引定义", font=s.font(27, bold=True), fill=s.BLUE)
    draw.text(
        (360, 640),
        "CREATE INDEX idx_orders_user_status_time",
        font=s.font(25, mono=True),
        fill=s.INK,
    )
    draw.text(
        (360, 690),
        "ON orders(user_id, status, created_at DESC);",
        font=s.font(25, mono=True),
        fill=s.INK,
    )
    draw.text(
        (175, 748),
        "结论：读取行数减少 99.99%，查询耗时降低 97.3%，且不再产生 filesort。",
        font=s.font(23, bold=True),
        fill=s.GREEN,
    )
    s.save(image, path)


def database_performance(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "查询收益与写入代价",
        "同一数据集、同一测试机，预热后重复执行并取中位数",
        accent=s.TEAL,
    )
    panels = [
        ("目标查询耗时", 78.4, 2.1, "ms", s.BLUE),
        ("批量插入耗时", 312.0, 458.0, "ms", s.ORANGE),
    ]
    y = 235
    for title, before, after, unit, accent in panels:
        draw.text((110, y), title, font=s.font(28, bold=True), fill=s.INK)
        max_value = max(before, after)
        for index, (label, value, color) in enumerate(
            (("建索引前", before, s.MUTED), ("建索引后", after, accent))
        ):
            bar_y = y + 70 + index * 76
            width = 1020 * value / max_value
            draw.rectangle((270, bar_y, 270 + width, bar_y + 45), fill=color)
            draw.text((110, bar_y + 5), label, font=s.font(21), fill=s.INK)
            draw.text(
                (300 + width, bar_y + 5),
                f"{value:.1f} {unit}",
                font=s.font(22, bold=True),
                fill=s.INK,
            )
        y += 280
    s.rounded(draw, (1160, 250, 1500, 710), fill=s.SOFT, outline=s.GRID, width=2)
    s.multiline(
        draw,
        (1200, 290),
        "工程判断\n\n查询收益明显\n\n写入耗时增加 46.8%\n\n空间增加 11.6 MB\n\n适合读多写少的订单查询",
        s.font(23),
        fill=s.INK,
        max_width=260,
        line_gap=12,
    )
    s.save(image, path)


def responsive_layout(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "响应式课程卡片在三种视口下的布局",
        "同一套 HTML 结构通过 Grid、minmax 和媒体查询自动调整列数",
        accent=s.TEAL,
    )
    viewports = [
        ((70, 240, 730, 745), "1440 px · 四列", 4),
        ((795, 240, 1155, 745), "768 px · 两列", 2),
        ((1220, 240, 1530, 745), "390 px · 单列", 1),
    ]
    for box, label, columns in viewports:
        x1, y1, x2, y2 = box
        s.rounded(draw, box, fill="#F5F7FA", outline=s.GRID, width=3, radius=16)
        draw.rectangle((x1, y1, x2, y1 + 46), fill="#263746")
        draw.text((x1 + 20, y1 + 10), label, font=s.font(19, bold=True), fill=s.WHITE)
        gap = 12
        content_left = x1 + 22
        content_right = x2 - 22
        card_width = (content_right - content_left - gap * (columns - 1)) / columns
        rows = 2 if columns > 1 else 3
        for row in range(rows):
            for column in range(columns):
                top = y1 + 78 + row * 184
                left = content_left + column * (card_width + gap)
                right = left + card_width
                bottom = top + 158
                s.rounded(
                    draw,
                    (left, top, right, bottom),
                    fill=s.WHITE,
                    outline="#C9D3DC",
                    width=2,
                    radius=10,
                )
                draw.rectangle((left, top, right, top + 52), fill=s.LIGHT_TEAL)
                draw.text(
                    (left + 12, top + 68),
                    f"课程 {row * columns + column + 1}",
                    font=s.font(16, bold=True),
                    fill=s.INK,
                )
                draw.rectangle((left + 12, bottom - 34, right - 12, bottom - 16), fill=s.TEAL)
    draw.text(
        (155, 795),
        "检查结果：无横向滚动条；标题最多两行；按钮高度 44 px；键盘焦点顺序与视觉顺序一致。",
        font=s.font(24, bold=True),
        fill=s.GREEN,
    )
    s.save(image, path)


def web_quality(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "多视口检查与质量记录",
        "在浏览器 100% 缩放下检查换行、溢出、焦点顺序和首屏稳定性",
        accent=s.BLUE,
    )
    cards = [
        ("无障碍", "100", s.GREEN, s.LIGHT_GREEN),
        ("性能", "98", s.BLUE, s.LIGHT_BLUE),
        ("布局偏移", "0.002", s.TEAL, s.LIGHT_TEAL),
    ]
    x = 110
    for title, value, accent, fill in cards:
        s.rounded(draw, (x, 225, x + 390, 385), fill=fill, outline=accent, width=3)
        draw.text((x + 28, 250), title, font=s.font(24), fill=s.MUTED)
        draw.text((x + 28, 292), value, font=s.font(48, bold=True), fill=accent)
        x += 495
    headers = ["视口", "列数", "横向滚动", "标题溢出", "按钮高度", "结果"]
    rows = [
        ("1440 px", "4", "无", "无", "44 px", "通过"),
        ("1024 px", "3", "无", "无", "44 px", "通过"),
        ("768 px", "2", "无", "无", "44 px", "通过"),
        ("390 px", "1", "无", "无", "44 px", "通过"),
        ("320 px", "1", "无", "无", "44 px", "通过"),
    ]
    s.draw_table(
        draw,
        (100, 455, 1500, 790),
        headers,
        rows,
        widths=[0.16, 0.10, 0.18, 0.18, 0.18, 0.20],
        header_fill=s.LIGHT_BLUE,
        font_size=21,
    )
    s.save(image, path)


def junit_matrix(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "成绩等级函数的边界值测试",
        "有效区间与无效区间均覆盖，失败用例可独立定位",
        accent=s.ORANGE,
    )
    headers = ["输入", "预期", "首次结果", "修正后", "说明"]
    rows = [
        ("-1", "异常", "通过", "通过", "低于有效范围"),
        ("0 / 59", "F", "通过", "通过", "F 区间边界"),
        ("60 / 69", "D", "通过", "通过", "D 区间边界"),
        ("70 / 79", "C", "通过", "通过", "C 区间边界"),
        ("80 / 89", "B", "通过", "通过", "B 区间边界"),
        ("90 / 100", "A", "90 失败", "通过", "修正 > 为 >="),
        ("101", "异常", "通过", "通过", "高于有效范围"),
    ]
    s.draw_table(
        draw,
        (85, 225, 1515, 790),
        headers,
        rows,
        widths=[0.18, 0.14, 0.18, 0.18, 0.32],
        header_fill=s.LIGHT_ORANGE,
        font_size=22,
    )
    s.save(image, path)


def merge_flow(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "归并排序的分解与合并过程",
        "输入序列递归二分，底层有序子序列逐层稳定合并",
        accent=s.BLUE,
    )
    levels = [
        (235, ["[7, 3, 5, 3, -2, 9]"]),
        (390, ["[7, 3, 5]", "[3, -2, 9]"]),
        (545, ["[7]", "[3, 5]", "[3]", "[-2, 9]"]),
        (700, ["[-2, 3, 3, 5, 7, 9]"]),
    ]
    for y, labels in levels:
        total_width = 1420
        gap = 35
        box_width = (total_width - gap * (len(labels) - 1)) / len(labels)
        for index, label in enumerate(labels):
            left = 90 + index * (box_width + gap)
            box = (left, y, left + box_width, y + 82)
            fill = s.LIGHT_GREEN if y == 700 else s.SOFT
            outline = s.GREEN if y == 700 else s.BLUE
            s.rounded(draw, box, fill=fill, outline=outline, width=3, radius=12)
            s.centered_text(
                draw, box, label, s.font(23, bold=y in {235, 700}, mono=True), fill=s.INK
            )
    for start, end in [
        ((800, 317), (430, 390)),
        ((800, 317), (1170, 390)),
        ((430, 472), (250, 545)),
        ((430, 472), (610, 545)),
        ((1170, 472), (990, 545)),
        ((1170, 472), (1350, 545)),
        ((800, 627), (800, 700)),
    ]:
        s.arrow(draw, start, end, fill=s.MUTED, width=4)
    draw.text(
        (495, 810),
        "相等元素优先取左侧：left[i] <= right[j]，从而保持稳定性。",
        font=s.font(24, bold=True),
        fill=s.GREEN,
    )
    s.save(image, path)


def benchmark_chart(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "归并排序性能测试",
        "固定随机种子，每个规模预热一次、执行七次并取中位数",
        accent=s.TEAL,
    )
    labels = ["1k", "5k", "10k", "50k", "100k"]
    values = [1.34, 7.92, 17.10, 102.63, 221.48]
    left, top, right, bottom = 150, 250, 1110, 760
    max_value = 240
    for tick in range(0, 241, 40):
        y = bottom - tick / max_value * (bottom - top)
        draw.line((left, y, right, y), fill=s.GRID, width=1)
        draw.text((85, y - 12), str(tick), font=s.font(18, mono=True), fill=s.MUTED)
    bar_width = 120
    for index, (label, value) in enumerate(zip(labels, values)):
        x = left + 65 + index * 180
        bar_top = bottom - value / max_value * (bottom - top)
        draw.rectangle((x, bar_top, x + bar_width, bottom), fill=s.TEAL)
        s.centered_text(
            draw,
            (x - 25, bar_top - 42, x + bar_width + 25, bar_top),
            f"{value:.2f}",
            s.font(19, bold=True),
            fill=s.INK,
        )
        s.centered_text(
            draw,
            (x - 10, bottom + 10, x + bar_width + 10, bottom + 52),
            label,
            s.font(20, bold=True),
            fill=s.INK,
        )
    s.rounded(draw, (1180, 285, 1510, 700), fill=s.SOFT, outline=s.GRID, width=2)
    s.multiline(
        draw,
        (1220, 325),
        "测试结论\n\n18 项测试全部通过\n\n100k 元素：221.48 ms\n\n规模扩大 10 倍\n耗时约扩大 12.95 倍\n\n趋势符合 O(n log n)",
        s.font(22),
        fill=s.INK,
        max_width=250,
        line_gap=11,
    )
    s.save(image, path)


def calibration_chart(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "温度传感器标定曲线",
        "8 个温度点线性拟合：T = 12.497U - 6.247，R² = 0.9998",
        accent=s.GREEN,
    )
    voltages = [0.50, 1.30, 2.10, 2.50, 2.90, 3.70, 4.50, 5.30]
    temperatures = [0, 10, 20, 25, 30, 40, 50, 60]
    left, top, right, bottom = 170, 230, 1250, 760
    for tick in range(0, 61, 10):
        y = bottom - tick / 60 * (bottom - top)
        draw.line((left, y, right, y), fill=s.GRID, width=1)
        draw.text((105, y - 12), str(tick), font=s.font(18, mono=True), fill=s.MUTED)
    for tick in range(0, 6):
        x = left + tick / 5.5 * (right - left)
        draw.line((x, top, x, bottom), fill=s.GRID, width=1)
        draw.text((x - 8, bottom + 15), str(tick), font=s.font(18, mono=True), fill=s.MUTED)
    points = []
    for voltage, temperature in zip(voltages, temperatures):
        x = left + voltage / 5.5 * (right - left)
        y = bottom - temperature / 60 * (bottom - top)
        points.append((x, y))
        draw.ellipse((x - 8, y - 8, x + 8, y + 8), fill=s.GREEN, outline=s.WHITE)
    draw.line(points, fill=s.TEAL, width=4)
    draw.text((585, 810), "输出电压 U / V", font=s.font(22, bold=True), fill=s.INK)
    s.rounded(draw, (1310, 260, 1520, 700), fill=s.LIGHT_GREEN, outline=s.GREEN, width=2)
    s.multiline(
        draw,
        (1340, 300),
        "拟合结果\n\nR²  0.9998\n\n最大绝对误差\n0.42 ℃\n\n满量程相对误差\n0.70%\n\n判定：合格",
        s.font(21),
        fill=s.INK,
        max_width=150,
        line_gap=12,
    )
    s.save(image, path)


def android_architecture(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "任务清单应用总体结构",
        "界面状态、业务规则、数据库与后台提醒通过明确接口协作",
        accent=s.BLUE,
    )
    layers = [
        ("界面层", "Compose 页面 · Navigation · ViewModel · UiState", s.LIGHT_BLUE, s.BLUE),
        ("领域层", "用例 · TaskFilter · 输入校验 · 日期边界", s.LIGHT_TEAL, s.TEAL),
        ("数据层", "Repository · Room DAO · 备份序列化", s.LIGHT_ORANGE, s.ORANGE),
        ("系统服务", "WorkManager · 通知 · Storage Access Framework", s.LIGHT_GREEN, s.GREEN),
    ]
    y = 220
    for index, (title, detail, fill, accent) in enumerate(layers):
        box = (210, y, 1390, y + 115)
        s.rounded(draw, box, fill=fill, outline=accent, width=3)
        draw.text((260, y + 26), title, font=s.font(28, bold=True), fill=accent)
        draw.text((510, y + 30), detail, font=s.font(24), fill=s.INK)
        if index < len(layers) - 1:
            s.arrow(draw, (800, y + 115), (800, y + 150), fill=s.MUTED, width=4)
        y += 155
    draw.text(
        (255, 835),
        "单一数据源：Room 保存持久状态，ViewModel 组合界面状态，Worker 执行前再次核对数据库。",
        font=s.font(23, bold=True),
        fill=s.BLUE,
    )
    s.save(image, path)


def android_ui(s, path: Path) -> None:
    image = Image.new("RGB", (s.WIDTH, s.HEIGHT), "#EEF1F5")
    draw = ImageDraw.Draw(image)
    s.rounded(draw, (500, 35, 1100, 865), fill=s.WHITE, outline="#7B8794", width=4, radius=32)
    draw.rectangle((500, 35, 1100, 125), fill="#294C60")
    draw.text((545, 70), "我的任务", font=s.font(31, bold=True), fill=s.WHITE)
    stats = [("未完成", "286"), ("今天", "18"), ("已过期", "23")]
    x = 535
    for label, value in stats:
        s.rounded(draw, (x, 155, x + 160, 250), fill=s.LIGHT_BLUE, outline=s.BLUE, width=2, radius=12)
        draw.text((x + 18, 172), label, font=s.font(18), fill=s.MUTED)
        draw.text((x + 18, 205), value, font=s.font(28, bold=True), fill=s.BLUE)
        x += 180
    s.rounded(draw, (535, 285, 1065, 345), fill=s.SOFT, outline=s.GRID, width=2, radius=12)
    draw.text((565, 302), "搜索标题或备注", font=s.font(20), fill=s.MUTED)
    tasks = [
        ("数据库实验报告", "今天 22:00", "高"),
        ("移动应用测试", "明天 18:30", "中"),
        ("整理课程笔记", "6 月 25 日", "低"),
    ]
    y = 380
    colors = [s.RED, s.ORANGE, s.GREEN]
    for (title, due, priority), color in zip(tasks, colors):
        s.rounded(draw, (535, y, 1065, y + 125), fill=s.WHITE, outline=s.GRID, width=2, radius=12)
        draw.ellipse((560, y + 42, 590, y + 72), outline=s.TEAL, width=3)
        draw.text((620, y + 24), title, font=s.font(23, bold=True), fill=s.INK)
        draw.text((620, y + 70), due, font=s.font(19), fill=s.MUTED)
        s.rounded(draw, (960, y + 38, 1030, y + 78), fill=color, radius=9)
        s.centered_text(draw, (960, y + 38, 1030, y + 78), priority, s.font(18, bold=True), fill=s.WHITE)
        y += 145
    draw.ellipse((955, 765, 1045, 855), fill=s.TEAL)
    s.centered_text(draw, (955, 765, 1045, 855), "+", s.font(48, bold=True), fill=s.WHITE)
    draw.text((90, 260), "运行结果", font=s.font(38, bold=True), fill=s.INK)
    s.multiline(
        draw,
        (90, 335),
        "• 500 条任务稳定加载\n\n• 组合筛选中位耗时 18 ms\n\n• 完成任务后统计同步更新\n\n• 无网络状态下可正常增删改查\n\n• 提醒与备份恢复测试通过",
        s.font(25),
        fill=s.INK,
        max_width=340,
        line_gap=15,
    )
    s.save(image, path)


def android_tests(s, path: Path) -> None:
    image, draw = s.new_canvas(
        "核心功能测试记录",
        "覆盖数据校验、状态切换、日期边界、后台提醒和备份恢复",
        accent=s.ORANGE,
    )
    headers = ["测试组", "用例数量", "关键检查", "结果"]
    rows = [
        ("输入校验", "2", "空标题、非法时间关系", "通过"),
        ("任务状态", "2", "新增、完成状态切换", "通过"),
        ("查询筛选", "2", "组合筛选、过期边界", "通过"),
        ("后台提醒", "3", "替换、完成取消、删除取消", "通过"),
        ("备份恢复", "3", "导出、损坏拒绝、完整恢复", "通过"),
    ]
    s.draw_table(
        draw,
        (100, 235, 1500, 680),
        headers,
        rows,
        widths=[0.20, 0.16, 0.46, 0.18],
        header_fill=s.LIGHT_ORANGE,
        font_size=23,
    )
    s.rounded(draw, (180, 735, 1420, 820), fill=s.LIGHT_GREEN, outline=s.GREEN, width=2)
    s.centered_text(
        draw,
        (180, 735, 1420, 820),
        "测试结论：12/12 通过；冷启动 430 ms；新增任务保存反馈 74 ms。",
        s.font(25, bold=True),
        fill=s.GREEN,
    )
    s.save(image, path)


def build_all(repo_root: Path) -> list[Path]:
    s = load_showcase_helpers(repo_root)
    base = repo_root / "examples" / "template-examples"
    outputs: list[Path] = []

    engineering = base / "neutral-engineering-lab" / "assets"
    database_plan(s, engineering / "explain-plan.png")
    database_performance(s, engineering / "query-performance.png")
    outputs.extend(sorted(engineering.glob("*.png")))

    modern = base / "neutral-modern-minimal" / "assets"
    responsive_layout(s, modern / "responsive-layout.png")
    web_quality(s, modern / "quality-check.png")
    outputs.extend(sorted(modern.glob("*.png")))

    compact = base / "neutral-compact-header-lab" / "assets"
    s.terminal_image(
        "Ubuntu 24.04 — 共享目录权限验证",
        [
            ("$ sudo groupadd project", "#7DD3FC"),
            ("$ sudo usermod -aG project alice", s.TERMINAL_TEXT),
            ("$ sudo usermod -aG project bob", s.TERMINAL_TEXT),
            ("$ sudo chown root:project /srv/project", s.TERMINAL_TEXT),
            ("$ sudo chmod 2770 /srv/project", s.TERMINAL_TEXT),
            ("$ ls -ld /srv/project", "#7DD3FC"),
            ("drwxrws--- 2 root project 4096 Jun 12 14:20 /srv/project", "#86EFAC"),
            ("alice$ touch /srv/project/task.txt", "#7DD3FC"),
            ("bob$ echo verified >> /srv/project/task.txt", "#86EFAC"),
            ("carol$ cd /srv/project", "#7DD3FC"),
            ("bash: cd: /srv/project: Permission denied", "#FCA5A5"),
        ],
        compact / "permission-terminal.png",
    )
    outputs.extend(sorted(compact.glob("*.png")))

    review = base / "neutral-review-panel-lab" / "assets"
    junit_matrix(s, review / "boundary-tests.png")
    s.terminal_image(
        "Maven Surefire — JUnit 回归测试",
        [
            ("$ mvn -q test", "#7DD3FC"),
            ("Running GradeServiceParameterizedTest", s.TERMINAL_TEXT),
            ("Tests run: 14, Failures: 0, Errors: 0, Skipped: 0", "#86EFAC"),
            ("Running GradeServiceExceptionTest", s.TERMINAL_TEXT),
            ("Tests run: 2, Failures: 0, Errors: 0, Skipped: 0", "#86EFAC"),
            ("", s.TERMINAL_TEXT),
            ("Results:", "#7DD3FC"),
            ("Tests run: 16, Failures: 0, Errors: 0, Skipped: 0", "#86EFAC"),
            ("BUILD SUCCESS", "#86EFAC"),
            ("Branch coverage: 100%  |  Statement coverage: 100%", "#A5B4FC"),
        ],
        review / "junit-console.png",
    )
    outputs.extend(sorted(review.glob("*.png")))

    code = base / "neutral-code-notebook-lab" / "assets"
    merge_flow(s, code / "merge-sort-flow.png")
    benchmark_chart(s, code / "benchmark-chart.png")
    outputs.extend(sorted(code.glob("*.png")))

    data = base / "neutral-data-analysis-lab" / "assets"
    calibration_chart(s, data / "calibration-chart.png")
    headers = ["标准温度", "输出均值", "拟合温度", "绝对误差", "三次极差"]
    rows = [
        ("0 ℃", "0.500 V", "0.00 ℃", "0.00 ℃", "0.004 V"),
        ("10 ℃", "1.300 V", "10.00 ℃", "+0.00 ℃", "0.005 V"),
        ("20 ℃", "2.100 V", "20.00 ℃", "+0.00 ℃", "0.006 V"),
        ("25 ℃", "2.500 V", "25.00 ℃", "+0.00 ℃", "0.005 V"),
        ("40 ℃", "3.700 V", "40.00 ℃", "+0.00 ℃", "0.008 V"),
        ("50 ℃", "4.505 V", "50.05 ℃", "+0.05 ℃", "0.007 V"),
    ]
    image, draw = s.new_canvas(
        "温度传感器测量记录",
        "每个温度点稳定三分钟后连续测量三次，表中列出处理后均值",
        accent=s.GREEN,
    )
    s.draw_table(
        draw,
        (90, 225, 1510, 760),
        headers,
        rows,
        widths=[0.18, 0.20, 0.20, 0.20, 0.22],
        header_fill=s.LIGHT_GREEN,
        font_size=22,
    )
    s.save(image, data / "measurement-record.png")
    outputs.extend(sorted(data.glob("*.png")))

    project = base / "neutral-project-dossier" / "assets"
    android_architecture(s, project / "application-architecture.png")
    android_ui(s, project / "task-list-ui.png")
    android_tests(s, project / "test-results.png")
    outputs.extend(sorted(project.glob("*.png")))
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
    )
    args = parser.parse_args()
    for output in build_all(args.repo_root.resolve()):
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
