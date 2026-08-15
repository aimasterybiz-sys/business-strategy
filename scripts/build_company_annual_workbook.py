"""Сборка демо-workbook «Годовой обзор деятельности компании 2025–2026».

Стиль: структура BIS (панели A/B/C, «Graph N», ось Y справа, единицы в правом
верхнем углу, подвал курсивом) + визуал LVMH (кремовый фон, serif italic,
приглушённая палитра, плоский стиль без теней).

Все графики — нативные объекты Excel (openpyxl), серии берутся из диапазонов
ячеек через Reference. Тепловая карта — условное форматирование ColorScaleRule.

Запуск:  python3 scripts/build_company_annual_workbook.py
Вывод:   Company_Annual_BIS_Style.xlsx
"""

from __future__ import annotations

import csv
import pathlib

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawFont,
    Paragraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "Company_Annual_BIS_Style.xlsx"

# ── Палитра (LVMH) ──────────────────────────────────────────────────────────
CREAM = "F5EDE6"          # фон
INK = "3A322C"            # основной текст
PRIMARY = "1F3A5F"        # выручка / основная метрика
SECONDARY = "4A7BA7"      # трафик и охват
TERTIARY = "8FB8D9"       # конверсия и воронка
ACCENT = "7A9A7E"         # доля рынка
ACCENT2 = "C97B5A"        # конкуренты
SHADE = "EDCBC1"          # затенённый период — пыльно-розовый
GRID = "E3D9CE"
AXIS = "C0B3A6"
RULE = "9A8B7C"

SERIF = "Georgia"

# ── Раскладка листа ─────────────────────────────────────────────────────────
ROW_SUB = {"A": 4, "B": 34, "C": 64}
ANCHOR = {"A": "A5", "B": "A35", "C": "A65"}
ROW_FOOTER = 95
ROW_DATA = 101
UNIT_COL = 19  # столбец S — правый край графиков

WIDE = (35, 14)   # 365-дневные линии
NARROW = (28, 13.5)

CAMPAIGNS = [(45, 60), (210, 228)]
SHOCKS = [(132, 139), (300, 306)]


# ── Утилиты чтения данных ───────────────────────────────────────────────────
def read_csv(name: str) -> list[dict]:
    with (DATA / name).open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def num(value: str) -> float:
    return float(value)


# ── Стилевые помощники ──────────────────────────────────────────────────────
def chart_text(size: int = 800, italic: bool = True, bold: bool = False) -> RichText:
    cp = CharacterProperties(
        latin=DrawFont(typeface=SERIF), sz=size * 1, i=italic, b=bold, solidFill=INK
    )
    return RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)],
    )


def flat_area(color: str) -> GraphicalProperties:
    return GraphicalProperties(solidFill=color, ln=LineProperties(noFill=True))


def dress_chart(chart, size, y_fmt="#,##0", legend="b"):
    """Кремовый фон, плоский стиль, ось Y справа, serif italic, без теней."""
    chart.width, chart.height = size
    chart.graphical_properties = flat_area(CREAM)
    chart.plot_area.graphicalProperties = flat_area(CREAM)
    chart.roundedCorners = False

    for axis in (chart.x_axis, chart.y_axis):
        axis.delete = False
        axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=6350))
        axis.txPr = chart_text(800)
        axis.majorTickMark = "out"
        axis.minorTickMark = "none"

    # BIS: ось значений справа.
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "r"
    chart.y_axis.crosses = "max"
    chart.y_axis.numFmt = y_fmt
    chart.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID, w=6350))
    )
    chart.x_axis.majorGridlines = None
    chart.x_axis.tickLblPos = "low"

    if legend is None:
        chart.legend = None
    else:
        chart.legend.position = legend
        chart.legend.overlay = False
        chart.legend.txPr = chart_text(800)

    chart.dispBlanksAs = "gap"
    return chart


def style_line(series, color: str, width_pt: float = 1.25, dashed: bool = False):
    lp = LineProperties(solidFill=color, w=int(width_pt * 12700))
    if dashed:
        lp.prstDash = "dash"
    series.graphicalProperties = GraphicalProperties(ln=lp)
    series.smooth = False
    series.marker = Marker(symbol="none")
    return series


def style_points(series, color: str, size: int = 5, symbol: str = "circle"):
    series.marker = Marker(
        symbol=symbol,
        size=size,
        spPr=GraphicalProperties(solidFill=color, ln=LineProperties(noFill=True)),
    )
    series.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    return series


def apply_cats(chart, ws, table, header: str):
    """Текстовые категории через StrRef — иначе Excel показывает 1…365 вместо дат."""
    col = get_column_letter(table["cols"][header])
    src = AxDataSource(
        strRef=StrRef(f=f"'{ws.title}'!${col}${table['first']}:${col}${table['last']}")
    )
    for series in chart.series:
        series.cat = src
    return chart


def dress_secondary(axis, fmt: str = "#,##0"):
    """Вторая ось значений — слева, без сетки, тем же serif italic."""
    axis.delete = False
    axis.axPos = "l"
    axis.crosses = "autoZero"
    axis.numFmt = fmt
    axis.txPr = chart_text(800)
    axis.majorGridlines = None
    axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=AXIS, w=6350))
    axis.majorTickMark = "out"
    axis.minorTickMark = "none"
    return axis


def style_bars(chart, colors):
    for series, color in zip(chart.series, colors):
        series.graphicalProperties = flat_area(color)


# ── Каркас листа ────────────────────────────────────────────────────────────
def prepare_sheet(wb, name: str, title: str, lead: str, graph_no: int, last_row: int):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = CREAM

    ws.column_dimensions["A"].width = 12
    for col in range(2, 27):
        ws.column_dimensions[get_column_letter(col)].width = 9.5

    fill = PatternFill("solid", fgColor=CREAM)
    for row in range(1, last_row + 4):
        for col in range(1, 27):
            ws.cell(row=row, column=col).fill = fill

    ws["A1"] = title
    ws["A1"].font = Font(name=SERIF, size=17, bold=True, color=INK)
    ws["A2"] = lead
    ws["A2"].font = Font(name=SERIF, size=10, italic=True, color=RULE)

    cell = ws.cell(row=1, column=UNIT_COL, value=f"Graph {graph_no}")
    cell.font = Font(name=SERIF, size=12, bold=True, italic=True, color=RULE)
    cell.alignment = Alignment(horizontal="right")

    for col in range(1, UNIT_COL + 1):
        ws.cell(row=3, column=col).border = Border(
            bottom=Side(style="thin", color=RULE)
        )
    return ws


def panel_header(ws, panel: str, subtitle: str, unit: str):
    row = ROW_SUB[panel]
    cell = ws.cell(row=row, column=1, value=f"{panel}. {subtitle}")
    cell.font = Font(name=SERIF, size=11, bold=True, color=INK)
    unit_cell = ws.cell(row=row, column=UNIT_COL, value=unit)
    unit_cell.font = Font(name=SERIF, size=9, italic=True, color=RULE)
    unit_cell.alignment = Alignment(horizontal="right")


def footer(ws, lines: list[str]):
    for offset, line in enumerate(lines):
        cell = ws.cell(row=ROW_FOOTER + offset, column=1, value=line)
        cell.font = Font(name=SERIF, size=9, italic=True, color=RULE)


def write_table(ws, top: int, start_col: int, headers: list[str], rows: list[list]):
    """Пишет таблицу данных и возвращает координаты для Reference."""
    for offset, head in enumerate(headers):
        cell = ws.cell(row=top, column=start_col + offset, value=head)
        cell.font = Font(name=SERIF, size=9, bold=True, italic=True, color=INK)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color=RULE))
    for r, row in enumerate(rows, start=top + 1):
        for offset, value in enumerate(row):
            cell = ws.cell(row=r, column=start_col + offset, value=value)
            cell.font = Font(name=SERIF, size=9, color=INK)
            if isinstance(value, (int, float)):
                cell.number_format = "#,##0.##"
                cell.alignment = Alignment(horizontal="right")
    return {
        "hdr": top,
        "first": top + 1,
        "last": top + len(rows),
        "col": start_col,
        "cols": {h: start_col + i for i, h in enumerate(headers)},
    }


def ref(ws, table, header: str, with_header=True):
    col = table["cols"][header]
    return Reference(
        ws,
        min_col=col,
        min_row=table["hdr"] if with_header else table["first"],
        max_row=table["last"],
    )


def shade_column(day: int, ranges, value: float):
    return value if any(lo <= day <= hi for lo, hi in ranges) else None


# ── Лист 1: Мониторинг рынка ────────────────────────────────────────────────
def sheet_monitoring(wb, daily, monthly):
    ws = prepare_sheet(
        wb,
        "1. Мониторинг рынка",
        "Годовой обзор деятельности компании 2025–2026",
        "Мониторинг рынка: спрос, трафик и воронка сделок, май 2025 — апрель 2026",
        1,
        ROW_DATA + len(daily),
    )

    def index_of(key):
        values = [num(r[key]) for r in daily]
        base = sum(values[:28]) / 28
        return [round(v / base * 100, 1) for v in values]

    demand, traffic, leads, deals = (
        index_of("DemandIndex"),
        index_of("Traffic"),
        index_of("Leads"),
        index_of("Deals"),
    )

    rows = []
    for i, r in enumerate(daily):
        day = int(r["DayIndex"])
        rows.append(
            [
                r["Date"],
                demand[i],
                traffic[i],
                leads[i],
                deals[i],
                shade_column(day, CAMPAIGNS, 175),
            ]
        )
    t_daily = write_table(
        ws,
        ROW_DATA,
        1,
        ["Дата", "Спрос", "Трафик", "Лиды", "Сделки", "Кампания"],
        rows,
    )

    t_month = write_table(
        ws,
        ROW_DATA,
        8,
        ["Месяц", "Органика", "Платный", "Прямой"],
        [
            [m["Month"], int(m["Organic"]), int(m["Paid"]), int(m["Direct"])]
            for m in monthly
        ],
    )

    # Регрессия «расходы → выручка» с 95%-й доверительной полосой.
    x = np.array([num(r["MarketingSpend"]) for r in daily])
    y = np.array([num(r["Revenue"]) for r in daily])
    slope, intercept = np.polyfit(x, y, 1)
    resid = y - (slope * x + intercept)
    s_err = float(np.sqrt((resid**2).sum() / (len(x) - 2)))
    xbar, sxx = float(x.mean()), float(((x - x.mean()) ** 2).sum())
    order = np.argsort(x)
    band_rows = []
    for xi in x[order]:
        fit = slope * xi + intercept
        se = s_err * float(np.sqrt(1 / len(x) + (xi - xbar) ** 2 / sxx))
        band_rows.append(
            [
                round(float(xi), 1),
                round(float(fit), 1),
                round(float(fit + 1.96 * se), 1),
                round(float(fit - 1.96 * se), 1),
            ]
        )

    t_sc = write_table(
        ws,
        ROW_DATA,
        13,
        ["Расходы", "Выручка"],
        [[round(float(a), 1), round(float(b), 1)] for a, b in zip(x, y)],
    )
    t_band = write_table(
        ws, ROW_DATA, 16, ["X", "Регрессия", "Верх 95%", "Низ 95%"], band_rows
    )

    # A. Четыре индексные линии + затенение кампаний.
    panel_header(ws, "A", "Спрос, трафик и воронка сделок, 365 дней", "индекс, 100 = май 2025")
    area = AreaChart()
    area.add_data(ref(ws, t_daily, "Кампания"), titles_from_data=True)
    apply_cats(area, ws, t_daily, "Дата")
    area.series[0].graphicalProperties = flat_area(SHADE)
    line = LineChart()
    for name in ("Спрос", "Трафик", "Лиды", "Сделки"):
        line.add_data(ref(ws, t_daily, name), titles_from_data=True)
    apply_cats(line, ws, t_daily, "Дата")
    for series, color in zip(line.series, (PRIMARY, SECONDARY, TERTIARY, ACCENT)):
        style_line(series, color, 1.1)
    area += line
    combo = area
    dress_chart(combo, WIDE)
    combo.y_axis.scaling.min = 0
    combo.x_axis.tickLblSkip = 30
    combo.x_axis.tickMarkSkip = 30
    ws.add_chart(combo, ANCHOR["A"])

    # B. Стековые столбцы по каналам.
    panel_header(ws, "B", "Помесячная структура спроса по каналам", "визиты")
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap = "col", "stacked", 100
    for name in ("Органика", "Платный", "Прямой"):
        bar.add_data(ref(ws, t_month, name), titles_from_data=True)
    apply_cats(bar, ws, t_month, "Месяц")
    style_bars(bar, (PRIMARY, SECONDARY, TERTIARY))
    dress_chart(bar, NARROW)
    ws.add_chart(bar, ANCHOR["B"])

    # C. Диаграмма рассеяния + регрессия + доверительная полоса.
    panel_header(ws, "C", "Маркетинговые расходы против выручки, N = 365", "₽ тыс.")
    sc = ScatterChart()
    sc.x_axis.title = None
    xref = ref(ws, t_sc, "Расходы", with_header=False)
    points = Series(ref(ws, t_sc, "Выручка"), xref, title_from_data=True)
    style_points(points, SECONDARY, 5)
    points.trendline = Trendline(trendlineType="linear")
    sc.series.append(points)

    bx = ref(ws, t_band, "X", with_header=False)
    for name, color, dashed in (
        ("Регрессия", PRIMARY, False),
        ("Верх 95%", ACCENT2, True),
        ("Низ 95%", ACCENT2, True),
    ):
        s = Series(ref(ws, t_band, name), bx, title_from_data=True)
        style_line(s, color, 1.4 if not dashed else 1.0, dashed=dashed)
        sc.series.append(s)
    dress_chart(sc, NARROW)
    ws.add_chart(sc, ANCHOR["C"])

    footer(
        ws,
        [
            "Затенённые области отмечают периоды маркетинговых кампаний.",
            "ª Кампания 1 (июн. 2025).   ᵇ Кампания 2 (янв. 2026).",
            "Панель C: сплошная линия — линейная регрессия, пунктир — 95 % доверительный интервал среднего.",
            "Источники: внутренняя аналитика и открытые данные по конкурентам (демо-данные), 2025-05 … 2026-05.",
        ],
    )


# ── Лист 2: Тренды ──────────────────────────────────────────────────────────
def sheet_trends(wb, daily, monthly):
    ws = prepare_sheet(
        wb,
        "2. Тренды",
        "Годовой обзор деятельности компании 2025–2026",
        "Тренды спроса: сглаженная динамика, прогноз по категориям, сезонность",
        2,
        ROW_DATA + len(daily),
    )

    rows = []
    for r in daily:
        day = int(r["DayIndex"])
        rows.append(
            [
                r["Date"],
                num(r["DemandIndex"]),
                shade_column(day, CAMPAIGNS + SHOCKS, 100),
            ]
        )
    t_daily = write_table(ws, ROW_DATA, 1, ["Дата", "Спрос", "Период"], rows)

    categories = [
        ("Корпоративный сегмент", 8.4, 2.1),
        ("Малый бизнес", 5.6, 1.6),
        ("Госсектор", 2.1, 2.8),
        ("Розница", 6.9, 1.9),
        ("Партнёрский канал", 11.2, 3.4),
        ("Экспорт", -1.8, 2.6),
        ("Подписка", 14.5, 2.2),
        ("Услуги внедрения", 3.7, 1.5),
    ]
    t_fc = write_table(
        ws,
        ROW_DATA,
        5,
        ["Категория", "N", "Прогноз", "ДИ"],
        [[name, i + 1, g, ci] for i, (name, g, ci) in enumerate(categories)],
    )

    t_month = write_table(
        ws,
        ROW_DATA,
        10,
        ["Месяц", "Сезонность", "Средний чек"],
        [[m["Month"], num(m["Seasonality"]), int(m["AvgCheck"])] for m in monthly],
    )

    # A. Спрос + сглаживание movingAvg 28 + затенение всех периодов.
    panel_header(ws, "A", "Индекс спроса и сглаженный тренд, 365 дней", "индекс")
    area = AreaChart()
    area.add_data(ref(ws, t_daily, "Период"), titles_from_data=True)
    apply_cats(area, ws, t_daily, "Дата")
    area.series[0].graphicalProperties = flat_area(SHADE)
    line = LineChart()
    line.add_data(ref(ws, t_daily, "Спрос"), titles_from_data=True)
    apply_cats(line, ws, t_daily, "Дата")
    style_line(line.series[0], SECONDARY, 1.0)
    line.series[0].trendline = Trendline(
        trendlineType="movingAvg",
        period=28,
        spPr=GraphicalProperties(ln=LineProperties(solidFill=PRIMARY, w=25400)),
    )
    area += line
    combo = area
    dress_chart(combo, WIDE)
    combo.y_axis.scaling.min = 0
    combo.x_axis.tickLblSkip = 30
    combo.x_axis.tickMarkSkip = 30
    ws.add_chart(combo, ANCHOR["A"])

    # B. Прогноз роста по категориям — scatter с планками погрешностей.
    panel_header(ws, "B", "Прогноз роста по категориям, ±95 % ДИ", "% г/г")
    sc = ScatterChart()
    palette = [PRIMARY, SECONDARY, TERTIARY, ACCENT, ACCENT2, PRIMARY, SECONDARY, ACCENT]
    xcol = get_column_letter(t_fc["cols"]["N"])
    ycol = get_column_letter(t_fc["cols"]["Прогноз"])
    cicol = get_column_letter(t_fc["cols"]["ДИ"])
    title_col = get_column_letter(t_fc["cols"]["Категория"])
    for i in range(len(categories)):
        row = t_fc["first"] + i
        xr = Reference(ws, min_col=t_fc["cols"]["N"], min_row=row, max_row=row)
        yr = Reference(ws, min_col=t_fc["cols"]["Прогноз"], min_row=row, max_row=row)
        s = Series(yr, xr, title=categories[i][0])
        style_points(s, palette[i], 8, "diamond")
        cell = f"'{ws.title}'!${cicol}${row}"
        s.errBars = ErrorBars(
            errDir="y",
            errBarType="both",
            errValType="cust",
            noEndCap=False,
            plus=NumDataSource(NumRef(f=cell)),
            minus=NumDataSource(NumRef(f=cell)),
            spPr=GraphicalProperties(ln=LineProperties(solidFill=RULE, w=9525)),
        )
        s.dLbls = DataLabelList(
            showVal=True,
            showSerName=False,
            showCatName=False,
            showLegendKey=False,
            showBubbleSize=False,
            numFmt='+#,##0.0;-#,##0.0',
            txPr=chart_text(800),
        )
        sc.series.append(s)
    dress_chart(sc, NARROW, y_fmt='+#,##0.0;-#,##0.0', legend="r")
    sc.x_axis.scaling.min = 0
    sc.x_axis.scaling.max = len(categories) + 1
    ws.add_chart(sc, ANCHOR["B"])

    # C. Сезонность и средний чек — сгруппированные столбцы, вторая ось.
    panel_header(ws, "C", "Сезонность спроса и средний чек по месяцам", "индекс / ₽")
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap = "col", "clustered", -20
    bar.add_data(ref(ws, t_month, "Сезонность"), titles_from_data=True)
    apply_cats(bar, ws, t_month, "Месяц")
    style_bars(bar, (SECONDARY,))
    check = BarChart()
    check.type, check.grouping, check.overlap = "col", "clustered", -20
    check.add_data(ref(ws, t_month, "Средний чек"), titles_from_data=True)
    style_bars(check, (ACCENT,))
    check.y_axis.axId = 200
    check.y_axis.title = None
    dress_chart(bar, NARROW)
    dress_secondary(check.y_axis, "#,##0")
    bar += check
    combo = bar
    ws.add_chart(combo, ANCHOR["C"])

    footer(
        ws,
        [
            "Затенённые области отмечают периоды кампаний или рыночных шоков.",
            "ª Кампания 1 (июн. 2025).   ᵇ Кампания 2 (янв. 2026).",
            "ᶜ Рыночный шок 1 (сен. 2025).   ᵈ Рыночный шок 2 (мар. 2026).",
            "Панель A: жирная линия — скользящее среднее за 28 дней. Панель B: планки — 95 % доверительный интервал прогноза.",
            "Источники: внутренняя аналитика и открытые данные по конкурентам (демо-данные), 2025-05 … 2026-05.",
        ],
    )


# ── Лист 3: Конкуренты ──────────────────────────────────────────────────────
def sheet_competitors(wb, daily, monthly):
    ws = prepare_sheet(
        wb,
        "3. Конкуренты",
        "Годовой обзор деятельности компании 2025–2026",
        "Конкурентное поле: доля голоса, ценовая позиция, доля рынка и упоминания",
        3,
        ROW_DATA + len(daily),
    )

    rows = []
    for r in daily:
        day = int(r["DayIndex"])
        rows.append(
            [
                r["Date"],
                num(r["SOV_Us"]),
                num(r["SOV_CompA"]),
                num(r["SOV_CompB"]),
                num(r["SOV_CompC"]),
                shade_column(day, SHOCKS, 35),
                num(r["PriceIndexUs"]),
                num(r["PriceIndexMarket"]),
                round(num(r["PriceIndexUs"]) - num(r["PriceIndexMarket"]), 2),
            ]
        )
    t_daily = write_table(
        ws,
        ROW_DATA,
        1,
        [
            "Дата",
            "Мы",
            "Конкурент A",
            "Конкурент B",
            "Конкурент C",
            "Шок",
            "Цена: мы",
            "Цена: рынок",
            "Разница",
        ],
        rows,
    )

    t_month = write_table(
        ws,
        ROW_DATA,
        11,
        ["Месяц", "Доля: мы", "Доля: A", "Доля: B", "Доля: C", "Упоминания: мы"],
        [
            [
                m["Month"],
                num(m["MarketShareUs"]),
                num(m["MarketShareA"]),
                num(m["MarketShareB"]),
                num(m["MarketShareC"]),
                int(m["MentionsUs"]),
            ]
            for m in monthly
        ],
    )

    # A. Доля голоса: мы и три конкурента.
    panel_header(ws, "A", "Доля голоса: мы против трёх конкурентов, 365 дней", "% упоминаний")
    area = AreaChart()
    area.add_data(ref(ws, t_daily, "Шок"), titles_from_data=True)
    apply_cats(area, ws, t_daily, "Дата")
    area.series[0].graphicalProperties = flat_area(SHADE)
    line = LineChart()
    for name in ("Мы", "Конкурент A", "Конкурент B", "Конкурент C"):
        line.add_data(ref(ws, t_daily, name), titles_from_data=True)
    apply_cats(line, ws, t_daily, "Дата")
    for series, color, width in zip(
        line.series, (PRIMARY, ACCENT2, SECONDARY, TERTIARY), (1.8, 1.0, 1.0, 1.0)
    ):
        style_line(series, color, width)
    area += line
    combo = area
    dress_chart(combo, WIDE, y_fmt="#,##0.0")
    combo.y_axis.scaling.min = 0
    combo.x_axis.tickLblSkip = 30
    combo.x_axis.tickMarkSkip = 30
    ws.add_chart(combo, ANCHOR["A"])

    # B. Ценовая позиция и разница в пунктах — вторая ось.
    panel_header(ws, "B", "Ценовой индекс: мы против средней по рынку", "индекс / п.п.")
    line = LineChart()
    for name in ("Цена: мы", "Цена: рынок"):
        line.add_data(ref(ws, t_daily, name), titles_from_data=True)
    apply_cats(line, ws, t_daily, "Дата")
    style_line(line.series[0], PRIMARY, 1.5)
    style_line(line.series[1], ACCENT2, 1.2, dashed=True)
    diff = LineChart()
    diff.add_data(ref(ws, t_daily, "Разница"), titles_from_data=True)
    apply_cats(diff, ws, t_daily, "Дата")
    style_line(diff.series[0], ACCENT, 1.0)
    diff.y_axis.axId = 200
    dress_chart(line, NARROW, y_fmt="#,##0.0")
    line.x_axis.tickLblSkip = 45
    line.x_axis.tickMarkSkip = 45
    dress_secondary(diff.y_axis, '+#,##0.0;-#,##0.0')
    line += diff
    ws.add_chart(line, ANCHOR["B"])

    # C. Доля рынка по игрокам + упоминания на второй оси.
    panel_header(ws, "C", "Доля рынка по игрокам и упоминания бренда", "% / шт.")
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap = "col", "clustered", -15
    for name in ("Доля: мы", "Доля: A", "Доля: B", "Доля: C"):
        bar.add_data(ref(ws, t_month, name), titles_from_data=True)
    apply_cats(bar, ws, t_month, "Месяц")
    style_bars(bar, (PRIMARY, ACCENT2, SECONDARY, TERTIARY))
    mentions = LineChart()
    mentions.add_data(ref(ws, t_month, "Упоминания: мы"), titles_from_data=True)
    style_line(mentions.series[0], ACCENT, 2.0)
    mentions.y_axis.axId = 200
    dress_chart(bar, NARROW, y_fmt="#,##0.0")
    dress_secondary(mentions.y_axis, "#,##0")
    bar += mentions
    ws.add_chart(bar, ANCHOR["C"])

    footer(
        ws,
        [
            "Затенённые области отмечают периоды рыночных шоков.",
            "ᶜ Рыночный шок 1 (сен. 2025).   ᵈ Рыночный шок 2 (мар. 2026).",
            "Панель B: правая ось — разница между нашим ценовым индексом и средней по рынку, в процентных пунктах.",
            "Источники: внутренняя аналитика и открытые данные по конкурентам (демо-данные), 2025-05 … 2026-05.",
        ],
    )


# ── Лист 4: Цифровой след ───────────────────────────────────────────────────
KPI_COLUMNS = [
    ("KPI_Revenue", "Выручка, ₽ тыс.", False),
    ("KPI_Traffic", "Трафик, визиты", False),
    ("KPI_Conversion", "Конверсия, %", False),
    ("KPI_CPL", "CPL, ₽", True),  # меньше — лучше
    ("KPI_Share", "Доля голоса, %", False),
    ("KPI_NPS", "NPS", False),
]


def sheet_digital(wb, daily, monthly):
    ws = prepare_sheet(
        wb,
        "4. Цифровой след",
        "Годовой обзор деятельности компании 2025–2026",
        "Цифровой след: трафик и видимость, экономика лида, помесячная карта KPI",
        4,
        ROW_DATA + len(daily),
    )

    rows = [
        [
            r["Date"],
            int(r["Traffic"]),
            num(r["SearchVisibility"]),
            num(r["ConversionRate"]),
            int(r["CPL"]),
        ]
        for r in daily
    ]
    t_daily = write_table(
        ws,
        ROW_DATA,
        1,
        ["Дата", "Визиты", "Видимость", "Конверсия", "CPL"],
        rows,
    )

    # A. Визиты + видимость в поиске (вторая ось).
    panel_header(ws, "A", "Визиты сайта и видимость в поиске, 365 дней", "визиты / %")
    visits = LineChart()
    visits.add_data(ref(ws, t_daily, "Визиты"), titles_from_data=True)
    apply_cats(visits, ws, t_daily, "Дата")
    style_line(visits.series[0], SECONDARY, 1.0)
    vis = LineChart()
    vis.add_data(ref(ws, t_daily, "Видимость"), titles_from_data=True)
    apply_cats(vis, ws, t_daily, "Дата")
    style_line(vis.series[0], PRIMARY, 1.8)
    vis.y_axis.axId = 200
    dress_chart(visits, WIDE)
    visits.x_axis.tickLblSkip = 30
    visits.x_axis.tickMarkSkip = 30
    dress_secondary(vis.y_axis, "#,##0.0")
    visits += vis
    ws.add_chart(visits, ANCHOR["A"])

    # B. Конверсия воронки и стоимость лида (вторая ось).
    panel_header(ws, "B", "Конверсия воронки и стоимость лида", "% / ₽")
    conv = LineChart()
    conv.add_data(ref(ws, t_daily, "Конверсия"), titles_from_data=True)
    apply_cats(conv, ws, t_daily, "Дата")
    style_line(conv.series[0], TERTIARY, 1.2)
    conv.series[0].trendline = Trendline(
        trendlineType="movingAvg",
        period=28,
        spPr=GraphicalProperties(ln=LineProperties(solidFill=PRIMARY, w=25400)),
    )
    cpl = LineChart()
    cpl.add_data(ref(ws, t_daily, "CPL"), titles_from_data=True)
    apply_cats(cpl, ws, t_daily, "Дата")
    style_line(cpl.series[0], ACCENT2, 1.0)
    cpl.y_axis.axId = 200
    dress_chart(conv, NARROW, y_fmt="#,##0.00")
    conv.x_axis.tickLblSkip = 45
    conv.x_axis.tickMarkSkip = 45
    dress_secondary(cpl.y_axis, "#,##0")
    conv += cpl
    ws.add_chart(conv, ANCHOR["B"])

    # C. Тепловая карта KPI — условное форматирование, не объект-график.
    panel_header(ws, "C", "Помесячная карта KPI, перцентильная раскраска", "условное форматирование")
    top = ROW_SUB["C"] + 2
    headers = ["Месяц"] + [label for _, label, _ in KPI_COLUMNS]
    body = [
        [m["Month"]] + [num(m[key]) for key, _, _ in KPI_COLUMNS] for m in monthly
    ]
    t_kpi = write_table(ws, top, 1, headers, body)
    for cell_row in range(t_kpi["first"], t_kpi["last"] + 1):
        ws.cell(row=cell_row, column=1).font = Font(name=SERIF, size=9, italic=True, color=INK)

    for offset, (_, label, reverse) in enumerate(KPI_COLUMNS):
        col = get_column_letter(t_kpi["cols"][label])
        rng = f"{col}{t_kpi['first']}:{col}{t_kpi['last']}"
        low, high = ("C97B5A", "1F3A5F") if not reverse else ("1F3A5F", "C97B5A")
        ws.conditional_formatting.add(
            rng,
            ColorScaleRule(
                start_type="percentile", start_value=5, start_color=low,
                mid_type="percentile", mid_value=50, mid_color="EDCBC1",
                end_type="percentile", end_value=95, end_color=high,
            ),
        )
        for cell_row in range(t_kpi["first"], t_kpi["last"] + 1):
            cell = ws.cell(row=cell_row, column=t_kpi["cols"][label])
            cell.number_format = "#,##0.##"
            cell.alignment = Alignment(horizontal="center")

    footer(
        ws,
        [
            "Панель C — тепловая карта на условном форматировании (ColorScaleRule, перцентили 5 / 50 / 95), не объект-график.",
            "Для CPL шкала инвертирована: меньшее значение окрашено как лучший результат.",
            "Панель B: жирная линия — скользящее среднее конверсии за 28 дней.",
            "Источники: внутренняя аналитика и открытые данные по конкурентам (демо-данные), 2025-05 … 2026-05.",
        ],
    )


def main() -> None:
    daily = read_csv("company_daily_2025_2026.csv")
    monthly = read_csv("company_monthly_summary.csv")

    wb = Workbook()
    wb.remove(wb.active)
    sheet_monitoring(wb, daily, monthly)
    sheet_trends(wb, daily, monthly)
    sheet_competitors(wb, daily, monthly)
    sheet_digital(wb, daily, monthly)
    wb.save(OUT)
    print(f"Сохранено: {OUT.name}")


if __name__ == "__main__":
    main()
