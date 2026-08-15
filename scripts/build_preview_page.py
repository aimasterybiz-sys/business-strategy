"""Веб-превью workbook «Годовой обзор деятельности компании 2025–2026».

Читает те же CSV, что и build_company_annual_workbook.py, и рисует те же
12 панелей инлайновым SVG — чтобы посмотреть дашборд без Excel.

Запуск:  python3 scripts/build_preview_page.py
Вывод:   preview/dashboard.html
"""

from __future__ import annotations

import contextlib
import csv
import html
import json
import math
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "preview" / "dashboard.html"

CREAM = "#F5EDE6"
PAPER_EDGE = "#E4D7CA"
INK = "#3A322C"
MUTED = "#8B7B6B"
GRID = "#E3D9CE"
AXIS = "#C0B3A6"

PRIMARY = "#1F3A5F"
SECONDARY = "#4A7BA7"
TERTIARY = "#8FB8D9"
ACCENT = "#3F6B52"
ACCENT2 = "#C97B5A"
SHADE = "#EDCBC1"
RAMP = ["#E8EEF4", "#7FA3C4", "#1F3A5F"]

CAMPAIGNS = [(45, 60), (210, 228)]
SHOCKS = [(132, 139), (300, 306)]

VB_W = 1160
VB_H = 300
ML, MR, MT, MB = 12, 74, 16, 38
GUTTER = 52          # левый отступ под подписи второй оси
PX0, PX1 = ML, VB_W - MR
PY0, PY1 = MT, VB_H - MB


@contextlib.contextmanager
def gutter(active: bool):
    """Панели с двумя шкалами получают поле слева — иначе подписи второй оси
    уезжают за край области рисования."""
    global PX0
    PX0 = ML + (GUTTER if active else 0)
    try:
        yield
    finally:
        PX0 = ML

MONTH_RU = {
    "01": "янв", "02": "фев", "03": "мар", "04": "апр", "05": "май", "06": "июн",
    "07": "июл", "08": "авг", "09": "сен", "10": "окт", "11": "ноя", "12": "дек",
}


# ── данные ──────────────────────────────────────────────────────────────────
def read_csv(name: str) -> list[dict]:
    with (DATA / name).open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def fmt(value: float, dec: int = 0) -> str:
    return f"{value:,.{dec}f}".replace(",", " ")


def month_label(iso: str) -> str:
    return f"{MONTH_RU[iso[5:7]]} {iso[2:4]}"


def nice_ticks(lo: float, hi: float, count: int = 5) -> list[float]:
    """Деления оси. Магнитуда через log10 — вариант «по длине str(int())»
    ломался на диапазонах меньше единицы и выдавал по 20 подписей."""
    if hi <= lo:
        hi = lo + 1
    raw = (hi - lo) / count
    mag = 10 ** math.floor(math.log10(raw))
    for multiple in (1, 2, 2.5, 5):
        if raw <= multiple * mag:
            step = multiple * mag
            break
    else:
        step = 10 * mag
    ticks, value = [], math.ceil(lo / step) * step
    while value <= hi + step * 1e-9:
        ticks.append(round(value, 10))
        value += step
    return ticks


def moving_avg(values: list[float], window: int) -> list[float | None]:
    out: list[float | None] = []
    for i in range(len(values)):
        if i + 1 < window:
            out.append(None)
        else:
            out.append(sum(values[i + 1 - window : i + 1]) / window)
    return out


# ── примитивы SVG ───────────────────────────────────────────────────────────
def esc(text: str) -> str:
    return html.escape(str(text), quote=True)


class Scale:
    def __init__(self, lo: float, hi: float, y0: float = PY0, y1: float = PY1):
        pad = (hi - lo) * 0.06 or 1
        self.lo, self.hi = lo - pad, hi + pad
        self.y0, self.y1 = y0, y1

    def y(self, value: float) -> float:
        span = self.hi - self.lo or 1
        return self.y1 - (value - self.lo) / span * (self.y1 - self.y0)


def band_x(index: int, count: int) -> float:
    if count <= 1:
        return (PX0 + PX1) / 2
    return PX0 + index / (count - 1) * (PX1 - PX0)


def frame(scale: Scale, dec: int = 0, side: str = "right") -> str:
    parts = []
    for tick in nice_ticks(scale.lo, scale.hi):
        y = scale.y(tick)
        if not (PY0 - 1 <= y <= PY1 + 1):
            continue
        parts.append(
            f'<line class="grid" x1="{PX0}" x2="{PX1}" y1="{y:.1f}" y2="{y:.1f}"/>'
        )
        if side == "right":
            parts.append(
                f'<text class="tick" x="{PX1 + 8}" y="{y + 3.5:.1f}">{fmt(tick, dec)}</text>'
            )
        else:
            parts.append(
                f'<text class="tick tick-l" x="{PX0 - 8}" y="{y + 3.5:.1f}">{fmt(tick, dec)}</text>'
            )
    parts.append(f'<line class="axis" x1="{PX0}" x2="{PX1}" y1="{PY1}" y2="{PY1}"/>')
    return "".join(parts)


def shade_bands(ranges, count: int) -> str:
    out = []
    for lo, hi in ranges:
        x1 = band_x(lo - 1, count)
        x2 = band_x(hi - 1, count)
        out.append(
            f'<rect class="shade" x="{x1:.1f}" y="{PY0}" '
            f'width="{max(x2 - x1, 1.5):.1f}" height="{PY1 - PY0}"/>'
        )
    return "".join(out)


def date_ticks(dates: list[str]) -> str:
    out = []
    seen = set()
    for i, iso in enumerate(dates):
        key = iso[:7]
        if key in seen:
            continue
        seen.add(key)
        x = band_x(i, len(dates))
        out.append(f'<line class="axis" x1="{x:.1f}" x2="{x:.1f}" y1="{PY1}" y2="{PY1 + 4}"/>')
        out.append(
            f'<text class="tick tick-x" x="{x:.1f}" y="{PY1 + 16}">{month_label(iso)}</text>'
        )
    return "".join(out)


def polyline(values, scale: Scale, count: int, color: str, width: float = 1.4,
             dash: str | None = None) -> str:
    chunks, current = [], []
    for i, value in enumerate(values):
        if value is None:
            if len(current) > 1:
                chunks.append(current)
            current = []
            continue
        current.append(f"{band_x(i, count):.1f},{scale.y(value):.1f}")
    if len(current) > 1:
        chunks.append(current)
    dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
    return "".join(
        f'<polyline fill="none" stroke="{color}" stroke-width="{width}" '
        f'stroke-linejoin="round" stroke-linecap="round"{dash_attr} '
        f'points="{" ".join(chunk)}"/>'
        for chunk in chunks
    )


def legend(items) -> str:
    chips = "".join(
        f'<span class="chip"><i style="background:{color}"></i>{esc(name)}</span>'
        for name, color in items
    )
    return f'<div class="legend">{chips}</div>'


# ── панели ──────────────────────────────────────────────────────────────────
def panel(pid: str, letter: str, title: str, unit: str, body: str,
          legend_items=None, note: str | None = None, payload=None) -> str:
    head = (
        f'<div class="p-head"><h3><b>{letter}.</b> {esc(title)}</h3>'
        f'<span class="unit">{esc(unit)}</span></div>'
    )
    leg = legend(legend_items) if legend_items else ""
    tip = f'<div class="tip" id="tip-{pid}" hidden></div>' if payload else ""
    data = (
        f'<script type="application/json" id="dat-{pid}">{json.dumps(payload, ensure_ascii=False)}</script>'
        if payload
        else ""
    )
    foot = f'<p class="p-note">{note}</p>' if note else ""
    cls = "plot hoverable" if payload else "plot"
    return (
        f'<figure class="panel" id="{pid}">{head}{leg}'
        f'<div class="plot-wrap">'
        f'<svg class="{cls}" viewBox="0 0 {VB_W} {VB_H}" preserveAspectRatio="none" '
        f'role="img" aria-label="{esc(title)}">{body}</svg>{tip}</div>'
        f"{foot}{data}</figure>"
    )


def daily_line_panel(pid, letter, title, unit, dates, series, bands, dec=0,
                     secondary=None, note=None):
    count = len(dates)
    flat = [v for s in series for v in s["values"] if v is not None]
    scale = Scale(min(flat), max(flat))
    stack = contextlib.ExitStack()
    stack.enter_context(gutter(bool(secondary)))
    body = [shade_bands(bands, count), frame(scale, dec), date_ticks(dates)]

    if secondary:
        s2flat = [v for v in secondary["values"] if v is not None]
        scale2 = Scale(min(s2flat), max(s2flat))
        for tick in nice_ticks(scale2.lo, scale2.hi):
            y = scale2.y(tick)
            if PY0 - 1 <= y <= PY1 + 1:
                body.append(
                    f'<text class="tick tick-l" x="{PX0 - 8}" y="{y + 3.5:.1f}">'
                    f"{fmt(tick, secondary.get('dec', 0))}</text>"
                )
        body.append(
            polyline(secondary["values"], scale2, count, secondary["color"],
                     secondary.get("width", 1.4), secondary.get("dash"))
        )

    # Рисуем с конца: первая серия в списке — самая важная — ложится поверх.
    for s in reversed(series):
        body.append(
            polyline(s["values"], scale, count, s["color"], s.get("width", 1.3), s.get("dash"))
        )

    body.append(
        f'<g class="cross" hidden><line y1="{PY0}" y2="{PY1}"/></g>'
        f'<rect class="hit" x="{PX0}" y="{PY0}" width="{PX1 - PX0}" height="{PY1 - PY0}"/>'
    )

    suffix = " · правая ось" if secondary else ""
    items = [(s["name"] + suffix, s["color"]) for s in series]
    payload_series = [
        {"name": s["name"], "color": s["color"], "values": s["values"], "dec": dec}
        for s in series
    ]
    if secondary:
        items.append((secondary["name"] + " · левая ось", secondary["color"]))
        payload_series.append(
            {
                "name": secondary["name"],
                "color": secondary["color"],
                "values": secondary["values"],
                "dec": secondary.get("dec", 0),
            }
        )
    payload = {
        "x": [f"{d[8:10]} {MONTH_RU[d[5:7]]} {d[:4]}" for d in dates],
        "series": payload_series,
        "x0": PX0,
        "x1": PX1,
        "vw": VB_W,
    }
    stack.close()
    return panel(pid, letter, title, unit, "".join(body), items, note, payload)


def bar_panel(pid, letter, title, unit, cats, series, stacked=False, line=None,
              dec=0, note=None):
    count = len(cats)
    stack = contextlib.ExitStack()
    stack.enter_context(gutter(bool(line)))
    slot = (PX1 - PX0) / count
    if stacked:
        totals = [sum(s["values"][i] for s in series) for i in range(count)]
        scale = Scale(0, max(totals))
        scale.lo = 0
    else:
        flat = [v for s in series for v in s["values"]]
        scale = Scale(0, max(flat))
        scale.lo = 0
    body = [frame(scale, dec)]

    if stacked:
        width = slot * 0.6
        for i in range(count):
            base = 0.0
            for s in series:
                value = s["values"][i]
                y_top, y_bottom = scale.y(base + value), scale.y(base)
                x = PX0 + slot * (i + 0.5) - width / 2
                body.append(
                    f'<rect class="bar" x="{x:.1f}" y="{y_top:.1f}" width="{width:.1f}" '
                    f'height="{max(y_bottom - y_top - 2, 1):.1f}" fill="{s["color"]}" rx="2">'
                    f'<title>{esc(cats[i])} · {esc(s["name"])}: {fmt(value, dec)}</title></rect>'
                )
                base += value
    else:
        group = slot * 0.72
        width = group / len(series)
        for i in range(count):
            for k, s in enumerate(series):
                value = s["values"][i]
                x = PX0 + slot * (i + 0.5) - group / 2 + k * width
                y = scale.y(value)
                body.append(
                    f'<rect class="bar" x="{x:.1f}" y="{y:.1f}" width="{max(width - 2, 1):.1f}" '
                    f'height="{max(scale.y(0) - y, 1):.1f}" fill="{s["color"]}" rx="2">'
                    f'<title>{esc(cats[i])} · {esc(s["name"])}: {fmt(value, dec)}</title></rect>'
                )

    items = [(s["name"] + (" · правая ось" if line else ""), s["color"]) for s in series]
    if line:
        lo, hi = min(line["values"]), max(line["values"])
        scale2 = Scale(lo, hi)
        for tick in nice_ticks(scale2.lo, scale2.hi):
            y = scale2.y(tick)
            if PY0 - 1 <= y <= PY1 + 1:
                body.append(
                    f'<text class="tick tick-l" x="{PX0 - 8}" y="{y + 3.5:.1f}">{fmt(tick)}</text>'
                )
        pts = " ".join(
            f"{PX0 + slot * (i + 0.5):.1f},{scale2.y(v):.1f}"
            for i, v in enumerate(line["values"])
        )
        body.append(
            f'<polyline fill="none" stroke="{line["color"]}" stroke-width="2" '
            f'stroke-linejoin="round" points="{pts}"/>'
        )
        for i, v in enumerate(line["values"]):
            body.append(
                f'<circle cx="{PX0 + slot * (i + 0.5):.1f}" cy="{scale2.y(v):.1f}" r="3.4" '
                f'fill="{line["color"]}" stroke="{CREAM}" stroke-width="1.6">'
                f'<title>{esc(cats[i])} · {esc(line["name"])}: {fmt(v)}</title></circle>'
            )
        items.append((line["name"] + " · левая ось", line["color"]))

    for i, cat in enumerate(cats):
        body.append(
            f'<text class="tick tick-x" x="{PX0 + slot * (i + 0.5):.1f}" y="{PY1 + 16}">'
            f"{esc(cat)}</text>"
        )
    body.append(f'<line class="axis" x1="{PX0}" x2="{PX1}" y1="{PY1}" y2="{PY1}"/>')
    stack.close()
    return panel(pid, letter, title, unit, "".join(body), items, note)


def scatter_panel(pid, letter, title, unit, xs, ys, fit, note=None):
    xlo, xhi = min(xs), max(xs)
    ylo, yhi = min(ys), max(ys)
    xspan = (xhi - xlo) * 1.05 or 1

    def px(value):
        return PX0 + (value - xlo + (xhi - xlo) * 0.025) / xspan * (PX1 - PX0)

    scale = Scale(ylo, yhi)
    body = [frame(scale, 0)]
    for x, y in zip(xs, ys):
        body.append(
            f'<circle class="dot" cx="{px(x):.1f}" cy="{scale.y(y):.1f}" r="2.6" '
            f'fill="{SECONDARY}"><title>расходы {fmt(x)} → выручка {fmt(y)}</title></circle>'
        )
    band = (
        " ".join(f"{px(x):.1f},{scale.y(u):.1f}" for x, u in zip(fit["x"], fit["upper"]))
        + " "
        + " ".join(
            f"{px(x):.1f},{scale.y(l):.1f}"
            for x, l in reversed(list(zip(fit["x"], fit["lower"])))
        )
    )
    body.insert(1, f'<polygon class="ci" points="{band}"/>')
    body.append(
        '<polyline fill="none" stroke="%s" stroke-width="2.2" points="%s"/>'
        % (
            PRIMARY,
            " ".join(f"{px(x):.1f},{scale.y(v):.1f}" for x, v in zip(fit["x"], fit["fit"])),
        )
    )
    for tick in nice_ticks(xlo, xhi, 6):
        if xlo <= tick <= xhi:
            body.append(
                f'<text class="tick tick-x" x="{px(tick):.1f}" y="{PY1 + 16}">{fmt(tick)}</text>'
            )
    body.append(f'<line class="axis" x1="{PX0}" x2="{PX1}" y1="{PY1}" y2="{PY1}"/>')
    items = [("Дни (N = 365)", SECONDARY), ("Регрессия", PRIMARY), ("95 % ДИ", SHADE)]
    return panel(pid, letter, title, unit, "".join(body), items, note)


def dot_error_panel(pid, letter, title, unit, cats, values, errs, note=None):
    count = len(cats)
    slot = (PX1 - PX0) / count
    lo = min(v - e for v, e in zip(values, errs))
    hi = max(v + e for v, e in zip(values, errs))
    scale = Scale(min(lo, 0), hi)
    body = [frame(scale, 1)]
    zero = scale.y(0)
    body.append(f'<line class="zero" x1="{PX0}" x2="{PX1}" y1="{zero:.1f}" y2="{zero:.1f}"/>')
    for i, (cat, value, err) in enumerate(zip(cats, values, errs)):
        x = PX0 + slot * (i + 0.5)
        color = PRIMARY if value >= 0 else ACCENT2
        top, bottom = scale.y(value + err), scale.y(value - err)
        body.append(
            f'<line class="err" x1="{x:.1f}" x2="{x:.1f}" y1="{top:.1f}" y2="{bottom:.1f}"/>'
            f'<line class="err" x1="{x - 5:.1f}" x2="{x + 5:.1f}" y1="{top:.1f}" y2="{top:.1f}"/>'
            f'<line class="err" x1="{x - 5:.1f}" x2="{x + 5:.1f}" y1="{bottom:.1f}" y2="{bottom:.1f}"/>'
            f'<circle cx="{x:.1f}" cy="{scale.y(value):.1f}" r="5.5" fill="{color}" '
            f'stroke="{CREAM}" stroke-width="1.8">'
            f'<title>{esc(cat)}: {value:+.1f} % ± {err:.1f}</title></circle>'
            f'<text class="val" x="{x:.1f}" y="{top - 8:.1f}">{value:+.1f}</text>'
        )
        words = cat.split(" ")
        for k, word in enumerate(words):
            body.append(
                f'<text class="tick tick-x" x="{x:.1f}" y="{PY1 + 15 + k * 11}">{esc(word)}</text>'
            )
    body.append(f'<line class="axis" x1="{PX0}" x2="{PX1}" y1="{PY1}" y2="{PY1}"/>')
    items = [("Рост прогнозируется", PRIMARY), ("Снижение", ACCENT2)]
    return panel(pid, letter, title, unit, "".join(body), items, note)


def heatmap_panel(pid, letter, title, unit, months, columns, note=None):
    head = "".join(f"<th>{esc(label)}</th>" for label, _, _ in columns)
    rows = []
    for i, month in enumerate(months):
        cells = []
        for label, values, reverse in columns:
            ordered = sorted(values)
            rank = ordered.index(values[i]) / max(len(ordered) - 1, 1)
            if reverse:
                rank = 1 - rank
            if rank < 0.5:
                color = mix(RAMP[0], RAMP[1], rank * 2)
            else:
                color = mix(RAMP[1], RAMP[2], (rank - 0.5) * 2)
            ink = "#FFFFFF" if rank > 0.62 else INK
            dec = 2 if "%" in label and "Конверсия" in label else (1 if "Доля" in label else 0)
            cells.append(
                f'<td style="background:{color};color:{ink}">{fmt(values[i], dec)}</td>'
            )
        rows.append(f"<tr><th>{esc(month)}</th>{''.join(cells)}</tr>")
    table = (
        f'<div class="hm-wrap"><table class="hm">'
        f"<thead><tr><th>Месяц</th>{head}</tr></thead>"
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )
    ramp = (
        f'<div class="legend"><span class="chip"><i class="ramp"></i>'
        f"перцентиль 5 → 95 внутри столбца</span></div>"
    )
    foot = f'<p class="p-note">{note}</p>' if note else ""
    return (
        f'<figure class="panel" id="{pid}">'
        f'<div class="p-head"><h3><b>{letter}.</b> {esc(title)}</h3>'
        f'<span class="unit">{esc(unit)}</span></div>{ramp}{table}{foot}</figure>'
    )


def mix(c1: str, c2: str, t: float) -> str:
    t = max(0.0, min(1.0, t))
    a = [int(c1[i : i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i : i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


# ── сборка страницы ─────────────────────────────────────────────────────────
def build() -> str:
    daily = read_csv("company_daily_2025_2026.csv")
    monthly = read_csv("company_monthly_summary.csv")
    dates = [r["Date"] for r in daily]
    months = [month_label(m["Month"]) for m in monthly]

    def col(key, cast=float):
        return [cast(r[key]) for r in daily]

    def mcol(key, cast=float):
        return [cast(m[key]) for m in monthly]

    def indexed(key):
        values = col(key)
        base = sum(values[:28]) / 28
        return [round(v / base * 100, 1) for v in values]

    sheets = []

    # ── Лист 1 ──
    p1a = daily_line_panel(
        "p1a", "A", "Спрос, трафик и воронка сделок, 365 дней", "индекс, 100 = май 2025",
        dates,
        [
            {"name": "Спрос", "color": PRIMARY, "values": indexed("DemandIndex"), "width": 1.5},
            {"name": "Трафик", "color": SECONDARY, "values": indexed("Traffic")},
            {"name": "Лиды", "color": TERTIARY, "values": indexed("Leads")},
            {"name": "Сделки", "color": ACCENT, "values": indexed("Deals")},
        ],
        CAMPAIGNS,
        note="Розовым залиты периоды кампаний: <sup>а</sup> июнь 2025, <sup>б</sup> январь 2026.",
    )
    p1b = bar_panel(
        "p1b", "B", "Помесячная структура спроса по каналам", "визиты",
        months,
        [
            {"name": "Органика", "color": PRIMARY, "values": mcol("Organic")},
            {"name": "Платный", "color": SECONDARY, "values": mcol("Paid")},
            {"name": "Прямой", "color": TERTIARY, "values": mcol("Direct")},
        ],
        stacked=True,
    )

    xs, ys = col("MarketingSpend"), col("Revenue")
    n = len(xs)
    xbar = sum(xs) / n
    ybar = sum(ys) / n
    sxx = sum((x - xbar) ** 2 for x in xs)
    slope = sum((x - xbar) * (y - ybar) for x, y in zip(xs, ys)) / sxx
    intercept = ybar - slope * xbar
    resid = [y - (slope * x + intercept) for x, y in zip(xs, ys)]
    s_err = (sum(r * r for r in resid) / (n - 2)) ** 0.5
    order = sorted(range(n), key=lambda i: xs[i])
    fx = [xs[i] for i in order]
    ff = [slope * x + intercept for x in fx]
    se = [s_err * (1 / n + (x - xbar) ** 2 / sxx) ** 0.5 for x in fx]
    p1c = scatter_panel(
        "p1c", "C", "Маркетинговые расходы против выручки, N = 365", "₽ тыс. · ось X — расходы",
        xs, ys,
        {"x": fx, "fit": ff,
         "upper": [f + 1.96 * s for f, s in zip(ff, se)],
         "lower": [f - 1.96 * s for f, s in zip(ff, se)]},
        note=f"Наклон {slope:+.2f} ₽ выручки на ₽ расходов; полоса — 95 % доверительный интервал среднего.",
    )
    sheets.append(("Мониторинг рынка", 1,
                   "Спрос, трафик и воронка сделок",
                   [p1a, p1b, p1c],
                   ["Затенённые области отмечают периоды маркетинговых кампаний.",
                    "<sup>а</sup> Кампания 1 (июн. 2025).  <sup>б</sup> Кампания 2 (янв. 2026).",
                    "Источники: внутренняя аналитика (демо-данные), 2025-05 … 2026-05."]))

    # ── Лист 2 ──
    demand = col("DemandIndex")
    p2a = daily_line_panel(
        "p2a", "A", "Индекс спроса и сглаженный тренд, 365 дней", "индекс",
        dates,
        [
            {"name": "Спрос, день", "color": TERTIARY, "values": demand, "width": 1.0},
            {"name": "Скользящее среднее, 28 дней", "color": PRIMARY,
             "values": moving_avg(demand, 28), "width": 2.4},
        ],
        CAMPAIGNS + SHOCKS,
        dec=0,
        note="Затенены и кампании (<sup>а</sup> <sup>б</sup>), и рыночные шоки (<sup>в</sup> <sup>г</sup>).",
    )
    categories = [
        ("Корпоративный сегмент", 8.4, 2.1), ("Малый бизнес", 5.6, 1.6),
        ("Госсектор", 2.1, 2.8), ("Розница", 6.9, 1.9),
        ("Партнёрский канал", 11.2, 3.4), ("Экспорт", -1.8, 2.6),
        ("Подписка", 14.5, 2.2), ("Услуги внедрения", 3.7, 1.5),
    ]
    p2b = dot_error_panel(
        "p2b", "B", "Прогноз роста по категориям, ±95 % ДИ", "% год к году",
        [c[0] for c in categories], [c[1] for c in categories], [c[2] for c in categories],
        note="Единственная панель, где планки погрешностей обязательны по ТЗ: ScatterChart + ErrorBars.",
    )
    p2c = bar_panel(
        "p2c", "C", "Сезонность спроса и средний чек по месяцам", "индекс (справа) · ₽ (слева)",
        months,
        [{"name": "Сезонность", "color": SECONDARY, "values": mcol("Seasonality")}],
        line={"name": "Средний чек", "color": ACCENT, "values": mcol("AvgCheck")},
        note="Две шкалы: сезонность на правой оси, средний чек на левой — так же, как в файле.",
    )
    sheets.append(("Тренды", 2,
                   "Сглаженная динамика, прогноз категорий, сезонность",
                   [p2a, p2b, p2c],
                   ["Затенённые области отмечают периоды кампаний или рыночных шоков.",
                    "<sup>а</sup> Кампания 1 (июн. 2025).  <sup>б</sup> Кампания 2 (янв. 2026).",
                    "<sup>в</sup> Рыночный шок 1 (сен. 2025).  <sup>г</sup> Рыночный шок 2 (мар. 2026).",
                    "Источники: внутренняя аналитика (демо-данные), 2025-05 … 2026-05."]))

    # ── Лист 3 ──
    p3a = daily_line_panel(
        "p3a", "A", "Доля голоса: мы против трёх конкурентов, 365 дней", "% упоминаний",
        dates,
        [
            {"name": "Мы", "color": PRIMARY, "values": col("SOV_Us"), "width": 2.2},
            {"name": "Конкурент A", "color": ACCENT2, "values": col("SOV_CompA")},
            {"name": "Конкурент B", "color": SECONDARY, "values": col("SOV_CompB")},
            {"name": "Конкурент C", "color": TERTIARY, "values": col("SOV_CompC")},
        ],
        SHOCKS,
        dec=1,
        note="Наша линия намеренно вдвое толще остальных — она читается первой.",
    )
    diff = [round(a - b, 2) for a, b in zip(col("PriceIndexUs"), col("PriceIndexMarket"))]
    p3b = daily_line_panel(
        "p3b", "B", "Ценовой индекс: мы против средней по рынку", "индекс (справа) · п.п. (слева)",
        dates,
        [
            {"name": "Цена: мы", "color": PRIMARY, "values": col("PriceIndexUs"), "width": 1.8},
            {"name": "Цена: рынок", "color": ACCENT2, "values": col("PriceIndexMarket"),
             "width": 1.4, "dash": "5 4"},
        ],
        [],
        dec=1,
        secondary={"name": "Разница, п.п.", "color": ACCENT, "values": diff, "dec": 1, "width": 1.2},
        note="Разница вынесена на левую ось: её размах в 30 раз меньше самого индекса.",
    )
    p3c = bar_panel(
        "p3c", "C", "Доля рынка по игрокам и упоминания бренда", "% (справа) · шт. (слева)",
        months,
        [
            {"name": "Мы", "color": PRIMARY, "values": mcol("MarketShareUs")},
            {"name": "A", "color": ACCENT2, "values": mcol("MarketShareA")},
            {"name": "B", "color": SECONDARY, "values": mcol("MarketShareB")},
            {"name": "C", "color": TERTIARY, "values": mcol("MarketShareC")},
        ],
        line={"name": "Упоминания: мы", "color": ACCENT, "values": mcol("MentionsUs")},
        dec=1,
    )
    sheets.append(("Конкуренты", 3,
                   "Доля голоса, ценовая позиция, доля рынка",
                   [p3a, p3b, p3c],
                   ["Затенённые области отмечают периоды рыночных шоков.",
                    "<sup>в</sup> Рыночный шок 1 (сен. 2025).  <sup>г</sup> Рыночный шок 2 (мар. 2026).",
                    "Источники: внутренняя аналитика и открытые данные по конкурентам (демо-данные)."]))

    # ── Лист 4 ──
    p4a = daily_line_panel(
        "p4a", "A", "Визиты сайта и видимость в поиске, 365 дней", "визиты (справа) · % (слева)",
        dates,
        [{"name": "Визиты", "color": SECONDARY, "values": col("Traffic", lambda v: int(float(v))), "width": 1.1}],
        [],
        secondary={"name": "Видимость в поиске", "color": PRIMARY,
                   "values": col("SearchVisibility"), "dec": 0, "width": 2.2},
        note="Видимость растёт с 44 % до 53 % — трафик следует за ней с лагом.",
    )
    conversion = col("ConversionRate")
    p4b = daily_line_panel(
        "p4b", "B", "Конверсия воронки и стоимость лида", "% (справа) · ₽ (слева)",
        dates,
        [
            {"name": "Конверсия, день", "color": TERTIARY, "values": conversion, "width": 1.0},
            {"name": "Конверсия, среднее за 28 дней", "color": PRIMARY,
             "values": moving_avg(conversion, 28), "width": 2.4},
        ],
        [],
        dec=2,
        secondary={"name": "CPL", "color": ACCENT2, "values": col("CPL", lambda v: int(float(v))),
                   "dec": 0, "width": 1.1},
    )
    p4c = heatmap_panel(
        "p4c", "C", "Помесячная карта KPI, перцентильная раскраска", "условное форматирование",
        months,
        [
            ("Выручка, ₽ тыс.", mcol("KPI_Revenue"), False),
            ("Трафик, визиты", mcol("KPI_Traffic"), False),
            ("Конверсия, %", mcol("KPI_Conversion"), False),
            ("CPL, ₽", mcol("KPI_CPL"), True),
            ("Доля голоса, %", mcol("KPI_Share"), False),
            ("NPS", mcol("KPI_NPS"), False),
        ],
        note="Единственная панель, которая в Excel не объект-график, а условное форматирование "
             "(ColorScaleRule). Шкала последовательная, одноцветная: светлое → тёмное. "
             "Для CPL инвертирована, меньше — лучше.",
    )
    sheets.append(("Цифровой след", 4,
                   "Трафик и видимость, экономика лида, карта KPI",
                   [p4a, p4b, p4c],
                   ["Панель C — условное форматирование, а не объект-график.",
                    "Панели A, B: правая ось — первая метрика, левая — вторая.",
                    "Источники: внутренняя аналитика (демо-данные), 2025-05 … 2026-05."]))

    nav = "".join(
        f'<a href="#sheet{no}">{no}. {esc(name)}</a>' for name, no, _, _, _ in sheets
    )
    blocks = []
    for name, no, lead, panels, notes in sheets:
        foot = "".join(f"<p>{note}</p>" for note in notes)
        blocks.append(
            f'<section class="sheet" id="sheet{no}">'
            f'<header class="s-head"><div><h2>{esc(name)}</h2>'
            f'<p class="lead">{esc(lead)}</p></div>'
            f'<span class="gno">Graph {no}</span></header>'
            f'<div class="panels">{"".join(panels)}</div>'
            f'<footer class="s-foot">{foot}</footer></section>'
        )

    return TEMPLATE.format(nav=nav, body="".join(blocks))


TEMPLATE = """<title>Годовой обзор деятельности 2025–2026</title>
<style>
  :root {{
    --paper: #F5EDE6; --edge: #E4D7CA; --ink: #3A322C; --muted: #8B7B6B;
    --rule: #C9BCAE; --grid: #E3D9CE; --axis: #C0B3A6;
    --p1: #1F3A5F; --p2: #4A7BA7; --p3: #8FB8D9; --p4: #3F6B52; --p5: #C97B5A;
    --shade: #EDCBC1;
    --serif: Georgia, 'Times New Roman', serif;
    --sans: ui-sans-serif, -apple-system, 'Segoe UI', Roboto, sans-serif;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; background: var(--edge); color: var(--ink);
    font-family: var(--serif); font-size: 16px; line-height: 1.5;
    -webkit-font-smoothing: antialiased;
  }}
  .bar {{
    position: sticky; top: 0; z-index: 20; background: rgba(245,237,230,.94);
    backdrop-filter: blur(8px); border-bottom: 1px solid var(--rule);
    display: flex; flex-wrap: wrap; gap: 4px 18px; align-items: baseline;
    padding: 10px clamp(16px, 4vw, 44px); font-family: var(--sans); font-size: 13px;
  }}
  .bar strong {{ font-family: var(--serif); font-size: 15px; font-weight: 700; margin-right: 8px; }}
  .bar a {{ color: var(--muted); text-decoration: none; border-bottom: 1px solid transparent; }}
  .bar a:hover, .bar a:focus-visible {{ color: var(--p1); border-bottom-color: var(--p1); }}
  .bar .hint {{ margin-left: auto; color: var(--muted); font-style: italic; }}
  main {{ max-width: 1240px; margin: 0 auto; padding: clamp(18px, 3vw, 34px) clamp(12px, 3vw, 28px) 72px; }}
  .masthead {{ padding: 28px 0 20px; border-bottom: 2px solid var(--ink); margin-bottom: 26px; }}
  .masthead h1 {{ font-size: clamp(26px, 4vw, 40px); margin: 0 0 6px; letter-spacing: -.01em; text-wrap: balance; }}
  .masthead p {{ margin: 0; color: var(--muted); font-style: italic; max-width: 62ch; }}
  .sheet {{ scroll-margin-top: 60px; background: var(--paper); border: 1px solid var(--rule); padding: clamp(16px, 2.4vw, 30px); margin-bottom: 26px; }}
  .s-head {{ display: flex; justify-content: space-between; align-items: baseline; gap: 20px;
    border-bottom: 1px solid var(--rule); padding-bottom: 10px; margin-bottom: 6px; }}
  .s-head h2 {{ margin: 0; font-size: clamp(19px, 2.4vw, 25px); }}
  .lead {{ margin: 2px 0 0; color: var(--muted); font-style: italic; font-size: 14px; }}
  .gno {{ font-style: italic; font-weight: 700; color: var(--muted); white-space: nowrap; font-size: 15px; }}
  .panels {{ display: flex; flex-direction: column; gap: 26px; padding-top: 16px; }}
  .panel {{ margin: 0; }}
  .p-head {{ display: flex; justify-content: space-between; align-items: baseline; gap: 16px; }}
  .p-head h3 {{ margin: 0; font-size: 15px; font-weight: 400; }}
  .p-head b {{ font-weight: 700; }}
  .unit {{ font-size: 12px; font-style: italic; color: var(--muted); text-align: right; white-space: nowrap; }}
  .legend {{ display: flex; flex-wrap: wrap; gap: 4px 16px; margin: 7px 0 4px; }}
  .chip {{ display: inline-flex; align-items: center; gap: 6px; font-size: 12px; color: var(--muted); }}
  .chip i {{ width: 14px; height: 3px; border-radius: 2px; display: inline-block; }}
  .chip i.ramp {{ width: 60px; height: 9px; border-radius: 2px;
    background: linear-gradient(90deg, #E8EEF4, #7FA3C4, #1F3A5F); }}
  .plot-wrap {{ position: relative; }}
  svg.plot {{ width: 100%; height: clamp(190px, 24vw, 290px); display: block; overflow: visible; }}
  .grid {{ stroke: var(--grid); stroke-width: 1; }}
  .axis {{ stroke: var(--axis); stroke-width: 1; }}
  .zero {{ stroke: var(--ink); stroke-width: 1; stroke-dasharray: 3 3; opacity: .5; }}
  .shade {{ fill: var(--shade); opacity: .55; }}
  .ci {{ fill: var(--shade); opacity: .5; }}
  .err {{ stroke: #8B7B6B; stroke-width: 1.4; }}
  .dot {{ opacity: .55; }}
  text {{ font-family: var(--serif); fill: var(--muted); }}
  .tick {{ font-size: 11px; font-style: italic; text-anchor: start; }}
  .tick-l {{ text-anchor: end; }}
  .tick-x {{ text-anchor: middle; font-size: 11px; }}
  .val {{ font-size: 11px; text-anchor: middle; fill: var(--ink); font-style: italic; }}
  .hit {{ fill: transparent; }}
  .cross[hidden] {{ display: none; }}
  .cross line {{ stroke: var(--ink); stroke-width: 1; stroke-dasharray: 3 3; opacity: .55; }}
  .tip {{ position: absolute; pointer-events: none; background: var(--paper);
    border: 1px solid var(--rule); padding: 7px 10px; font-size: 12px; line-height: 1.45;
    box-shadow: 0 2px 10px rgba(58,50,44,.14); min-width: 150px; z-index: 5; }}
  .tip .t-x {{ font-style: italic; color: var(--muted); margin-bottom: 3px; }}
  .tip .t-row {{ display: flex; justify-content: space-between; gap: 14px; }}
  .tip .t-row span {{ display: inline-flex; align-items: center; gap: 6px; }}
  .tip i {{ width: 10px; height: 3px; border-radius: 2px; }}
  .tip b {{ font-variant-numeric: tabular-nums; font-weight: 700; }}
  .p-note {{ margin: 8px 0 0; font-size: 12.5px; font-style: italic; color: var(--muted); max-width: 88ch; }}
  .p-note sup {{ font-style: normal; }}
  .hm-wrap {{ overflow-x: auto; margin-top: 6px; }}
  table.hm {{ border-collapse: separate; border-spacing: 2px; width: 100%; font-size: 12.5px;
    font-variant-numeric: tabular-nums; }}
  table.hm th {{ font-weight: 700; font-style: italic; color: var(--muted); padding: 5px 8px;
    text-align: right; white-space: nowrap; font-size: 12px; }}
  table.hm thead th {{ border-bottom: 1px solid var(--rule); }}
  table.hm tbody th {{ text-align: left; }}
  table.hm td {{ padding: 6px 10px; text-align: right; white-space: nowrap; }}
  .s-foot {{ border-top: 1px solid var(--rule); margin-top: 22px; padding-top: 10px; }}
  .s-foot p {{ margin: 0 0 2px; font-size: 12.5px; font-style: italic; color: var(--muted); }}
  .s-foot sup {{ font-style: normal; }}
  .colophon {{ font-size: 13px; color: var(--muted); font-style: italic; max-width: 74ch; }}
  .colophon code {{ font-family: var(--sans); font-style: normal; font-size: 12px;
    background: var(--paper); border: 1px solid var(--rule); padding: 1px 5px; }}
  @media (prefers-reduced-motion: reduce) {{ * {{ transition: none !important; }} }}
</style>

<nav class="bar">
  <strong>Годовой обзор 2025–2026</strong>
  {nav}
  <span class="hint">наведите курсор на график — появятся значения за день</span>
</nav>

<main>
  <div class="masthead">
    <h1>Годовой обзор деятельности компании</h1>
    <p>Веб-превью книги <code>Company_Annual_BIS_Style.xlsx</code>: те же 12 панелей,
    те же данные, та же палитра. Структура BIS, визуал LVMH. Данные синтетические,
    май 2025 — апрель 2026.</p>
  </div>
  {body}
  <p class="colophon">Превью собирается из тех же CSV, что и книга Excel:
  <code>python3 scripts/build_preview_page.py</code>. В самом файле все графики —
  нативные объекты Excel, серии привязаны к диапазонам ячеек, поэтому правка данных
  пересчитывает график.</p>
</main>

<script>
(function () {{
  var fmt = function (v, d) {{
    return v == null ? '—' : v.toLocaleString('ru-RU', {{
      minimumFractionDigits: d, maximumFractionDigits: d
    }});
  }};
  document.querySelectorAll('svg.hoverable').forEach(function (svg) {{
    var fig = svg.closest('.panel');
    var node = document.getElementById('dat-' + fig.id);
    if (!node) return;
    var data = JSON.parse(node.textContent);
    var tip = document.getElementById('tip-' + fig.id);
    var cross = svg.querySelector('.cross');
    var line = cross.querySelector('line');
    var hit = svg.querySelector('.hit');
    var n = data.x.length;

    function move(ev) {{
      var box = svg.getBoundingClientRect();
      var vx = (ev.clientX - box.left) / box.width * data.vw;
      var t = (vx - data.x0) / (data.x1 - data.x0);
      var i = Math.round(t * (n - 1));
      if (i < 0) i = 0;
      if (i > n - 1) i = n - 1;
      var px = data.x0 + i / (n - 1) * (data.x1 - data.x0);
      line.setAttribute('x1', px);
      line.setAttribute('x2', px);
      cross.removeAttribute('hidden');
      var rows = data.series.map(function (s) {{
        return '<div class="t-row"><span><i style="background:' + s.color + '"></i>' +
          s.name + '</span><b>' + fmt(s.values[i], s.dec) + '</b></div>';
      }}).join('');
      tip.innerHTML = '<div class="t-x">' + data.x[i] + '</div>' + rows;
      tip.hidden = false;
      var left = px / data.vw * box.width + 16;
      if (left + tip.offsetWidth > box.width) left = px / data.vw * box.width - tip.offsetWidth - 16;
      tip.style.left = Math.max(0, left) + 'px';
      tip.style.top = '4px';
    }}

    function leave() {{
      cross.setAttribute('hidden', '');
      tip.hidden = true;
    }}

    hit.addEventListener('pointermove', move);
    hit.addEventListener('pointerdown', move);
    hit.addEventListener('pointerleave', leave);
  }});
}})();
</script>
"""


def main() -> None:
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(build(), encoding="utf-8")
    print(f"Сохранено: {OUT.relative_to(ROOT)}  ({OUT.stat().st_size // 1024} КБ)")


if __name__ == "__main__":
    main()
