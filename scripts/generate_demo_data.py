"""Генератор синтетических данных для демо-workbook «Годовой обзор деятельности компании».

Создаёт data/company_daily_2025_2026.csv (365 строк) и
data/company_monthly_summary.csv (12 строк) по схеме из
docs/scenario-company-annual-review.md.

Запуск:  python3 scripts/generate_demo_data.py
"""

from __future__ import annotations

import csv
import datetime as dt
import math
import pathlib

import numpy as np

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

START = dt.date(2025, 5, 1)
N_DAYS = 365

# Затенённые периоды (номера дней, включительно) — см. часть A сценария.
CAMPAIGNS = [(45, 60), (210, 228)]
SHOCKS = [(132, 139), (300, 306)]

RNG = np.random.default_rng(20260515)


def in_ranges(day: int, ranges) -> int:
    return int(any(lo <= day <= hi for lo, hi in ranges))


def build_daily() -> list[dict]:
    rows = []
    for i in range(N_DAYS):
        day = i + 1
        date = START + dt.timedelta(days=i)
        dow = date.weekday()

        campaign = in_ranges(day, CAMPAIGNS)
        shock = in_ranges(day, SHOCKS)

        # Сезонность: годовая волна + недельный провал на выходных.
        season = 1.0 + 0.13 * math.sin(2 * math.pi * (i - 40) / 365)
        weekend = 0.72 if dow >= 5 else 1.0
        trend = 1.0 + 0.22 * i / N_DAYS  # общий рост за год
        boost = 1.0 + 0.28 * campaign
        hit = 1.0 - 0.22 * shock

        base = season * weekend * trend * boost * hit

        demand = 62 * base * RNG.normal(1, 0.045)
        traffic = 4800 * base * RNG.normal(1, 0.055)
        leads = traffic * (0.029 + 0.004 * campaign) * RNG.normal(1, 0.09)
        deals = leads * (0.145 - 0.02 * shock) * RNG.normal(1, 0.13)

        organic = traffic * (0.53 - 0.10 * campaign) * RNG.normal(1, 0.04)
        paid = traffic * (0.31 + 0.11 * campaign) * RNG.normal(1, 0.07)
        direct = max(traffic - organic - paid, traffic * 0.08)

        spend = (52 + 46 * campaign) * base * RNG.normal(1, 0.11)
        revenue = deals * (19_800 * (1 - 0.03 * shock) * RNG.normal(1, 0.06)) / 1000
        avg_check = revenue * 1000 / max(deals, 1)

        # Доли голоса: сумма по четырём игрокам ~ 81%, остальное — длинный хвост.
        sov_us = 18.0 + 2.6 * campaign + 0.0035 * i + RNG.normal(0, 0.45)
        sov_a = 27.6 - 0.0028 * i + RNG.normal(0, 0.5)
        sov_b = 21.0 + 0.0009 * i + RNG.normal(0, 0.42)
        sov_c = 14.2 - 0.0011 * i + RNG.normal(0, 0.38)

        price_market = 100 + 3.4 * i / N_DAYS + 1.9 * shock + RNG.normal(0, 0.28)
        price_us = price_market + 1.3 - 2.4 * campaign + RNG.normal(0, 0.35)

        visibility = 44 + 8.5 * i / N_DAYS + 3.1 * campaign + RNG.normal(0, 0.9)
        conversion = leads / max(traffic, 1) * 100
        cpl = spend * 1000 / max(leads, 1)

        rows.append(
            {
                "Date": date.isoformat(),
                "DayIndex": day,
                "DemandIndex": round(demand, 1),
                "Traffic": int(round(traffic)),
                "Leads": int(round(leads)),
                "Deals": int(round(deals)),
                "ChannelOrganic": int(round(organic)),
                "ChannelPaid": int(round(paid)),
                "ChannelDirect": int(round(direct)),
                "MarketingSpend": round(spend, 1),
                "Revenue": round(revenue, 1),
                "AvgCheck": int(round(avg_check)),
                "SOV_Us": round(sov_us, 2),
                "SOV_CompA": round(sov_a, 2),
                "SOV_CompB": round(sov_b, 2),
                "SOV_CompC": round(sov_c, 2),
                "PriceIndexUs": round(price_us, 2),
                "PriceIndexMarket": round(price_market, 2),
                "SearchVisibility": round(visibility, 1),
                "ConversionRate": round(conversion, 2),
                "CPL": int(round(cpl)),
                "IsCampaign": campaign,
                "IsMarketShock": shock,
            }
        )
    return rows


def build_monthly(daily: list[dict]) -> list[dict]:
    buckets: dict[str, list[dict]] = {}
    for row in daily:
        month = row["Date"][:7]
        buckets.setdefault(month, []).append(row)

    months = sorted(buckets)[:12]
    out = []
    for idx, month in enumerate(months):
        rows = buckets[month]

        def total(key):
            return sum(r[key] for r in rows)

        def mean(key):
            return sum(r[key] for r in rows) / len(rows)

        revenue = total("Revenue")
        traffic = total("Traffic")
        growth = 3.4 + 2.6 * math.sin(2 * math.pi * (idx - 1) / 12) + RNG.normal(0, 0.6)

        out.append(
            {
                "Month": month,
                "Organic": total("ChannelOrganic"),
                "Paid": total("ChannelPaid"),
                "Direct": total("ChannelDirect"),
                "CategoryGrowth": round(growth, 1),
                "GrowthCI": round(1.4 + 0.9 * abs(RNG.normal(0, 1)), 1),
                "Seasonality": round(mean("DemandIndex") / 0.62, 1),
                "AvgCheck": int(round(mean("AvgCheck"))),
                "MarketShareUs": round(mean("SOV_Us"), 1),
                "MarketShareA": round(mean("SOV_CompA"), 1),
                "MarketShareB": round(mean("SOV_CompB"), 1),
                "MarketShareC": round(mean("SOV_CompC"), 1),
                "MentionsUs": int(round(mean("SOV_Us") * 19 + RNG.normal(0, 12))),
                "MentionsA": int(round(mean("SOV_CompA") * 22 + RNG.normal(0, 15))),
                "MentionsB": int(round(mean("SOV_CompB") * 21 + RNG.normal(0, 14))),
                "MentionsC": int(round(mean("SOV_CompC") * 20 + RNG.normal(0, 11))),
                "KPI_Revenue": int(round(revenue)),
                "KPI_Traffic": traffic,
                "KPI_Conversion": round(mean("ConversionRate"), 2),
                "KPI_CPL": int(round(mean("CPL"))),
                "KPI_Share": round(mean("SOV_Us"), 1),
                "KPI_NPS": int(round(40 + 6 * idx / 11 + RNG.normal(0, 1.6))),
            }
        )
    return out


def write_csv(path: pathlib.Path, rows: list[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"{path.relative_to(ROOT)}: {len(rows)} строк")


def main() -> None:
    DATA.mkdir(exist_ok=True)
    daily = build_daily()
    monthly = build_monthly(daily)
    write_csv(DATA / "company_daily_2025_2026.csv", daily)
    write_csv(DATA / "company_monthly_summary.csv", monthly)


if __name__ == "__main__":
    main()
