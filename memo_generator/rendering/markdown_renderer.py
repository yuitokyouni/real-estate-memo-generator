"""Render memo data as Markdown."""
from memo_generator.models.property_input import PropertyInput


def render_markdown(memo_data: dict) -> str:
    prop: PropertyInput = memo_data["property"]
    m = memo_data["metrics"]
    s = memo_data["sections"]
    hold = prop.hold_period_years

    irr_str = f"{m['irr']:.2%}" if m["irr"] is not None else "N/A"

    lines = [
        f"# INVESTMENT MEMORANDUM — CONFIDENTIAL",
        f"## {prop.property_name} | {prop.city}, {prop.state_or_country}",
        "",
        "---",
        "",
        "## EXECUTIVE SUMMARY",
        "",
        s.get("executive_summary", ""),
        "",
        "### Key Metrics",
        "",
        "| Metric | Value |",
        "|---|---|",
        f"| Purchase Price | ${prop.purchase_price:,.0f} |",
    ]

    if m.get("price_per_unit"):
        lines.append(f"| Price per Unit | ${m['price_per_unit']:,.0f} |")
    if m.get("price_per_sqft"):
        lines.append(f"| Price per SF | ${m['price_per_sqft']:.2f} |")

    lines += [
        f"| Going-in Cap Rate | {m['cap_rate']:.2%} |",
        f"| NOI (Year 1) | ${m['noi']:,.0f} |",
        f"| DSCR | {m['dscr']:.2f}x |",
        f"| LTV | {m['ltv']:.1%} |",
        f"| Cash-on-Cash (Year 1) | {m['cash_on_cash_yr1']:.2%} |",
        f"| Levered IRR ({hold}-yr) | {irr_str} |",
        f"| Equity Multiple | {m['equity_multiple']:.2f}x |",
        "",
        "---",
        "",
        "## MARKET ANALYSIS",
        "",
        s.get("market_analysis", ""),
        "",
        "---",
        "",
        "## INVESTMENT STRATEGY",
        "",
        s.get("investment_strategy", ""),
        "",
        "---",
        "",
        "## RISK FACTORS",
        "",
        s.get("risk_factors", ""),
        "",
        "---",
        "",
        "## 5-YEAR CASH FLOW PROJECTION",
        "",
        "| Year | Cash Flow |",
        "|---|---|",
        f"| 0 (Equity Invested) | $({abs(m['cash_flows'][0]):,.0f}) |",
    ]

    for i, cf in enumerate(m["cash_flows"][1:], start=1):
        lines.append(f"| Year {i} | ${cf:,.0f} |")

    lines += [
        "",
        "---",
        "",
        f"*This memorandum was prepared using automated analysis. All projections are estimates and subject to change.*",
    ]

    return "\n".join(lines)
