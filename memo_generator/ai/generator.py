"""Orchestrates Claude API calls to generate all memo sections."""
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from memo_generator.ai.client import generate_section, _load_prompt
from memo_generator.financials.calculator import calculate_all_metrics
from memo_generator.models.property_input import PropertyInput

_SYSTEM_PROMPT = _load_prompt("system_prompt.txt")


def _build_property_context(prop: PropertyInput, metrics: dict) -> str:
    """Build a structured context block to include in every section prompt."""
    lines = [
        "=== PROPERTY DATA ===",
        f"Property Name: {prop.property_name}",
        f"Type: {prop.property_type}",
        f"Location: {prop.address}, {prop.city}, {prop.state_or_country}",
        f"Year Built: {prop.year_built}",
    ]
    if prop.total_units:
        lines.append(f"Total Units: {prop.total_units}")
    if prop.total_sqft:
        lines.append(f"Total SF: {prop.total_sqft:,.0f}")

    lines += [
        "",
        "=== KEY FINANCIALS ===",
        f"Purchase Price: ${prop.purchase_price:,.0f}",
        f"Loan Amount: ${prop.loan_amount:,.0f}",
        f"Equity Invested: ${prop.equity_invested:,.0f}",
        f"Gross Rental Income (Annual): ${prop.gross_rental_income:,.0f}",
        f"Vacancy Rate: {prop.vacancy_rate:.1%}",
        f"Operating Expenses (Annual): ${prop.operating_expenses:,.0f}",
        f"CapEx Reserve (Annual): ${prop.capital_expenditures:,.0f}",
        f"Interest Rate: {prop.interest_rate:.2%}",
        f"Loan Term: {prop.loan_term_years} years",
        f"Amortization: {prop.amortization_years} years",
        "",
        "=== CALCULATED METRICS ===",
        f"NOI (Year 1): ${metrics['noi']:,.0f}",
        f"Going-in Cap Rate: {metrics['cap_rate']:.2%}",
        f"Annual Debt Service: ${metrics['annual_debt_service']:,.0f}",
        f"DSCR: {metrics['dscr']:.2f}x",
        f"LTV: {metrics['ltv']:.1%}",
        f"Cash-on-Cash Return (Year 1): {metrics['cash_on_cash_yr1']:.2%}",
        f"Levered IRR ({prop.hold_period_years}-yr hold): "
        + (f"{metrics['irr']:.2%}" if metrics['irr'] is not None else "N/A"),
        f"Equity Multiple: {metrics['equity_multiple']:.2f}x",
    ]

    if metrics.get("price_per_unit"):
        lines.append(f"Price per Unit: ${metrics['price_per_unit']:,.0f}")
    if metrics.get("price_per_sqft"):
        lines.append(f"Price per SF: ${metrics['price_per_sqft']:.2f}")

    if prop.investment_thesis:
        lines += ["", "=== INVESTMENT THESIS ===", prop.investment_thesis]
    if prop.value_add_strategy:
        lines += ["", "=== VALUE-ADD STRATEGY ===", prop.value_add_strategy]
    if prop.key_risks:
        lines += ["", "=== IDENTIFIED RISKS ==="] + [f"- {r}" for r in prop.key_risks]

    return "\n".join(lines)


def _generate_one(args: tuple) -> tuple[str, str]:
    """Worker function: load prompt, build user message, call Claude for one section."""
    section_key, prompt_file, system_prompt, context = args
    section_prompt = _load_prompt(prompt_file)
    user_message = f"{section_prompt}\n\n{context}"
    return section_key, generate_section(system_prompt=system_prompt, user_prompt=user_message)


def generate_all_sections(prop: PropertyInput, metrics: dict | None = None) -> dict[str, str]:
    """
    Generate all narrative sections of the investment memo in parallel.

    Args:
        prop: Property input data.
        metrics: Pre-calculated financial metrics. If None, they are calculated internally.

    Returns:
        Dict mapping section name to generated text.
    """
    if metrics is None:
        metrics = calculate_all_metrics(prop)
    context = _build_property_context(prop, metrics)

    tasks = [
        ("executive_summary", "executive_summary.txt", _SYSTEM_PROMPT, context),
        ("market_analysis", "market_analysis.txt", _SYSTEM_PROMPT, context),
        ("investment_strategy", "investment_strategy.txt", _SYSTEM_PROMPT, context),
        ("risk_factors", "risk_factors.txt", _SYSTEM_PROMPT, context),
    ]

    sections: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(_generate_one, task): task[0] for task in tasks}
        for future in as_completed(futures):
            section_key, text = future.result()
            sections[section_key] = text

    return sections


def generate_memo_data(prop: PropertyInput) -> dict:
    """Return all data needed to render the memo (metrics + AI sections)."""
    metrics = calculate_all_metrics(prop)
    sections = generate_all_sections(prop, metrics=metrics)
    return {
        "property": prop,
        "metrics": metrics,
        "sections": sections,
    }
