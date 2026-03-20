"""Financial calculation engine for real estate investment analysis."""
from __future__ import annotations

from typing import TYPE_CHECKING, Optional

import numpy_financial as npf

if TYPE_CHECKING:
    from memo_generator.models.property_input import PropertyInput


def calculate_noi(
    gross_rental_income: float,
    vacancy_rate: float,
    operating_expenses: float,
) -> float:
    """Net Operating Income = Effective Gross Income - Operating Expenses."""
    egi = gross_rental_income * (1 - vacancy_rate)
    return egi - operating_expenses


def calculate_going_in_cap_rate(noi: float, purchase_price: float) -> float:
    """Going-in Cap Rate = NOI / Purchase Price."""
    return noi / purchase_price


def calculate_annual_debt_service(
    loan_amount: float,
    interest_rate: float,
    amortization_years: int,
) -> float:
    """Annual debt service using numpy_financial.pmt (fully amortizing)."""
    monthly_rate = interest_rate / 12
    n_periods = amortization_years * 12
    monthly_payment = npf.pmt(monthly_rate, n_periods, -loan_amount)
    return monthly_payment * 12


def calculate_dscr(noi: float, annual_debt_service: float) -> float:
    """Debt Service Coverage Ratio = NOI / Annual Debt Service."""
    return noi / annual_debt_service


def calculate_ltv(loan_amount: float, purchase_price: float) -> float:
    """Loan-to-Value Ratio = Loan Amount / Purchase Price."""
    return loan_amount / purchase_price


def calculate_cash_on_cash(
    pre_tax_cash_flow: float,
    equity_invested: float,
) -> float:
    """Cash-on-Cash Return = Pre-tax Cash Flow / Equity Invested."""
    return pre_tax_cash_flow / equity_invested


def _remaining_loan_balance(
    loan_amount: float,
    interest_rate: float,
    amortization_years: int,
    years_elapsed: int,
) -> float:
    """Remaining loan balance after years_elapsed using PV of remaining payments."""
    monthly_rate = interest_rate / 12
    total_periods = amortization_years * 12
    elapsed_periods = years_elapsed * 12
    remaining_periods = total_periods - elapsed_periods

    if remaining_periods <= 0:
        return 0.0

    monthly_payment = npf.pmt(monthly_rate, total_periods, -loan_amount)
    remaining_balance = npf.pv(monthly_rate, remaining_periods, monthly_payment)
    return abs(remaining_balance)


def calculate_5year_cash_flows(prop: "PropertyInput") -> list[float]:
    """
    Build the levered cash flow series for IRR calculation.

    Returns:
        List of floats: [-equity_invested, cf_yr1, ..., cf_yr4, cf_yr5 + exit_proceeds]
    """
    noi_yr0 = calculate_noi(
        prop.gross_rental_income, prop.vacancy_rate, prop.operating_expenses
    )
    annual_ds = calculate_annual_debt_service(
        prop.loan_amount, prop.interest_rate, prop.amortization_years
    )
    cap_rate = calculate_going_in_cap_rate(noi_yr0, prop.purchase_price)
    exit_cap = prop.exit_cap_rate if prop.exit_cap_rate is not None else cap_rate

    cash_flows: list[float] = [-prop.equity_invested]

    for year in range(1, prop.hold_period_years + 1):
        noi = noi_yr0 * ((1 + prop.rent_growth_annual) ** year)
        capex = prop.capital_expenditures * ((1 + prop.expense_growth_annual) ** year)
        cf = noi - annual_ds - capex
        cash_flows.append(cf)

    # Add exit proceeds to final year
    noi_exit = noi_yr0 * ((1 + prop.rent_growth_annual) ** prop.hold_period_years)
    exit_value = noi_exit / exit_cap
    remaining_debt = _remaining_loan_balance(
        prop.loan_amount, prop.interest_rate, prop.amortization_years, prop.hold_period_years
    )
    exit_proceeds = exit_value - remaining_debt
    cash_flows[-1] += exit_proceeds

    return cash_flows


def calculate_irr(cash_flows: list[float]) -> Optional[float]:
    """Levered IRR using numpy_financial.irr. Returns None if calculation fails."""
    try:
        result = npf.irr(cash_flows)
        if result is None or result != result:  # NaN check
            return None
        return float(result)
    except Exception:
        return None


def calculate_equity_multiple(cash_flows: list[float]) -> float:
    """Equity Multiple = sum of all positive cash flows / initial equity invested."""
    equity_invested = abs(cash_flows[0])
    total_inflows = sum(cf for cf in cash_flows[1:] if cf > 0)
    return total_inflows / equity_invested


def calculate_all_metrics(prop: "PropertyInput") -> dict:
    """Calculate all key investment metrics and return as a dictionary."""
    noi = calculate_noi(prop.gross_rental_income, prop.vacancy_rate, prop.operating_expenses)
    cap_rate = calculate_going_in_cap_rate(noi, prop.purchase_price)
    annual_ds = calculate_annual_debt_service(
        prop.loan_amount, prop.interest_rate, prop.amortization_years
    )
    dscr = calculate_dscr(noi, annual_ds)
    ltv = calculate_ltv(prop.loan_amount, prop.purchase_price)
    pre_tax_cf = noi - annual_ds - prop.capital_expenditures
    coc = calculate_cash_on_cash(pre_tax_cf, prop.equity_invested)
    cash_flows = calculate_5year_cash_flows(prop)
    irr = calculate_irr(cash_flows)
    equity_multiple = calculate_equity_multiple(cash_flows)

    price_per_unit = (
        prop.purchase_price / prop.total_units if prop.total_units else None
    )
    price_per_sqft = (
        prop.purchase_price / prop.total_sqft if prop.total_sqft else None
    )

    return {
        "noi": noi,
        "cap_rate": cap_rate,
        "annual_debt_service": annual_ds,
        "dscr": dscr,
        "ltv": ltv,
        "pre_tax_cash_flow_yr1": pre_tax_cf,
        "cash_on_cash_yr1": coc,
        "irr": irr,
        "equity_multiple": equity_multiple,
        "cash_flows": cash_flows,
        "price_per_unit": price_per_unit,
        "price_per_sqft": price_per_sqft,
    }
