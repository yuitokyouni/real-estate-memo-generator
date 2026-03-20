"""Unit tests for the financial calculator."""
import pytest
from memo_generator.models.property_input import PropertyInput
from memo_generator.financials.calculator import (
    calculate_noi,
    calculate_going_in_cap_rate,
    calculate_annual_debt_service,
    calculate_dscr,
    calculate_ltv,
    calculate_cash_on_cash,
    calculate_all_metrics,
)

SAMPLE = {
    "property_name": "Test Property",
    "property_type": "multifamily",
    "address": "123 Main St",
    "city": "Austin",
    "state_or_country": "TX, USA",
    "year_built": 2001,
    "purchase_price": 18_500_000,
    "gross_rental_income": 1_980_000,
    "vacancy_rate": 0.06,
    "operating_expenses": 750_000,
    "capital_expenditures": 60_000,
    "loan_amount": 13_875_000,
    "interest_rate": 0.0675,
    "loan_term_years": 10,
    "amortization_years": 30,
    "equity_invested": 4_625_000,
    "hold_period_years": 5,
    "exit_cap_rate": 0.055,
}


def test_noi():
    noi = calculate_noi(1_980_000, 0.06, 750_000)
    assert abs(noi - 1_111_200) < 1


def test_cap_rate():
    noi = calculate_noi(1_980_000, 0.06, 750_000)
    cap = calculate_going_in_cap_rate(noi, 18_500_000)
    assert 0.05 < cap < 0.07


def test_dscr_positive():
    ds = calculate_annual_debt_service(13_875_000, 0.0675, 30)
    noi = calculate_noi(1_980_000, 0.06, 750_000)
    dscr = calculate_dscr(noi, ds)
    assert dscr > 1.0, "DSCR should be above 1.0 for a viable deal"


def test_ltv():
    ltv = calculate_ltv(13_875_000, 18_500_000)
    assert abs(ltv - 0.75) < 0.001


def test_all_metrics_keys():
    prop = PropertyInput(**SAMPLE)
    metrics = calculate_all_metrics(prop)
    required_keys = {"noi", "cap_rate", "annual_debt_service", "dscr", "ltv",
                     "cash_on_cash_yr1", "irr", "equity_multiple", "cash_flows"}
    assert required_keys.issubset(metrics.keys())


def test_irr_reasonable():
    prop = PropertyInput(**SAMPLE)
    metrics = calculate_all_metrics(prop)
    assert metrics["irr"] is not None
    assert 0.05 < metrics["irr"] < 0.50, "IRR should be between 5% and 50% for a typical deal"


def test_cash_flows_length():
    prop = PropertyInput(**SAMPLE)
    metrics = calculate_all_metrics(prop)
    assert len(metrics["cash_flows"]) == prop.hold_period_years + 1
