"""Pydantic v2 input schema for property data."""
from typing import Literal, Optional
from pydantic import BaseModel, Field, field_validator


class PropertyInput(BaseModel):
    # --- Identification ---
    property_name: str = Field(..., description="Name of the property")
    property_type: Literal["multifamily", "office", "retail", "industrial", "mixed-use"]
    address: str
    city: str
    state_or_country: str
    year_built: int = Field(..., ge=1800, le=2030)

    # --- Core Financials ---
    purchase_price: float = Field(..., gt=0)
    gross_rental_income: float = Field(..., gt=0, description="Annual GRI in USD")
    vacancy_rate: float = Field(..., ge=0.0, le=1.0)
    operating_expenses: float = Field(..., gt=0, description="Annual OpEx in USD")
    capital_expenditures: float = Field(..., ge=0, description="Annual CapEx reserve in USD")

    # --- Debt ---
    loan_amount: float = Field(..., ge=0)
    interest_rate: float = Field(..., ge=0.0, le=1.0)
    loan_term_years: int = Field(..., ge=1, le=40)
    amortization_years: int = Field(..., ge=1, le=40)

    # --- Equity ---
    equity_invested: float = Field(..., gt=0)

    # --- Optional Property Details ---
    total_units: Optional[int] = Field(None, gt=0)
    total_sqft: Optional[float] = Field(None, gt=0)
    occupancy_rate: Optional[float] = Field(None, ge=0.0, le=1.0)

    # --- Exit Assumptions ---
    hold_period_years: int = Field(5, ge=1, le=30)
    exit_cap_rate: Optional[float] = Field(None, ge=0.0, le=1.0)
    rent_growth_annual: float = Field(0.03, ge=0.0, le=1.0)
    expense_growth_annual: float = Field(0.025, ge=0.0, le=1.0)

    # --- Market Context ---
    market_cap_rate: Optional[float] = Field(None, ge=0.0, le=1.0)

    # --- Narrative Inputs ---
    investment_thesis: Optional[str] = None
    key_risks: list[str] = Field(default_factory=list)
    value_add_strategy: Optional[str] = None

    @field_validator("loan_amount")
    @classmethod
    def loan_not_exceed_price(cls, v, info):
        purchase_price = info.data.get("purchase_price")
        if purchase_price and v > purchase_price:
            raise ValueError("loan_amount must not exceed purchase_price")
        return v

    @field_validator("equity_invested")
    @classmethod
    def equity_consistent(cls, v, info):
        purchase_price = info.data.get("purchase_price")
        loan_amount = info.data.get("loan_amount")
        if purchase_price and loan_amount:
            expected = purchase_price - loan_amount
            if abs(v - expected) > 1.0:  # allow $1 rounding tolerance
                raise ValueError(
                    f"equity_invested ({v:,.0f}) should equal "
                    f"purchase_price - loan_amount ({expected:,.0f})"
                )
        return v
