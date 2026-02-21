"""
Rent vs Buy Calculator — Free Web Tool by ClearMetric
https://clearmetric.gumroad.com

Helps users decide whether renting or buying a home is better financially.
"""

import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Rent vs Buy Calculator — ClearMetric",
    page_icon="🏠",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom CSS (navy/blue theme matching FIRE calculator)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .stMetric { background: #f8f9fa; border-radius: 8px; padding: 12px; border-left: 4px solid #1a5276; }
    h1 { color: #1a5276; }
    h2, h3 { color: #2c3e50; }
    .verdict-buy { color: #27ae60; font-weight: bold; }
    .verdict-rent { color: #e74c3c; font-weight: bold; }
    .cta-box {
        background: linear-gradient(135deg, #1a5276 0%, #2e86c1 100%);
        color: white; padding: 24px; border-radius: 12px; text-align: center;
        margin: 20px 0;
    }
    .cta-box a { color: #f0d78c; text-decoration: none; font-weight: bold; font-size: 1.1rem; }
    div[data-testid="stSidebar"] { background: #f8f9fa; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("# 🏠 Rent vs Buy Calculator")
st.markdown("**Should you rent or buy?** Compare the true financial impact over time.")
st.markdown("---")

# ---------------------------------------------------------------------------
# Sidebar — User inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## Your Scenario")

    st.markdown("### Purchase Details")
    home_price = st.number_input("Home Purchase Price ($)", value=400_000, min_value=0, step=10_000, format="%d")
    down_pct = st.slider("Down Payment (%)", 0.0, 50.0, 20.0, 1.0) / 100
    mortgage_rate = st.slider("Mortgage Interest Rate (%)", 1.0, 15.0, 6.5, 0.25) / 100
    mortgage_term = st.selectbox("Mortgage Term", [15, 30], index=1)

    st.markdown("### Renting")
    monthly_rent = st.number_input("Monthly Rent ($)", value=2_000, min_value=0, step=100, format="%d")
    rent_increase = st.slider("Annual Rent Increase (%)", 0.0, 10.0, 3.0, 0.5) / 100

    st.markdown("### Ownership Costs")
    prop_tax_rate = st.slider("Property Tax Rate (%)", 0.0, 3.0, 1.2, 0.1) / 100
    home_insurance = st.number_input("Home Insurance ($/year)", value=1_500, min_value=0, step=100, format="%d")
    hoa_monthly = st.number_input("HOA Fees ($/month)", value=0, min_value=0, step=50, format="%d")
    maintenance_pct = st.slider("Home Maintenance (% of value/year)", 0.0, 3.0, 1.0, 0.1) / 100
    closing_costs_pct = st.slider("Closing Costs (%)", 0.0, 6.0, 3.0, 0.5) / 100

    st.markdown("### Assumptions")
    appreciation = st.slider("Home Appreciation Rate (%)", -2.0, 10.0, 3.5, 0.5) / 100
    investment_return = st.slider("Investment Return Rate (%)", 1.0, 15.0, 7.0, 0.5) / 100
    tax_bracket = st.slider("Income Tax Bracket (%)", 0.0, 50.0, 24.0, 1.0) / 100
    years = st.slider("Years to Compare", 1, 30, 10)

# Validation
if home_price > 0 and down_pct >= 1:
    st.sidebar.warning("⚠️ Down payment cannot be 100% or more.")
if monthly_rent == 0 and home_price > 0:
    st.sidebar.info("💡 Enter a rent amount to compare. Typical rent is 0.5–1% of home value per month.")

# ---------------------------------------------------------------------------
# Core calculations
# ---------------------------------------------------------------------------
down_payment = home_price * down_pct
loan_amount = home_price - down_payment
closing_costs = home_price * closing_costs_pct

# Monthly mortgage (P&I) — standard amortization
n_payments = mortgage_term * 12
monthly_rate = mortgage_rate / 12
if monthly_rate > 0 and loan_amount > 0:
    monthly_pi = loan_amount * (monthly_rate * (1 + monthly_rate) ** n_payments) / ((1 + monthly_rate) ** n_payments - 1)
else:
    monthly_pi = 0

# Annual ownership costs (fixed components in year 1)
annual_prop_tax = home_price * prop_tax_rate
annual_insurance = home_insurance
annual_hoa = hoa_monthly * 12
annual_maintenance = home_price * maintenance_pct

# Total monthly cost of buying (year 1)
monthly_buy_total = monthly_pi + (annual_prop_tax + annual_insurance + annual_hoa + annual_maintenance) / 12

# Build year-by-year data
def amortization_schedule(principal, annual_rate, years):
    """Return (balance_by_year, interest_paid_by_year, principal_paid_by_year)."""
    n = years * 12
    r = annual_rate / 12
    if r == 0:
        return [principal - principal * y / years for y in range(years + 1)], [0] * years, [principal / years] * years
    pmt = principal * (r * (1 + r) ** n) / ((1 + r) ** n - 1)
    balance = principal
    balances = [principal]
    interest_yr = []
    principal_yr = []
    for year in range(years):
        int_yr = 0
        prin_yr = 0
        for _ in range(12):
            interest = balance * r
            principal_paid = pmt - interest
            balance -= principal_paid
            int_yr += interest
            prin_yr += principal_paid
        balances.append(max(0, balance))
        interest_yr.append(int_yr)
        principal_yr.append(prin_yr)
    return balances, interest_yr, principal_yr

balances, interest_by_year, principal_by_year = amortization_schedule(loan_amount, mortgage_rate, mortgage_term)

# Extend if years > mortgage_term
if years > mortgage_term:
    for _ in range(years - mortgage_term):
        balances.append(0)
        interest_by_year.append(0)
        principal_by_year.append(0)

# Standard deduction (2024): assume single $14,600 / married $29,200 — use as threshold for itemizing
STANDARD_DEDUCTION = 14_600  # single; user can mentally adjust

rows = []
cum_cost_buy = down_payment + closing_costs
cum_cost_rent = 0
renter_wealth = down_payment  # Start with down payment invested
home_value = home_price

for y in range(years):
    # Home value (appreciates)
    home_value = home_price * (1 + appreciation) ** (y + 1)

    # Buyer costs this year
    if y < len(interest_by_year):
        annual_pi = monthly_pi * 12
        annual_interest = interest_by_year[y]
    else:
        annual_pi = 0
        annual_interest = 0

    prop_tax_yr = home_price * (1 + appreciation) ** y * prop_tax_rate
    maint_yr = home_price * (1 + appreciation) ** y * maintenance_pct
    annual_buy_cost = annual_pi + prop_tax_yr + annual_insurance + annual_hoa + maint_yr

    # Tax savings (itemized deduction for mortgage interest)
    itemized = annual_interest
    tax_savings = max(0, itemized - STANDARD_DEDUCTION) * tax_bracket if itemized > STANDARD_DEDUCTION else 0
    # Simplified: assume they itemize if mortgage interest > standard. Actually first year often itemize.
    # Use: tax_savings = annual_interest * tax_bracket (common simplification)
    tax_savings = annual_interest * tax_bracket

    net_buy_cost = annual_buy_cost - tax_savings
    cum_cost_buy += net_buy_cost

    # Renter costs
    rent_yr = monthly_rent * 12 * (1 + rent_increase) ** y
    cum_cost_rent += rent_yr

    # Renter: down payment invested. Each year, (buyer cost - rent) = savings when positive.
    # When buyer costs more, renter invests the difference. When rent > buyer cost, renter withdraws from investments.
    annual_savings = (annual_buy_cost - tax_savings) - rent_yr  # positive = renter has extra to invest
    renter_wealth = renter_wealth * (1 + investment_return) + annual_savings

    # Buyer equity
    if y < len(balances) - 1:
        mortgage_balance = balances[y + 1]
    else:
        mortgage_balance = 0
    buyer_equity = home_value - mortgage_balance

    rows.append({
        "Year": y + 1,
        "Home Value": home_value,
        "Mortgage Balance": mortgage_balance,
        "Buyer Equity": buyer_equity,
        "Cumulative Cost (Buy)": cum_cost_buy,
        "Cumulative Cost (Rent)": cum_cost_rent,
        "Renter Investments": renter_wealth,
        "Annual Buy Cost": annual_buy_cost,
        "Tax Savings": tax_savings,
        "Annual Rent": rent_yr,
    })

df = pd.DataFrame(rows)

# Break-even: first year when buyer equity > renter investments
break_even = None
for i, row in df.iterrows():
    if row["Buyer Equity"] > row["Renter Investments"]:
        break_even = int(row["Year"])
        break

# Verdict
if break_even is not None and break_even <= years:
    verdict = "buy"
    verdict_text = f"Buying is better after {break_even} years"
else:
    verdict = "rent"
    verdict_text = f"Renting is better for this {years}-year timeframe"

# ---------------------------------------------------------------------------
# Display — Key metrics
# ---------------------------------------------------------------------------
st.markdown("## Key Results")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Break-Even Year", f"{break_even}" if break_even else "N/A", help="Year when buying becomes cheaper than renting")
m2.metric("Total Cost (Buy)", f"${df.iloc[-1]['Cumulative Cost (Buy)']:,.0f}", help="Total cash spent on ownership over period")
m3.metric("Total Cost (Rent)", f"${df.iloc[-1]['Cumulative Cost (Rent)']:,.0f}", help="Total rent paid over period")
m4.metric("Monthly Payment (Buy)", f"${monthly_buy_total:,.0f}", help="P&I + tax + insurance + HOA + maintenance")

st.markdown("---")

# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------
st.markdown("## Cumulative Cost Comparison")
fig_cost = go.Figure()
fig_cost.add_trace(go.Scatter(x=df["Year"], y=df["Cumulative Cost (Buy)"], mode="lines", name="Buying", line=dict(color="#1a5276", width=3)))
fig_cost.add_trace(go.Scatter(x=df["Year"], y=df["Cumulative Cost (Rent)"], mode="lines", name="Renting", line=dict(color="#e74c3c", width=3)))
if break_even:
    fig_cost.add_vline(x=break_even, line_dash="dash", line_color="#27ae60", annotation_text=f"Break-even: Year {break_even}")
fig_cost.update_layout(
    yaxis_title="Cumulative Cost ($)", yaxis_tickformat="$,.0f",
    height=400, template="plotly_white", legend=dict(orientation="h", y=1.02),
    margin=dict(t=60, b=40),
)
st.plotly_chart(fig_cost, use_container_width=True)

st.markdown("## Net Wealth Comparison")
fig_wealth = go.Figure()
fig_wealth.add_trace(go.Scatter(x=df["Year"], y=df["Buyer Equity"], mode="lines", name="Buyer (Equity)", line=dict(color="#1a5276", width=3)))
fig_wealth.add_trace(go.Scatter(x=df["Year"], y=df["Renter Investments"], mode="lines", name="Renter (Investments)", line=dict(color="#27ae60", width=3)))
fig_wealth.update_layout(
    yaxis_title="Net Wealth ($)", yaxis_tickformat="$,.0f",
    height=400, template="plotly_white", legend=dict(orientation="h", y=1.02),
    margin=dict(t=60, b=40),
)
st.plotly_chart(fig_wealth, use_container_width=True)

st.markdown("## Monthly Cost Breakdown")
# Stacked bar: Year 1 vs Year 5 vs Year 10
sample_years = [1, min(5, years), min(10, years)]
sample_years = [y for y in sample_years if y <= years]
if not sample_years:
    sample_years = [1]

bar_data = []
for yr in sample_years:
    row = df.iloc[yr - 1]
    pi_yr = monthly_pi * 12 if yr <= mortgage_term else 0
    prop_tax_yr = home_price * (1 + appreciation) ** (yr - 1) * prop_tax_rate
    maint_yr = home_price * (1 + appreciation) ** (yr - 1) * maintenance_pct
    bar_data.append({
        "Year": yr,
        "P&I": pi_yr / 12,
        "Property Tax": prop_tax_yr / 12,
        "Insurance": annual_insurance / 12,
        "HOA": annual_hoa / 12,
        "Maintenance": maint_yr / 12,
        "Rent": row["Annual Rent"] / 12,
    })

fig_bar = go.Figure()
x_labels = [f"Year {y}" for y in sample_years]
colors = ["#1a5276", "#2e86c1", "#5dade2", "#aed6f1", "#d4e6f1"]
for i, comp in enumerate(["P&I", "Property Tax", "Insurance", "HOA", "Maintenance"]):
    fig_bar.add_trace(go.Bar(
        name=comp,
        x=x_labels,
        y=[b[comp] for b in bar_data],
        marker_color=colors[i % len(colors)],
    ))
fig_bar.add_trace(go.Bar(
    name="Rent",
    x=x_labels,
    y=[b["Rent"] for b in bar_data],
    marker_color="#e74c3c",
))
fig_bar.update_layout(barmode="stack", yaxis_title="Monthly Cost ($)", yaxis_tickformat="$,.0f",
                     height=400, template="plotly_white", legend=dict(orientation="h", y=1.02),
                     margin=dict(t=60, b=40))
st.plotly_chart(fig_bar, use_container_width=True)

# ---------------------------------------------------------------------------
# Year-by-year table
# ---------------------------------------------------------------------------
with st.expander("📊 Year-by-Year Comparison", expanded=False):
    display_df = df[["Year", "Home Value", "Buyer Equity", "Cumulative Cost (Buy)", "Cumulative Cost (Rent)", "Renter Investments"]]
    st.dataframe(
        display_df.style.format({
            "Home Value": "${:,.0f}",
            "Buyer Equity": "${:,.0f}",
            "Cumulative Cost (Buy)": "${:,.0f}",
            "Cumulative Cost (Rent)": "${:,.0f}",
            "Renter Investments": "${:,.0f}",
        }),
        use_container_width=True,
        height=400,
    )

# ---------------------------------------------------------------------------
# Verdict
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("## Verdict")
if verdict == "buy":
    st.success(f"**{verdict_text}.** Based on your inputs, buying becomes the better financial choice after year {break_even}.")
else:
    st.warning(f"**{verdict_text}.** For a {years}-year horizon, renting and investing the difference comes out ahead.")

# ---------------------------------------------------------------------------
# CTA — Paid Product
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("""
<div class="cta-box">
    <h3 style="color: white; margin: 0 0 8px 0;">Want the Full Excel Spreadsheet?</h3>
    <p style="margin: 0 0 16px 0;">
        Get the <strong>ClearMetric Rent vs Buy Calculator</strong> — a downloadable Excel template with:<br>
        ✓ All inputs in one place with teal input cells<br>
        ✓ Year-by-year comparison (30 years)<br>
        ✓ What-If Scenarios: current market, optimistic, conservative<br>
        ✓ How To Use guide with input explanations<br>
    </p>
    <a href="https://clearmetric.gumroad.com" target="_blank">
        Get It on Gumroad — $12.99 →
    </a>
</div>
""", unsafe_allow_html=True)

# Cross-sell
st.markdown("### More from ClearMetric")
cx1, cx2, cx3 = st.columns(3)
with cx1:
    st.markdown("""
    **🏠 Rental Property Analyzer** — $19.99
    12+ metrics, 5-year projection, 4-property comparison.
    [Get it →](https://clearmetric.gumroad.com)
    """)
with cx2:
    st.markdown("""
    **📊 Budget Planner** — $13.99
    Track income, expenses, savings with the 50/30/20 framework.
    [Get it →](https://clearmetric.gumroad.com)
    """)
with cx3:
    st.markdown("""
    **🔥 FIRE Calculator** — $14.99
    Find your FIRE number, scenario comparison, year-by-year projection.
    [Get it →](https://clearmetric.gumroad.com)
    """)

# Footer
st.markdown("---")
st.caption("© 2026 ClearMetric | [clearmetric.gumroad.com](https://clearmetric.gumroad.com) | "
           "This tool is for educational purposes only. Not financial advice. Consult a qualified financial advisor.")
