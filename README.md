# Binomial-Black-Scholes-Option-Model-in-Excel

This project implements a dynamic binomial tree framework in Excel to price European-style call and put options, simulate multi-period stock price movements, and demonstrate real-time delta hedging strategies. It also includes a Black-Scholes pricing comparison to validate model accuracy and strengthen understanding of financial risk management techniques.

---

## Features

- Stock price simulation using a multi-period binomial tree
- European option valuation via backward induction
- Delta hedging strategy with replicating portfolio (stock and bond)
- Fully dynamic inputs: stock price, strike, up/down factors, interest rate, and number of periods
- Black-Scholes analytical comparison for model validation
- Clean and structured layout to support learning and experimentation

---

## Inputs

| Parameter           | Description                                |
|---------------------|--------------------------------------------|
| Initial Stock Price | Starting value of the stock (S₀)           |
| Strike Price        | Exercise price of the option (K)           |
| Up Factor (u)       | Multiplier for upward price movement       |
| Down Factor (d)     | Multiplier for downward price movement     |
| Risk-Free Rate (r)  | Interest rate used for discounting         |
| Time Periods (N)    | Number of steps in the binomial model      |
| Volatility (σ)      | Used in the Black-Scholes price calculation|

---

## Model Overview

1. **Stock Price Tree**: Constructed using user-defined up/down factors and number of periods.
2. **Option Payoff**: Calculated at the terminal nodes for both call and put options.
3. **Backward Induction**: Option values are computed recursively at each node using risk-neutral valuation.
4. **Delta Hedging Table**: The hedge ratio (Δ), bond value (B), and portfolio value are computed at each node to simulate a replicating strategy.
5. **Black-Scholes Comparison**: Provides the analytical price of the option using the closed-form Black-Scholes formula to compare against the binomial result.

---

## Files Included

- `Binomial-Black-Scholes-Option-Model.xlsx` — Main Excel workbook including all trees, inputs, hedging logic, and comparison formulas.

---

## Concepts Covered

- Binomial Tree Option Pricing
- Black-Scholes Analytical Formula
- Risk-Neutral Valuation
- Delta Hedging and Portfolio Replication
- Financial Risk Management

---

## Technologies Used

- Microsoft Excel
- Excel Functions: `IF`, `POWER`, `EXP`, `LN`, `NORM.S.DIST`
- Dynamic referencing with `INDEX`, `MATCH`, and conditional logic

---

## Author

**Diya Bajaj**  
Undergraduate in Financial Mathematics & Analytics  
Specializing in Financial Risk Management

---

## Notes

- The model supports European-style options only.
- Increasing the number of periods improves accuracy and convergence toward the Black-Scholes value.
- The hedging strategy is based on discrete-time delta replication.

