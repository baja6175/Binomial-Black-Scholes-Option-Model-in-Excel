# 📈 Binomial-Black-Scholes-Option-Model-in-Excel

This Excel-based project implements a dynamic binomial tree framework to price European-style call and put options, simulate multi-period stock price movements, and demonstrate real-time delta hedging strategies. It also includes a Black-Scholes pricing comparison to validate model accuracy and deepen understanding of financial risk management.

---

## 🚀 Features

- 📊 **Stock Price Simulation** using a multi-period binomial tree
- 💰 **European Option Valuation** via backward induction
- ⚖️ **Delta Hedging Strategy** with real-time stock and bond positions
- 🔄 **Dynamic Inputs**: update stock price, strike, u/d factors, interest rate, and periods
- 📉 **Black-Scholes Comparison** for analytical pricing benchmark
- 📐 **Clean and Visual Tree Layout** for intuitive understanding

---

## 📥 Inputs

| Parameter           | Description                          |
|---------------------|--------------------------------------|
| Initial Stock Price | Starting value of the stock (S₀)     |
| Strike Price        | Exercise price of the option (K)     |
| Up Factor (u)       | Upward movement multiplier           |
| Down Factor (d)     | Downward movement multiplier         |
| Risk-Free Rate (r)  | Interest rate used for discounting   |
| Time Periods (N)    | Number of periods in the tree        |
| Volatility (σ)      | Used for Black-Scholes calculation   |

---

## 📘 How It Works

1. **Stock Tree** is built based on binomial up/down logic from user inputs
2. **Option Payoff** is calculated at the final nodes using payoff formulas
3. **Backward Induction** applies risk-neutral valuation to calculate earlier node values
4. **Delta Hedging** is computed using the replicating portfolio method:
   - Delta (Δ)
   - Bond value (B)
   - Portfolio value = Δ × Stock + B
5. **Black-Scholes Formula** is used to compute an analytical option price and compare with the binomial result

---

## 📎 Files Included

- `[Binomial-Black-Scholes-Option-Model.xlsx](https://github.com/user-attachments/files/20140639/Binomial-Black-Scholes-Option-Model.xlsx).xlsx` — the main Excel workbook with stock tree, option tree, delta hedge table, and Black-Scholes calculator

---

## 📚 Concepts Covered

- Binomial Option Pricing Model  
- Black-Scholes Model  
- Risk-Neutral Valuation  
- Delta Hedging and Replication  
- Financial Risk Management

---

## 🛠 Technologies Used

- **Microsoft Excel**  
- **Excel Formulas**: `IF`, `POWER`, `NORM.S.DIST`, `EXP`, `LN`, and custom payoff logic  
- Optional use of `VLOOKUP`, `INDEX`, and `MATCH` for dynamic referencing

---

## 📌 Author

**Diya Bajaj**  
Undergraduate in Financial Mathematics & Analytics  
Specializing in Financial Risk Management  

---

## 🧠 Notes

- The model assumes European-style options (no early exercise).
- Increasing the number of periods (N) improves convergence toward the Black-Scholes price.
- Risk-neutral probabilities are derived from user-defined up/down factors and the risk-free rate.

---

