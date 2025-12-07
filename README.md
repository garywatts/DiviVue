📈 **DiviVue — Open-Source Google Sheets Portfolio Tracker**
_______________________________________________
DiviVue is a lightweight, open-source stock portfolio tracker built using Google Sheets and Google Apps Script. The project was created as a free alternative to the many paid online portfolio trackers currently on the market. It leverages the power of the GOOGLEFINANCE function and a structured ledger system to provide a simple, transparent way to track your investments.

This tool is aimed at everyday investors who want something they can host and control themselves. I’m not a finance professional, and the current version is fairly basic in areas such as CGT and taxable income calculations — this is where I hope the community can help expand and improve the project over time.

DiviVue is fully open-source, and contributions are welcome. If you find it useful, feel free to support the project via **[Buy Me a Coffee](https://www.buymeacoffee.com/garywatts)**

See the Wiki above for install instructions <a href="https://github.com/garywatts/DiviVue/wiki">https://github.com/garywatts/DiviVue/wiki</a>

🎯 **Project Goals**
_______________________________________________
Most core features have already been built, with more improvements planned. DiviVue currently aims to:

✔ **Portfolio & Trading**

Add and track buy and sell trades, dividends, and DRP entries, including manual price entry for DRP allocations.

Maintain holdings with realised and unrealised P&L.

Perform capital gains calculations using FIFO rules, including the 12-month 50% CGT discount for long-term holdings.

✔ **Income Tracking**

Track dividends, franking credits, unfranked components, and DRP reinvestments.

Generate a taxable income report (financial-year dividends + franking).

✔ **Capital Gains**

Generate capital gains reports with a selectable date range.

✔ **Portfolio Insights**

Provide a performance chart over time (placeholder chart with date selector in place).

Include per-stock detail pages showing holdings, transactions, and returns.

✔ **Architecture**

Use a hybrid calculation model where Apps Script performs the heavy lifting (as much as possible) and writes helper columns back to the Sheet.

Treat Google Sheets as the raw ledger, not just a calculator.

Leverage Google’s GOOGLEFINANCE for live market data where supported.

Include wiring for live ASX price lookup via UrlFetch as a fallback for GoogleFinance limitations.

✔ **Dashboard**

Provide a top-level dashboard showing:

Total portfolio value

Total gain/loss

Dividends YTD

🤝 **Contributing**
_______________________________________________

DiviVue is designed to evolve with community input. Improvements especially welcome in:

CGT rules and expansion from FIFO

Taxable income calculations

Multi-country support

Performance charting

UI/UX enhancements

If you’d like to support development, contributions, pull requests, and feature ideas are all appreciated — or you can shout me a coffee. **[Buy Me a Coffee](https://www.buymeacoffee.com/garywatts)**
