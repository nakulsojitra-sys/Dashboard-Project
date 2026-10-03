# Ledgerline Financial Dashboard

Ledgerline is a single-page financial analytics dashboard built with Vite and vanilla JavaScript. It uses a deterministic, fictional dataset of 2,400 2025 transactions and labels the data clearly as illustrative/demo data.

## Run locally

    npm install
    npm run generate:data
    npm run dev

Open the address Vite prints, typically http://localhost:5173.

## Included interactions

- Combined date, category, region, department, transaction type, status, amount, and search filters
- Filter-driven KPIs, insights, records, and Chart.js visualizations
- Daily, weekly, monthly, and quarterly trend aggregation
- Sortable transaction columns and pagination
- CSV, Excel, and PDF exports of the selected ledger
- Empty and data-load error states

## Project structure

    src/main.js                    Dashboard state, filters, calculations, exports, and chart rendering
    src/styles.css                 Responsive editorial visual system
    scripts/generate-data.mjs      Deterministic financial-record dataset generator
    public/data/financial_records.json

The data generator maintains the relationship Net_Amount = Amount + Tax. Recognized revenue and operating expense metrics include completed income and expense transactions respectively.
