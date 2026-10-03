import Chart from "chart.js/auto";
import { createIcons, icons } from "lucide";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import "./styles.css";

const DATA_URL = "/data/financial_records.json";
const DEFAULT_FILTERS = {
  startDate: "2025-01-01",
  endDate: "2025-12-31",
  category: "",
  region: "",
  department: "",
  transactionType: "",
  status: "",
  minAmount: "",
  maxAmount: "",
  search: ""
};

const FILTER_OPTIONS = {
  category: ["Revenue", "Expense", "Investment", "Operations", "Other"],
  region: ["North", "South", "East", "West", "International"],
  department: ["Finance", "Marketing", "Operations", "HR", "Technology", "Sales"],
  transactionType: ["Income", "Expense", "Transfer", "Investment"],
  status: ["Completed", "Pending", "Failed", "Refunded"]
};

const PALETTE = {
  ink: "#17312b",
  teal: "#006b67",
  coral: "#d45a44",
  gold: "#b98b2e",
  moss: "#6d8b73",
  mist: "#d9e4de",
  clay: "#b96f5d",
  quiet: "#75817b",
  pale: "#edf2ee"
};

const state = {
  records: [],
  filters: { ...DEFAULT_FILTERS },
  grain: "Monthly",
  sort: { key: "Date", direction: "desc" },
  page: 1,
  pageSize: 10,
  charts: [],
  loadError: false
};

const app = document.querySelector("#app");
const compactCurrency = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  notation: "compact",
  maximumFractionDigits: 1
});
const preciseCurrency = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  maximumFractionDigits: 0
});
const numberFormat = new Intl.NumberFormat("en-US");
const fullDateFormat = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric",
  year: "numeric"
});
const shortDateFormat = new Intl.DateTimeFormat("en-US", {
  month: "short",
  day: "numeric"
});

function sum(records, field) {
  return records.reduce((total, record) => total + Number(record[field] || 0), 0);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function asDate(value) {
  return new Date(value + "T12:00:00");
}

function formatDate(value, compact = false) {
  return (compact ? shortDateFormat : fullDateFormat).format(asDate(value));
}

function formatPeriod(value) {
  const date = new Date(value + "T12:00:00");
  return new Intl.DateTimeFormat("en-US", { month: "short", year: "numeric" }).format(date);
}

function formatCurrency(value, compact = true) {
  return (compact ? compactCurrency : preciseCurrency).format(value || 0);
}

function formatPercent(value, fallback = "N/A") {
  if (!Number.isFinite(value)) return fallback;
  const sign = value > 0 ? "+" : "";
  return sign + value.toFixed(1) + "%";
}

function pluralize(count, word) {
  return numberFormat.format(count) + " " + word + (count === 1 ? "" : "s");
}

function filteredRecords({ ignoreDates = false } = {}) {
  const filters = state.filters;
  const query = filters.search.trim().toLowerCase();

  return state.records.filter((record) => {
    if (!ignoreDates && (record.Date < filters.startDate || record.Date > filters.endDate)) return false;
    if (filters.category && record.Category !== filters.category) return false;
    if (filters.region && record.Region !== filters.region) return false;
    if (filters.department && record.Department !== filters.department) return false;
    if (filters.transactionType && record.Transaction_Type !== filters.transactionType) return false;
    if (filters.status && record.Status !== filters.status) return false;
    if (filters.minAmount && record.Net_Amount < Number(filters.minAmount)) return false;
    if (filters.maxAmount && record.Net_Amount > Number(filters.maxAmount)) return false;

    if (query) {
      const haystack = [
        record.Description,
        record.Customer,
        record.Category,
        record.Subcategory,
        record.Department,
        record.Region,
        record.Transaction_ID
      ]
        .join(" ")
        .toLowerCase();
      if (!haystack.includes(query)) return false;
    }

    return true;
  });
}

function calculateGrowth(records) {
  if (records.length < 8) return null;

  const orderedDates = records.map((record) => record.Date).sort();
  const start = asDate(orderedDates[0]).getTime();
  const end = asDate(orderedDates[orderedDates.length - 1]).getTime();
  const midpoint = start + (end - start) / 2;
  const early = records.filter((record) => asDate(record.Date).getTime() <= midpoint);
  const late = records.filter((record) => asDate(record.Date).getTime() > midpoint);
  const earlyRevenue = sum(early, "Revenue");
  const lateRevenue = sum(late, "Revenue");

  if (!earlyRevenue || !late.length) return null;
  return ((lateRevenue - earlyRevenue) / earlyRevenue) * 100;
}

function calculateMetrics(records) {
  const revenue = sum(records, "Revenue");
  const expenses = sum(records, "Expense");
  const profit = revenue - expenses;
  const completed = records.filter((record) => record.Status === "Completed");
  const grossVolume = sum(records, "Net_Amount");

  return {
    revenue,
    expenses,
    profit,
    margin: revenue ? (profit / revenue) * 100 : null,
    transactionCount: records.length,
    averageTransaction: records.length ? grossVolume / records.length : 0,
    completedCount: completed.length,
    growth: calculateGrowth(records)
  };
}

function groupBy(records, valueFn, metricFn) {
  const grouped = new Map();
  records.forEach((record) => {
    const key = valueFn(record);
    grouped.set(key, (grouped.get(key) || 0) + metricFn(record));
  });
  return Array.from(grouped.entries())
    .map(([label, value]) => ({ label, value }))
    .sort((left, right) => right.value - left.value);
}

function groupCount(records, valueFn) {
  const grouped = new Map();
  records.forEach((record) => {
    const key = valueFn(record);
    grouped.set(key, (grouped.get(key) || 0) + 1);
  });
  return Array.from(grouped.entries())
    .map(([label, value]) => ({ label, value }))
    .sort((left, right) => right.value - left.value);
}

function getWeekStart(dateValue) {
  const date = asDate(dateValue);
  const day = date.getDay();
  const shift = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + shift);
  return date.toISOString().slice(0, 10);
}

function trendBuckets(records) {
  const buckets = new Map();

  records.forEach((record) => {
    let key = record.Date;
    let label = formatDate(record.Date, true);

    if (state.grain === "Weekly") {
      key = getWeekStart(record.Date);
      label = formatDate(key, true);
    }

    if (state.grain === "Monthly") {
      key = record.Date.slice(0, 7);
      label = new Intl.DateTimeFormat("en-US", { month: "short" }).format(asDate(record.Date));
    }

    if (state.grain === "Quarterly") {
      key = record.Quarter;
      label = record.Quarter;
    }

    if (!buckets.has(key)) {
      buckets.set(key, { key, label, revenue: 0, expenses: 0, profit: 0 });
    }

    const bucket = buckets.get(key);
    bucket.revenue += record.Revenue;
    bucket.expenses += record.Expense;
    bucket.profit += record.Profit;
  });

  return Array.from(buckets.values()).sort((left, right) => left.key.localeCompare(right.key));
}

function topBreakdowns(records) {
  const expenseCategory = groupBy(
    records.filter((record) => record.Expense > 0),
    (record) => record.Subcategory,
    (record) => record.Expense
  ).slice(0, 6);
  const regionRevenue = groupBy(
    records.filter((record) => record.Revenue > 0),
    (record) => record.Region,
    (record) => record.Revenue
  );
  const departmentProfit = groupBy(records, (record) => record.Department, (record) => record.Profit);
  const status = groupCount(records, (record) => record.Status);

  return { expenseCategory, regionRevenue, departmentProfit, status };
}

function meaningfulInsights(records, metrics, breakdowns) {
  if (!records.length) return [];

  const insights = [];
  const topExpense = breakdowns.expenseCategory[0];
  const topRegion = breakdowns.regionRevenue[0];
  const topDepartment = breakdowns.departmentProfit[0];
  const pendingCount = records.filter((record) => record.Status === "Pending").length;

  if (topExpense && metrics.expenses) {
    const share = (topExpense.value / metrics.expenses) * 100;
    insights.push({
      label: "Spend concentration",
      value: topExpense.label,
      detail: formatPercent(share, "0.0%").replace("+", "") + " of completed expenses"
    });
  }

  if (topRegion && metrics.revenue) {
    const share = (topRegion.value / metrics.revenue) * 100;
    insights.push({
      label: "Revenue leader",
      value: topRegion.label + " region",
      detail: formatCurrency(topRegion.value) + " | " + share.toFixed(1) + "% of revenue"
    });
  }

  if (topDepartment && topDepartment.value !== 0) {
    insights.push({
      label: "Profit contributor",
      value: topDepartment.label,
      detail: formatCurrency(topDepartment.value) + " net contribution"
    });
  }

  if (pendingCount) {
    insights.push({
      label: "Settlement watch",
      value: pluralize(pendingCount, "transaction") + " pending",
      detail: ((pendingCount / records.length) * 100).toFixed(1) + "% of selected records"
    });
  }

  return insights.slice(0, 4);
}

function selectOptions(items, selected, placeholder) {
  const options = ['<option value="">' + placeholder + "</option>"];
  items.forEach((item) => {
    options.push(
      '<option value="' +
        escapeHtml(item) +
        '"' +
        (item === selected ? " selected" : "") +
        ">" +
        escapeHtml(item) +
        "</option>"
    );
  });
  return options.join("");
}

function activeFilterChips() {
  const labels = {
    category: "Category",
    region: "Region",
    department: "Department",
    transactionType: "Type",
    status: "Status",
    minAmount: "Min amount",
    maxAmount: "Max amount",
    search: "Search"
  };
  const chips = [];

  Object.entries(labels).forEach(([key, label]) => {
    const value = state.filters[key];
    if (!value) return;
    const display = key.includes("Amount") ? formatCurrency(Number(value), false) : value;
    chips.push(
      '<button class="filter-chip" type="button" data-remove-filter="' +
        key +
        '" aria-label="Remove ' +
        escapeHtml(label) +
        ' filter">' +
        escapeHtml(label) +
        ": " +
        escapeHtml(display) +
        '<i data-lucide="x" aria-hidden="true"></i></button>'
    );
  });

  const datesAreCustom =
    state.filters.startDate !== DEFAULT_FILTERS.startDate || state.filters.endDate !== DEFAULT_FILTERS.endDate;
  if (datesAreCustom) {
    chips.unshift(
      '<button class="filter-chip" type="button" data-remove-filter="dates" aria-label="Remove date range filter">' +
        escapeHtml(formatPeriod(state.filters.startDate) + " to " + formatPeriod(state.filters.endDate)) +
        '<i data-lucide="x" aria-hidden="true"></i></button>'
    );
  }

  return chips.length
    ? '<div class="active-filters"><span>Active</span>' + chips.join("") + "</div>"
    : "";
}

function metricBlock(label, value, note, tone = "") {
  return (
    '<article class="metric-block ' +
    tone +
    '">' +
    '<p class="metric-label">' +
    label +
    "</p>" +
    '<p class="metric-value">' +
    value +
    "</p>" +
    '<p class="metric-note">' +
    note +
    "</p>" +
    "</article>"
  );
}

function renderKpis(metrics) {
  const growthNote =
    metrics.growth === null
      ? "Selected period"
      : formatPercent(metrics.growth) + " late-period revenue";
  const marginNote =
    metrics.margin === null
      ? "No recognized revenue"
      : metrics.margin.toFixed(1) + "% retained after expenses";

  return (
    '<section class="kpi-section" aria-label="Key financial metrics">' +
    metricBlock("Recognized revenue", formatCurrency(metrics.revenue), growthNote, "is-teal") +
    metricBlock(
      "Operating expenses",
      formatCurrency(metrics.expenses),
      metrics.expenses ? (metrics.expenses / Math.max(metrics.revenue, 1) * 100).toFixed(1) + "% of revenue" : "No completed expenses",
      "is-coral"
    ) +
    metricBlock("Net profit", formatCurrency(metrics.profit), metrics.profit >= 0 ? "Positive operating result" : "Operating loss", "is-ink") +
    metricBlock("Profit margin", metrics.margin === null ? "N/A" : metrics.margin.toFixed(1) + "%", marginNote, "is-gold") +
    metricBlock("Transactions", numberFormat.format(metrics.transactionCount), numberFormat.format(metrics.completedCount) + " completed", "is-ink") +
    metricBlock("Average transaction", formatCurrency(metrics.averageTransaction), "Gross selected value", "is-ink") +
    "</section>"
  );
}

function renderLedgerBalance(metrics) {
  const balance = metrics.revenue - metrics.expenses;
  const balanceLabel = balance >= 0 ? "In balance" : "Needs attention";
  return (
    '<aside class="ledger-balance" aria-label="Filtered ledger balance">' +
    '<div class="eyebrow">Selected ledger</div>' +
    '<div class="balance-state"><span class="state-dot"></span>' +
    escapeHtml(balanceLabel) +
    "</div>" +
    '<p class="balance-amount">' +
    formatCurrency(balance) +
    "</p>" +
    '<dl class="balance-list">' +
    "<div><dt>Revenue</dt><dd>" +
    formatCurrency(metrics.revenue, false) +
    "</dd></div>" +
    "<div><dt>Expenses</dt><dd>" +
    formatCurrency(metrics.expenses, false) +
    "</dd></div>" +
    "<div><dt>Completed</dt><dd>" +
    numberFormat.format(metrics.completedCount) +
    "</dd></div>" +
    "</dl>" +
    "</aside>"
  );
}

function renderLargestTransactions(records) {
  const largest = [...records].sort((left, right) => right.Net_Amount - left.Net_Amount).slice(0, 5);
  if (!largest.length) {
    return '<p class="quiet-empty">No transactions within the current selection.</p>';
  }

  return (
    '<ol class="transaction-list">' +
    largest
      .map((record) => {
        return (
          "<li>" +
          '<div class="transaction-list-main"><strong>' +
          escapeHtml(record.Description) +
          "</strong><span>" +
          escapeHtml(record.Category + " | " + formatDate(record.Date, true)) +
          "</span></div>" +
          "<b>" +
          formatCurrency(record.Net_Amount, false) +
          "</b>" +
          "</li>"
        );
      })
      .join("") +
    "</ol>"
  );
}

function sortRecords(records) {
  const { key, direction } = state.sort;
  const multiplier = direction === "asc" ? 1 : -1;
  return [...records].sort((left, right) => {
    const leftValue = key === "Amount" ? left.Net_Amount : left[key];
    const rightValue = key === "Amount" ? right.Net_Amount : right[key];
    if (typeof leftValue === "number") return (leftValue - rightValue) * multiplier;
    return String(leftValue).localeCompare(String(rightValue)) * multiplier;
  });
}

function sortButton(label, key) {
  const active = state.sort.key === key;
  const icon = active && state.sort.direction === "asc" ? "arrow-up" : "arrow-down";
  return (
    '<button type="button" class="sort-button' +
    (active ? " is-active" : "") +
    '" data-sort="' +
    key +
    '">' +
    label +
    '<i data-lucide="' +
    icon +
    '" aria-hidden="true"></i></button>'
  );
}

function statusPill(status) {
  return '<span class="status-pill status-' + status.toLowerCase() + '">' + escapeHtml(status) + "</span>";
}

function renderTable(records) {
  const sorted = sortRecords(records);
  const pageCount = Math.max(1, Math.ceil(sorted.length / state.pageSize));
  state.page = Math.min(state.page, pageCount);
  const start = (state.page - 1) * state.pageSize;
  const visibleRows = sorted.slice(start, start + state.pageSize);
  const pageLabel = sorted.length ? start + 1 + "–" + Math.min(start + state.pageSize, sorted.length) : "0";

  const rows = visibleRows.length
    ? visibleRows
        .map((record) => {
          return (
            "<tr>" +
            "<td>" +
            formatDate(record.Date, true) +
            "</td>" +
            '<td class="description-cell"><strong>' +
            escapeHtml(record.Description) +
            "</strong><span>" +
            escapeHtml(record.Customer) +
            "</span></td>" +
            "<td>" +
            escapeHtml(record.Category) +
            "</td>" +
            "<td>" +
            escapeHtml(record.Department) +
            "</td>" +
            "<td>" +
            escapeHtml(record.Region) +
            "</td>" +
            "<td>" +
            escapeHtml(record.Transaction_Type) +
            "</td>" +
            '<td class="amount-cell">' +
            formatCurrency(record.Net_Amount, false) +
            "</td>" +
            "<td>" +
            statusPill(record.Status) +
            "</td>" +
            "</tr>"
          );
        })
        .join("")
    : '<tr><td colspan="8"><div class="empty-table"><i data-lucide="filter-x"></i><strong>No financial records match your filters.</strong><span>Try changing a filter or clear the selection.</span></div></td></tr>';

  return (
    '<section class="records-section" aria-labelledby="records-title">' +
    '<div class="section-heading records-heading">' +
    '<div><p class="eyebrow">Transaction register</p><h2 id="records-title">Financial records</h2></div>' +
    '<div class="table-count">' +
    pluralize(records.length, "record") +
    "</div>" +
    "</div>" +
    '<div class="table-shell"><div class="table-scroll"><table>' +
    "<thead><tr>" +
    "<th>" +
    sortButton("Date", "Date") +
    "</th><th>" +
    sortButton("Description", "Description") +
    "</th><th>" +
    sortButton("Category", "Category") +
    "</th><th>" +
    sortButton("Department", "Department") +
    "</th><th>" +
    sortButton("Region", "Region") +
    "</th><th>" +
    sortButton("Type", "Transaction_Type") +
    "</th><th>" +
    sortButton("Amount", "Amount") +
    "</th><th>" +
    sortButton("Status", "Status") +
    "</th></tr></thead>" +
    "<tbody>" +
    rows +
    "</tbody></table></div>" +
    '<div class="pagination"><span>Showing ' +
    pageLabel +
    " of " +
    numberFormat.format(sorted.length) +
    '</span><div class="pagination-controls"><button type="button" class="icon-button" data-page="previous" ' +
    (state.page === 1 ? "disabled" : "") +
    ' aria-label="Previous page" title="Previous page"><i data-lucide="chevron-left"></i></button>' +
    '<span class="page-readout">Page ' +
    state.page +
    " / " +
    pageCount +
    '</span><button type="button" class="icon-button" data-page="next" ' +
    (state.page === pageCount ? "disabled" : "") +
    ' aria-label="Next page" title="Next page"><i data-lucide="chevron-right"></i></button></div></div>' +
    "</div></section>"
  );
}

function renderInsights(insights) {
  const items = insights.length
    ? insights
        .map((insight, index) => {
          return (
            '<article class="insight-item"><span class="insight-index">0' +
            (index + 1) +
            "</span><div><p>" +
            escapeHtml(insight.label) +
            "</p><h3>" +
            escapeHtml(insight.value) +
            "</h3><span>" +
            escapeHtml(insight.detail) +
            "</span></div></article>"
          );
        })
        .join("")
    : '<div class="insights-empty">No calculated insights are available for this filter combination.</div>';

  return (
    '<section class="insights-section" aria-labelledby="insights-title"><div class="section-heading">' +
    '<div><p class="eyebrow">Readout</p><h2 id="insights-title">Data-led observations</h2></div>' +
    '<p class="section-note">Computed from the selected records</p></div><div class="insights-grid">' +
    items +
    "</div></section>"
  );
}

function renderReports(records, metrics) {
  return (
    '<section class="report-section" aria-labelledby="reports-title"><div class="report-heading">' +
    '<div><p class="eyebrow">Reporting desk</p><h2 id="reports-title">Take the selected ledger with you</h2>' +
    '<p>Export files use the active filters and include ' +
    numberFormat.format(records.length) +
    " records.</p></div>" +
    '<div class="report-total"><span>Net profit</span><strong>' +
    formatCurrency(metrics.profit) +
    "</strong></div></div>" +
    '<div class="export-actions"><button type="button" class="export-button" data-export="csv"><i data-lucide="file-down"></i><span>Export CSV</span></button>' +
    '<button type="button" class="export-button" data-export="excel"><i data-lucide="table-properties"></i><span>Export Excel</span></button>' +
    '<button type="button" class="export-button" data-export="pdf"><i data-lucide="file-text"></i><span>Generate PDF</span></button></div>' +
    "</section>"
  );
}

function renderDatasetInfo() {
  return (
    '<footer class="dataset-footer"><div><span class="footer-mark"></span><strong>Ledgerline</strong><span>Financial intelligence</span></div>' +
    '<dl><div><dt>Records</dt><dd>2,400</dd></div><div><dt>Data range</dt><dd>Jan 2025 – Dec 2025</dd></div>' +
    '<div><dt>Last updated</dt><dd>Dec 31, 2025</dd></div><div><dt>Source</dt><dd>Illustrative / Demo Data</dd></div></dl></footer>'
  );
}

function chartPanel(id, eyebrow, title, caption, className = "") {
  return (
    '<article class="chart-panel ' +
    className +
    '">' +
    '<div class="chart-panel-heading"><div><p class="eyebrow">' +
    eyebrow +
    "</p><h3>" +
    title +
    '</h3></div><p>' +
    caption +
    "</p></div><div class=\"chart-canvas\"><canvas id=\"" +
    id +
    '"></canvas></div></article>'
  );
}

function renderDashboard() {
  if (state.loadError) {
    app.innerHTML =
      '<main class="error-view"><div><span class="error-mark">!</span><p class="eyebrow">Data connection</p><h1>Unable to load financial data.</h1><p>Please check the illustrative dataset and try again.</p><button type="button" class="primary-button" data-retry><i data-lucide="refresh-cw"></i>Try again</button></div></main>';
    createIcons({ icons });
    document.querySelector("[data-retry]").addEventListener("click", loadDataset);
    return;
  }

  const records = filteredRecords();
  const metrics = calculateMetrics(records);
  const trend = trendBuckets(records);
  const breakdowns = topBreakdowns(records);
  const insights = meaningfulInsights(records, metrics, breakdowns);
  const noRecords = !records.length;

  app.innerHTML =
    '<main class="dashboard-shell">' +
    '<header class="dashboard-header"><div class="brand-lockup"><div class="brand-mark"><span></span><span></span><span></span></div><div><p class="eyebrow">Ledgerline</p><h1>Financial Overview</h1></div></div>' +
    '<div class="header-meta"><div class="header-period"><span>Data period</span><strong>FY 2025</strong></div><div class="dataset-badge"><span></span>Illustrative / Demo Data</div>' +
    '<label class="global-search"><i data-lucide="search" aria-hidden="true"></i><span class="sr-only">Search records</span><input data-filter="search" type="search" value="' +
    escapeHtml(state.filters.search) +
    '" placeholder="Search records" /></label></div></header>' +
    '<section class="filter-ribbon" aria-label="Financial filters"><div class="filter-ribbon-top"><div><p class="eyebrow">Selection controls</p><p class="filter-intro">Refine this financial view by period, ownership, status, or value.</p></div><button type="button" class="clear-button" data-clear-filters><i data-lucide="rotate-ccw"></i>Clear all</button></div>' +
    '<div class="filter-grid"><label class="control-field date-field"><span>Start date</span><input data-filter="startDate" type="date" value="' +
    state.filters.startDate +
    '" /></label><label class="control-field date-field"><span>End date</span><input data-filter="endDate" type="date" value="' +
    state.filters.endDate +
    '" /></label><label class="control-field"><span>Category</span><select data-filter="category">' +
    selectOptions(FILTER_OPTIONS.category, state.filters.category, "All categories") +
    '</select></label><label class="control-field"><span>Region</span><select data-filter="region">' +
    selectOptions(FILTER_OPTIONS.region, state.filters.region, "All regions") +
    '</select></label><label class="control-field"><span>Department</span><select data-filter="department">' +
    selectOptions(FILTER_OPTIONS.department, state.filters.department, "All departments") +
    '</select></label><label class="control-field"><span>Transaction type</span><select data-filter="transactionType">' +
    selectOptions(FILTER_OPTIONS.transactionType, state.filters.transactionType, "All types") +
    '</select></label><label class="control-field"><span>Status</span><select data-filter="status">' +
    selectOptions(FILTER_OPTIONS.status, state.filters.status, "All statuses") +
    '</select></label><label class="control-field amount-field"><span>Amount range</span><div><input data-filter="minAmount" inputmode="decimal" type="number" min="0" value="' +
    state.filters.minAmount +
    '" placeholder="Min" /><span>to</span><input data-filter="maxAmount" inputmode="decimal" type="number" min="0" value="' +
    state.filters.maxAmount +
    '" placeholder="Max" /></div></label></div>' +
    activeFilterChips() +
    "</section>" +
    renderKpis(metrics) +
    '<section class="main-analysis" aria-labelledby="momentum-title"><div class="section-heading"><div><p class="eyebrow">Financial momentum</p><h2 id="momentum-title">Revenue, expense, and net result</h2></div><div class="segmented-control" role="group" aria-label="Trend granularity">' +
    ["Daily", "Weekly", "Monthly", "Quarterly"]
      .map((grain) => {
        return (
          '<button type="button" data-grain="' +
          grain +
          '" class="' +
          (state.grain === grain ? "is-selected" : "") +
          '" aria-pressed="' +
          String(state.grain === grain) +
          '">' +
          grain +
          "</button>"
        );
      })
      .join("") +
    "</div></div>" +
    '<div class="main-analysis-grid">' +
    chartPanel("trend-chart", "Flow over time", "A moving financial picture", "Hover to inspect a period.", "hero-chart") +
    renderLedgerBalance(metrics) +
    "</div></section>" +
    '<section class="secondary-analysis" aria-labelledby="analysis-title"><div class="section-heading"><div><p class="eyebrow">Composition</p><h2 id="analysis-title">Where money is moving</h2></div><p class="section-note">' +
    (noRecords ? "No records in the selected view" : "Selected ledger breakdowns") +
    "</p></div><div class=\"analysis-grid\">" +
    chartPanel("expense-chart", "Expenses", "Spend by category", "Completed operating expenses") +
    chartPanel("region-chart", "Revenue", "Revenue by region", "Recognized revenue only") +
    chartPanel("profit-chart", "Profit", "Contribution by department", "Revenue less expenses") +
    '<article class="quick-list-panel"><div class="chart-panel-heading"><div><p class="eyebrow">Top ledger entries</p><h3>Largest transactions</h3></div><p>Selected value</p></div>' +
    renderLargestTransactions(records) +
    "</article></div></section>" +
    '<section class="status-section" aria-labelledby="status-title"><div><p class="eyebrow">Settlement health</p><h2 id="status-title">Transaction status mix</h2><p>Live status distribution across the current filtered register.</p></div>' +
    '<div class="status-chart-wrap"><canvas id="status-chart"></canvas></div><div class="status-stat"><strong>' +
    (records.length ? ((metrics.completedCount / records.length) * 100).toFixed(1) : "0.0") +
    '%</strong><span>completed</span></div></section>' +
    renderTable(records) +
    renderInsights(insights) +
    renderReports(records, metrics) +
    renderDatasetInfo() +
    "</main>";

  createIcons({ icons, attrs: { "stroke-width": 1.8 } });
  bindEvents();
  drawCharts({ trend, breakdowns });
}

function commonChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 300 },
    plugins: {
      legend: {
        labels: {
          boxWidth: 9,
          boxHeight: 9,
          usePointStyle: true,
          pointStyle: "circle",
          color: PALETTE.ink,
          font: { size: 11, family: "Arial, sans-serif" },
          padding: 18
        }
      },
      tooltip: {
        backgroundColor: PALETTE.ink,
        titleColor: "#ffffff",
        bodyColor: "#ffffff",
        padding: 10,
        displayColors: true,
        callbacks: {
          label(context) {
            const label = context.dataset.label ? context.dataset.label + ": " : "";
            return label + formatCurrency(context.raw, false);
          }
        }
      }
    }
  };
}

function emptyChartMessage(id, message) {
  const canvas = document.querySelector("#" + id);
  if (!canvas) return;
  const parent = canvas.parentElement;
  parent.innerHTML = '<div class="chart-empty"><i data-lucide="chart-no-axes-column"></i><span>' + message + "</span></div>";
  createIcons({ icons, attrs: { "stroke-width": 1.6 } });
}

function pushChart(chart) {
  state.charts.push(chart);
}

function drawCharts({ trend, breakdowns }) {
  state.charts.forEach((chart) => chart.destroy());
  state.charts = [];

  if (!trend.length) {
    ["trend-chart", "expense-chart", "region-chart", "profit-chart", "status-chart"].forEach((id) =>
      emptyChartMessage(id, "No data for this selection")
    );
    return;
  }

  const mainOptions = commonChartOptions();
  mainOptions.scales = {
    x: {
      grid: { display: false },
      border: { display: false },
      ticks: { color: PALETTE.quiet, font: { size: 10 }, maxRotation: 0 }
    },
    y: {
      grid: { color: "#e3e9e4", drawTicks: false },
      border: { display: false },
      ticks: {
        color: PALETTE.quiet,
        font: { size: 10 },
        callback(value) {
          return formatCurrency(value);
        }
      }
    }
  };
  mainOptions.elements = { line: { tension: 0.35, borderWidth: 2.2 }, point: { radius: 0, hoverRadius: 4 } };

  pushChart(
    new Chart(document.querySelector("#trend-chart"), {
      type: "line",
      data: {
        labels: trend.map((bucket) => bucket.label),
        datasets: [
          {
            label: "Revenue",
            data: trend.map((bucket) => bucket.revenue),
            borderColor: PALETTE.teal,
            backgroundColor: "rgba(0, 107, 103, 0.08)",
            fill: true
          },
          {
            label: "Expenses",
            data: trend.map((bucket) => bucket.expenses),
            borderColor: PALETTE.coral,
            backgroundColor: "rgba(212, 90, 68, 0.03)",
            fill: true
          },
          {
            label: "Net result",
            data: trend.map((bucket) => bucket.profit),
            borderColor: PALETTE.gold,
            backgroundColor: "transparent",
            fill: false,
            borderDash: [5, 4]
          }
        ]
      },
      options: mainOptions
    })
  );

  createBarChart("expense-chart", breakdowns.expenseCategory, "Expenses", PALETTE.coral);
  createBarChart("region-chart", breakdowns.regionRevenue, "Revenue", PALETTE.teal);
  createProfitChart(breakdowns.departmentProfit);
  createStatusChart(breakdowns.status);
}

function createBarChart(id, data, label, color) {
  if (!data.length) {
    emptyChartMessage(id, "No completed values to compare");
    return;
  }

  const options = commonChartOptions();
  options.plugins.legend.display = false;
  options.scales = {
    x: {
      grid: { display: false },
      border: { display: false },
      ticks: { color: PALETTE.quiet, font: { size: 10 }, maxRotation: 0 }
    },
    y: {
      grid: { color: "#e7ece8", drawTicks: false },
      border: { display: false },
      ticks: {
        color: PALETTE.quiet,
        font: { size: 10 },
        callback(value) {
          return formatCurrency(value);
        }
      }
    }
  };

  pushChart(
    new Chart(document.querySelector("#" + id), {
      type: "bar",
      data: {
        labels: data.map((item) => item.label),
        datasets: [
          {
            label,
            data: data.map((item) => item.value),
            backgroundColor: color,
            borderRadius: 3,
            maxBarThickness: 28
          }
        ]
      },
      options
    })
  );
}

function createProfitChart(data) {
  if (!data.length) {
    emptyChartMessage("profit-chart", "No contribution values to compare");
    return;
  }

  const options = commonChartOptions();
  options.indexAxis = "y";
  options.plugins.legend.display = false;
  options.scales = {
    x: {
      grid: { color: "#e7ece8", drawTicks: false },
      border: { display: false },
      ticks: {
        color: PALETTE.quiet,
        font: { size: 10 },
        callback(value) {
          return formatCurrency(value);
        }
      }
    },
    y: {
      grid: { display: false },
      border: { display: false },
      ticks: { color: PALETTE.quiet, font: { size: 10 } }
    }
  };

  pushChart(
    new Chart(document.querySelector("#profit-chart"), {
      type: "bar",
      data: {
        labels: data.map((item) => item.label),
        datasets: [
          {
            label: "Profit contribution",
            data: data.map((item) => item.value),
            backgroundColor: data.map((item) => (item.value >= 0 ? PALETTE.moss : PALETTE.coral)),
            borderRadius: 3,
            maxBarThickness: 22
          }
        ]
      },
      options
    })
  );
}

function createStatusChart(data) {
  if (!data.length) {
    emptyChartMessage("status-chart", "No transaction statuses to show");
    return;
  }

  const options = commonChartOptions();
  options.cutout = "68%";
  options.plugins.legend.position = "right";
  options.plugins.tooltip.callbacks.label = function label(context) {
    return context.label + ": " + numberFormat.format(context.raw) + " records";
  };

  pushChart(
    new Chart(document.querySelector("#status-chart"), {
      type: "doughnut",
      data: {
        labels: data.map((item) => item.label),
        datasets: [
          {
            data: data.map((item) => item.value),
            backgroundColor: [PALETTE.teal, PALETTE.gold, PALETTE.coral, PALETTE.clay],
            borderWidth: 0,
            spacing: 3
          }
        ]
      },
      options
    })
  );
}

function bindEvents() {
  document.querySelectorAll("[data-filter]").forEach((element) => {
    const eventName = element.dataset.filter === "search" ? "input" : "change";
    element.addEventListener(eventName, (event) => {
      const key = event.currentTarget.dataset.filter;
      const value = event.currentTarget.value;
      state.filters[key] = value;
      state.page = 1;
      renderDashboard();

      if (key === "search") {
        requestAnimationFrame(() => {
          const input = document.querySelector('[data-filter="search"]');
          if (input) {
            input.focus();
            input.setSelectionRange(value.length, value.length);
          }
        });
      }
    });
  });

  document.querySelector("[data-clear-filters]").addEventListener("click", () => {
    state.filters = { ...DEFAULT_FILTERS };
    state.page = 1;
    renderDashboard();
  });

  document.querySelectorAll("[data-remove-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      const key = button.dataset.removeFilter;
      if (key === "dates") {
        state.filters.startDate = DEFAULT_FILTERS.startDate;
        state.filters.endDate = DEFAULT_FILTERS.endDate;
      } else {
        state.filters[key] = DEFAULT_FILTERS[key];
      }
      state.page = 1;
      renderDashboard();
    });
  });

  document.querySelectorAll("[data-grain]").forEach((button) => {
    button.addEventListener("click", () => {
      state.grain = button.dataset.grain;
      renderDashboard();
    });
  });

  document.querySelectorAll("[data-sort]").forEach((button) => {
    button.addEventListener("click", () => {
      const key = button.dataset.sort;
      if (state.sort.key === key) {
        state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
      } else {
        state.sort = { key, direction: "asc" };
      }
      renderDashboard();
    });
  });

  document.querySelectorAll("[data-page]").forEach((button) => {
    button.addEventListener("click", () => {
      state.page += button.dataset.page === "next" ? 1 : -1;
      renderDashboard();
      document.querySelector(".records-section").scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });

  document.querySelectorAll("[data-export]").forEach((button) => {
    button.addEventListener("click", () => exportRecords(button.dataset.export));
  });
}

function exportRows(records) {
  return records.map((record) => ({
    Transaction_ID: record.Transaction_ID,
    Date: record.Date,
    Description: record.Description,
    Customer: record.Customer,
    Category: record.Category,
    Subcategory: record.Subcategory,
    Department: record.Department,
    Region: record.Region,
    Transaction_Type: record.Transaction_Type,
    Status: record.Status,
    Amount: record.Amount,
    Tax: record.Tax,
    Net_Amount: record.Net_Amount,
    Revenue: record.Revenue,
    Expense: record.Expense,
    Profit: record.Profit,
    Currency: record.Currency
  }));
}

function downloadBlob(content, type, name) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function exportRecords(format) {
  const records = filteredRecords();
  const rows = exportRows(records);
  const stamp = state.filters.startDate + "_to_" + state.filters.endDate;

  if (format === "csv") {
    const headers = Object.keys(rows[0] || { Transaction_ID: "" });
    const escapeCell = (value) => '"' + String(value ?? "").replace(/"/g, '""') + '"';
    const csv = [headers.join(",")]
      .concat(rows.map((row) => headers.map((header) => escapeCell(row[header])).join(",")))
      .join("\n");
    downloadBlob(csv, "text/csv;charset=utf-8", "ledgerline_" + stamp + ".csv");
    return;
  }

  if (format === "excel") {
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet["!cols"] = Object.keys(rows[0] || { Transaction_ID: "" }).map((key) => ({ wch: Math.max(13, key.length + 2) }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Selected ledger");
    XLSX.writeFile(workbook, "ledgerline_" + stamp + ".xlsx");
    return;
  }

  const metrics = calculateMetrics(records);
  const document = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
  document.setTextColor(23, 49, 43);
  document.setFontSize(20);
  document.text("Ledgerline financial report", 40, 42);
  document.setFontSize(9);
  document.setTextColor(90, 106, 99);
  document.text(
    "Illustrative / Demo Data | " + formatPeriod(state.filters.startDate) + " to " + formatPeriod(state.filters.endDate),
    40,
    60
  );
  document.setFontSize(11);
  document.setTextColor(23, 49, 43);
  document.text(
    "Revenue " + formatCurrency(metrics.revenue, false) + "     Expenses " + formatCurrency(metrics.expenses, false) + "     Net profit " + formatCurrency(metrics.profit, false),
    40,
    84
  );
  autoTable(document, {
    startY: 105,
    head: [["Date", "Description", "Category", "Department", "Region", "Type", "Amount", "Status"]],
    body: rows.slice(0, 45).map((row) => [
      row.Date,
      row.Description,
      row.Category,
      row.Department,
      row.Region,
      row.Transaction_Type,
      formatCurrency(row.Net_Amount, false),
      row.Status
    ]),
    styles: { fontSize: 7, cellPadding: 4, textColor: [35, 52, 46] },
    headStyles: { fillColor: [23, 49, 43], textColor: [255, 255, 255] },
    alternateRowStyles: { fillColor: [242, 246, 243] }
  });
  document.save("ledgerline_" + stamp + ".pdf");
}

async function loadDataset() {
  state.loadError = false;
  app.innerHTML =
    '<div class="loading-shell" aria-live="polite"><div class="loading-brand">Ledgerline</div><div class="loading-copy">Preparing your financial overview</div><div class="loading-bar"><span></span></div></div>';

  try {
    const response = await fetch(DATA_URL);
    if (!response.ok) throw new Error("Dataset request failed");
    state.records = await response.json();
    renderDashboard();
  } catch (error) {
    state.loadError = true;
    renderDashboard();
  }
}

loadDataset();
