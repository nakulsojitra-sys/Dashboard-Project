import { mkdir, writeFile } from "node:fs/promises";

const RECORD_COUNT = 2400;
const OUTPUT_DIRECTORY = new URL("../public/data/", import.meta.url);
const OUTPUT_FILE = new URL("financial_records.json", OUTPUT_DIRECTORY);

let seed = 20251003;

function random() {
  seed = (seed * 1664525 + 1013904223) % 4294967296;
  return seed / 4294967296;
}

function pick(items) {
  return items[Math.floor(random() * items.length)];
}

function weightedPick(items) {
  const total = items.reduce((sum, item) => sum + item.weight, 0);
  let cursor = random() * total;

  for (const item of items) {
    cursor -= item.weight;
    if (cursor <= 0) return item.value;
  }

  return items[items.length - 1].value;
}

function money(value) {
  return Math.round(value * 100) / 100;
}

function isoDate(offset) {
  const date = new Date(Date.UTC(2025, 0, 1 + offset));
  return date.toISOString().slice(0, 10);
}

function monthLabel(date) {
  return new Intl.DateTimeFormat("en-US", { month: "long" }).format(
    new Date(date + "T00:00:00Z")
  );
}

function quarter(date) {
  const month = Number(date.slice(5, 7));
  return "Q" + Math.ceil(month / 3);
}

const regions = {
  North: ["Aster & Co.", "Cedar Health", "Northline Stores", "Ridgeway Group"],
  South: ["Meridian Hotels", "Arbor Foods", "Horizon Labs", "Harbor Retail"],
  East: ["KiteWorks", "Mariner Systems", "Lumen Partners", "Orchid Health"],
  West: ["Solace Studio", "Canyon Retail", "Vantage Media", "Westfield Works"],
  International: ["Nori Global", "Atlas Partners", "Kestrel Group", "Mosaic Labs"]
};

const businessUnits = {
  Finance: "Corporate",
  Marketing: "Growth",
  Operations: "Core Operations",
  HR: "People",
  Technology: "Platform",
  Sales: "Commercial"
};

const incomeProfiles = [
  { subcategory: "Platform subscription", description: "Annual platform subscription", department: "Sales", min: 4800, max: 18800 },
  { subcategory: "Advisory services", description: "Financial advisory engagement", department: "Finance", min: 9200, max: 42000 },
  { subcategory: "Implementation", description: "Implementation milestone", department: "Technology", min: 6500, max: 28500 },
  { subcategory: "Enterprise license", description: "Enterprise license renewal", department: "Sales", min: 18000, max: 68000 },
  { subcategory: "Training", description: "Onsite enablement programme", department: "Operations", min: 3100, max: 12200 }
];

const expenseProfiles = [
  { subcategory: "Cloud infrastructure", description: "Cloud infrastructure services", department: "Technology", min: 2400, max: 16400 },
  { subcategory: "Campaign media", description: "Campaign media placement", department: "Marketing", min: 1800, max: 14200 },
  { subcategory: "Professional services", description: "External professional services", department: "Finance", min: 2200, max: 18800 },
  { subcategory: "Team development", description: "Professional development programme", department: "HR", min: 900, max: 6900 },
  { subcategory: "Facilities", description: "Facilities and workplace services", department: "Operations", min: 1300, max: 8600 },
  { subcategory: "Client travel", description: "Client visit and travel", department: "Sales", min: 700, max: 5900 }
];

const investmentProfiles = [
  { subcategory: "Equipment", description: "Equipment investment", department: "Technology", min: 4800, max: 24500 },
  { subcategory: "Product research", description: "Product research investment", department: "Technology", min: 6500, max: 32700 },
  { subcategory: "Market expansion", description: "Market expansion investment", department: "Sales", min: 7600, max: 38900 }
];

function createRecord(index) {
  const date = isoDate(Math.floor(random() * 365));
  const region = pick(Object.keys(regions));
  const transactionType = weightedPick([
    { value: "Income", weight: 53 },
    { value: "Expense", weight: 34 },
    { value: "Investment", weight: 8 },
    { value: "Transfer", weight: 5 }
  ]);

  let category;
  let profile;
  let amount;
  let taxRate;
  let paymentMethod;

  if (transactionType === "Income") {
    profile = pick(incomeProfiles);
    category = "Revenue";
    amount = money(profile.min + random() * (profile.max - profile.min));
    taxRate = region === "International" ? 0 : 0.08;
    paymentMethod = pick(["Bank transfer", "ACH", "Card"]);
  } else if (transactionType === "Expense") {
    profile = pick(expenseProfiles);
    category = random() < 0.72 ? "Expense" : "Operations";
    amount = money(profile.min + random() * (profile.max - profile.min));
    taxRate = 0.06;
    paymentMethod = pick(["Corporate card", "Bank transfer", "ACH"]);
  } else if (transactionType === "Investment") {
    profile = pick(investmentProfiles);
    category = "Investment";
    amount = money(profile.min + random() * (profile.max - profile.min));
    taxRate = 0.04;
    paymentMethod = "Bank transfer";
  } else {
    profile = {
      subcategory: "Internal allocation",
      description: "Internal business-unit allocation",
      department: pick(Object.keys(businessUnits))
    };
    category = "Other";
    amount = money(1500 + random() * 10800);
    taxRate = 0;
    paymentMethod = "Internal ledger";
  }

  const status = weightedPick([
    { value: "Completed", weight: 88 },
    { value: "Pending", weight: 7 },
    { value: "Failed", weight: 3 },
    { value: "Refunded", weight: 2 }
  ]);
  const tax = money(amount * taxRate);
  const netAmount = money(amount + tax);
  const isCompleted = status === "Completed";
  const revenue = transactionType === "Income" && isCompleted ? netAmount : 0;
  const expense = transactionType === "Expense" && isCompleted ? netAmount : 0;

  return {
    Transaction_ID: "LL-2025-" + String(index + 1).padStart(5, "0"),
    Date: date,
    Account: transactionType === "Income" ? "Operating revenue" : "Operating account",
    Transaction_Type: transactionType,
    Category: category,
    Subcategory: profile.subcategory,
    Description: profile.description,
    Customer: pick(regions[region]),
    Department: profile.department,
    Region: region,
    Payment_Method: paymentMethod,
    Currency: "USD",
    Amount: amount,
    Tax: tax,
    Net_Amount: netAmount,
    Status: status,
    Budget: money(amount * (transactionType === "Income" ? 0.96 : 1.14)),
    Revenue: revenue,
    Expense: expense,
    Profit: money(revenue - expense),
    Quarter: quarter(date),
    Month: monthLabel(date),
    Year: 2025,
    Business_Unit: businessUnits[profile.department]
  };
}

const records = Array.from({ length: RECORD_COUNT }, (_, index) => createRecord(index))
  .sort((left, right) => left.Date.localeCompare(right.Date))
  .map((record, index) => ({
    ...record,
    Transaction_ID: "LL-2025-" + String(index + 1).padStart(5, "0")
  }));

await mkdir(OUTPUT_DIRECTORY, { recursive: true });
await writeFile(OUTPUT_FILE, JSON.stringify(records, null, 2) + "\n", "utf8");

console.log("Generated " + records.length + " illustrative financial records.");
