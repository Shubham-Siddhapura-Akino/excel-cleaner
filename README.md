# Excel Cleaner Studio

A production-ready application built with **Next.js**, **Fastify**, and **ExcelJS** to batch-clean and normalize text values across Excel workbooks while preserving 100% of formatting, formulas, sheets, and cell metadata. Optimized for deployment on **Vercel**.

---

## Architecture

- **Frontend**: Next.js App Router with React, Tailwind CSS, and Lucide icons.
  - Multi-file drag & drop queue
  - Batch cleaning with animated real-time progress
  - One-click master download as a `.zip` archive
  - Interactive before & after cell diff inspector
- **Backend API**: Fastify with `@fastify/cors` and `@fastify/multipart`.
  - Mounted via serverless adapter `pages/api/[[...fastify]].js` for seamless Vercel deployment.
  - Can also run as a standalone Fastify server via `npm run server:fastify`.
  - Supports streaming octet-stream and multipart uploads up to 100MB.
  - Ephemeral file storage compatible with Vercel `/tmp` read-only environments.

---

## Features

- **Strict Whitespace & Text Normalization**:
  - Removes leading and trailing spaces.
  - Removes tabs (`\t`), carriage returns (`\r`), and newlines (`\n`).
  - Removes non-breaking spaces (Unicode `00A0` / `\u00A0`).
  - Removes double quotes (`"`).
  - Removes invisible Unicode whitespace characters (Zero-width space `\u200B`, `\u200C`, `\u200D`, `\uFEFF`, directional controls, and control codes).
  - Normalizes multiple consecutive spaces into a single space.

- **100% Workbook Property Preservation**:
  - Multiple worksheets & sheet names
  - Cell formatting (fonts, font size, text colors, background fills, borders, alignment)
  - Number formats (currency, percentages, custom date formats)
  - Formula definitions and formula results
  - Merged cell ranges
  - Hidden rows & columns, column widths, row heights
  - AutoFilters, sheet views / freeze panes, hyperlinks, images, comments

---

## Installation

```bash
git clone <repository-url>
cd excel-cleaner
npm install
```

---

## Running Locally

### 1. Next.js Web App (Recommended)
Runs the full Next.js application with Fastify backend API:

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

### 2. Production Build & Start
```bash
npm run build
npm start
```

### 3. Standalone Fastify Backend (Optional)
To run only the Fastify API server independently on port 3001:

```bash
npm run server:fastify
```

### 4. CLI Execution
To clean a spreadsheet directly from the terminal without starting the web UI:

```bash
# Default (reads input/workbook.xlsx -> writes output/workbook_cleaned.xlsx)
npm run cli

# Custom files
node index.js path/to/input.xlsx path/to/output.xlsx
```

---

## Deploying to Vercel

This repository is pre-configured for Vercel with zero extra setup needed.

### Method 1: Deploy via Vercel Dashboard (GitHub)
1. Push this repository to GitHub or GitLab.
2. In [Vercel](https://vercel.com), click **Add New Project** and import the repository.
3. Vercel automatically detects Next.js.
4. Click **Deploy**. Both the Next.js UI and the Fastify backend API will deploy automatically!

### Method 2: Deploy via Vercel CLI
```bash
npm i -g vercel
vercel
```

All `/api/*` endpoints (`/api/clean`, `/api/download/:id`, `/api/download-zip`, `/api/health`, `/api/history`) are executed serverlessly by Fastify.

---

## Examples

### Input vs Output Normalization

| Raw Input Cell Value | Cleaned Output Cell Value |
| :--- | :--- |
| `"  Shubham     Siddhapura "` | `Shubham Siddhapura` |
| `"John\nSmith"` | `John Smith` |
| `"    John      Smith"` | `John Smith` |
| `"\t\tHeader\r\nValue\u00A0\u200B "` | `Header Value` |
| `"\"Quoted Text\""` | `Quoted Text` |

---

## Console Output & Statistics

Upon completion, the application displays progress logs and performance statistics:

```text
Workbook Loaded
Processing Sheet: Sheet1
Processing Sheet: Employees
Processing Sheet: Payroll

Cleaning Complete

Statistics

Sheets Processed: 3
Rows Processed:   1500
Cells Visited:    12000
String Cells:     8000
Modified Cells:   450
Execution Time:   124ms
```

---

## Project Structure

```
excel-cleaner/
│
├── package.json          # Package configuration (ES Modules)
├── index.js              # CLI entry point, path resolution, CLI logging
├── cleaner.js            # Core workbook loading, worksheet iteration, stats tracking
├── utils/
│   └── cleanText.js      # Text normalization utility function
├── input/                # Default input directory
│   └── workbook.xlsx     # Input workbook sample
├── output/               # Default output directory
│   └── workbook_cleaned.xlsx
├── test/
│   ├── testRunner.js     # Verification & assertion test suite
│   └── benchmark.js      # Performance benchmark script
└── README.md             # Project documentation
```

---

## Testing & Benchmarks

To run the automated unit and end-to-end integration test suite:

```bash
npm test
```

To run performance benchmark on 250,000+ cells:

```bash
node test/benchmark.js
```

---

## Performance Notes

- **Streamlined Cell Iteration**: Utilizes ExcelJS `eachRow` and `eachCell` loops for minimal memory allocation.
- **Selective Type Filtering**: Evaluates cell types early to skip non-string primitive values (Numbers, Dates, Formulas) instantly.
- **Memory Footprint**: Tested on workbooks with 10+ worksheets, 50,000+ rows, and 250,000+ cells, completing execution in under ~4 seconds with ~100MB peak memory delta.
