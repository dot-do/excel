# excel.do

> Spreadsheets at the Edge. Talk to Your Data. AI-First.

Microsoft charges $12.50/user/month for Excel. Google locks you into Sheets. Both require you to click through menus, remember functions, and manually format cells. Spreadsheets should be as simple as describing what you want.

**excel.do** is the AI-native spreadsheet engine. Edge-first. Natural language. Zero configuration.

## AI-Native API

```typescript
import { excel } from 'excel.do'           // Full SDK
import { excel } from 'excel.do/tiny'      // Minimal client
import { excel } from 'excel.do/formulas'  // Formula-only operations
```

Natural language for spreadsheet workflows:

```typescript
import { excel } from 'excel.do'

// Talk to it like a colleague
const report = await excel`quarterly sales by region`
const summary = await excel`top 10 customers by revenue`
const forecast = await excel`project next quarter based on trends`

// Chain like sentences
await excel`load sales_data.xlsx`
  .map(sheet => excel`pivot by product category`)
  .map(pivot => excel`add percentage of total column`)
  .map(report => excel`export as PDF with charts`)

// Formatting that describes itself
await excel`format ${data} as financial report`
  .bold(`headers`)
  .currency(`revenue columns`)
  .highlight(`values over 1M in green`)
```

## The Problem

Spreadsheet software is stuck in the 1990s:

| What They Charge | The Reality |
|------------------|-------------|
| **Excel 365** | $12.50/user/month, desktop-first |
| **Google Sheets** | Free-ish, but your data is the product |
| **Airtable** | $20/seat/month for basic features |
| **Formula Learning Curve** | Hours memorizing VLOOKUP, INDEX/MATCH, SUMIFS |
| **Collaboration** | Merge conflicts, version chaos |
| **Integration** | Export/import cycles, broken links |

### The Complexity Tax

- **200+ Excel functions** to memorize
- **Nested formulas** that become unreadable
- **Manual formatting** click by click
- **No version control** - "Final_v2_REAL_final.xlsx"
- **No API access** without expensive add-ons

### The Data Trap

- Files scattered across drives and emails
- No real-time collaboration without cloud lock-in
- Formulas break when data moves
- Charts disconnect from sources

## The Solution

**excel.do** reimagines spreadsheets for the AI era:

```
Legacy Spreadsheets              excel.do
-----------------------------------------------------------------
Memorize 200+ functions          Just describe what you want
Click through format menus       Natural language styling
Manual pivot table creation      AI generates pivots
Export/import cycles             API-first, real-time
Desktop or cloud lock-in         Edge-native, runs anywhere
No version control               Git-native via gitx.do
$12.50/user/month                Deploy your own
```

## One-Click Deploy

```bash
npx create-dotdo excel
```

A full spreadsheet engine. Running on infrastructure you control. AI-native from day one.

```typescript
import { Excel } from 'excel.do'

export default Excel({
  name: 'company-sheets',
  domain: 'sheets.mycompany.com',
  storage: {
    hot: 'sqlite',      // Fast access
    warm: 'r2',         // Archive
  },
})
```

## Features

### Loading Data

```typescript
// Load anything
const data = await excel`sales_data.xlsx`
const csv = await excel`customers.csv`
const api = await excel`fetch from CRM API`

// AI infers structure
await excel`sales_data.xlsx`           // returns workbook
await excel`sales_data.xlsx Sheet2`    // specific sheet
await excel`first 100 rows of sales`   // filtered
```

### Querying Data

```typescript
// Natural queries
const results = await excel`orders over $10,000`
const monthly = await excel`group by month, sum revenue`
const top = await excel`top 5 products by units sold`

// AI understands context
await excel`show me Q4`                 // knows it's dates
await excel`find duplicates`            // identifies matching rows
await excel`rows with missing data`     // spots nulls
```

### Formulas

```typescript
// AI writes the formula
await excel`sum column A`
await excel`average of sales where region is West`
await excel`running total of revenue by date`

// Or describe complex logic
await excel`commission: 5% up to 100k, 7% from 100k to 500k, 10% above`
```

### Pivots and Aggregation

```typescript
// Pivots in plain English
await excel`pivot sales by region and quarter`
await excel`breakdown by category with subtotals`
await excel`cross-tab product vs channel`

// Aggregations just work
await excel`total, average, and count by department`
```

### Formatting

```typescript
// Describe the style
await excel`format as financial report`
await excel`bold headers, alternate row colors`
await excel`highlight negative values in red`

// Conditional formatting
await excel`green if above target, red if below`
await excel`data bars for percentages`
await excel`color scale from white to blue by value`
```

### Charts

```typescript
// Charts from description
await excel`bar chart of sales by region`
await excel`line chart showing monthly trends`
await excel`pie chart of market share`

// Compound visualizations
await excel`dashboard with revenue trends and top products`
```

### Export

```typescript
// Export naturally
await excel`save as quarterly_report.xlsx`
await excel`export to PDF`
await excel`send to Google Sheets`

// Formatted exports
await excel`export as PDF with company letterhead`
await excel`CSV with European number format`
```

## Real-Time Collaboration

```typescript
// Live editing powered by gitx.do
await excel`open sales_forecast.xlsx`
  .share(`sales-team@company.com`)
  .track()   // version history automatic

// Merge like git
await excel`merge Sarah's changes with mine`

// Conflict resolution
await excel`accept their revenue numbers, keep my formulas`
```

## Promise Pipelining

Chain operations without waiting:

```typescript
// One network round trip for the entire chain
await excel`quarterly_data.xlsx`
  .map(data => excel`filter to current year`)
  .map(filtered => excel`pivot by product line`)
  .map(pivot => excel`add year-over-year growth`)
  .map(analysis => excel`format as executive summary`)
  .map(report => excel`export as PDF`)

// Parallel processing
const [sales, costs, inventory] = await Promise.all([
  excel`sum revenue by region`,
  excel`sum expenses by category`,
  excel`inventory levels by warehouse`,
])

await excel`combine ${sales}, ${costs}, ${inventory} into P&L`
```

## AI-Native Features

### Smart Suggestions

```typescript
// AI analyzes your data and suggests
await excel`analyze this dataset`
// Returns: insights, recommended charts, potential issues

await excel`what formulas would help here?`
// Returns: relevant formulas for your data patterns
```

### Natural Transformations

```typescript
// Describe the transformation
await excel`split full name into first and last`
await excel`convert dates to fiscal quarters`
await excel`normalize phone numbers`
await excel`extract domains from emails`
```

### Anomaly Detection

```typescript
// AI spots problems
await excel`find outliers in revenue`
await excel`flag suspicious transactions`
await excel`data quality issues`
```

### Forecasting

```typescript
// Built-in ML
await excel`forecast next 6 months`
await excel`predict Q4 based on historical trends`
await excel`what-if: 10% price increase`
```

## Architecture

### Edge-Native Design

```
Client Request --> Cloudflare Edge --> Durable Object --> SQLite
                        |                    |              |
                   Global CDN           Per-Workbook     Hot Data
                                        Isolation        (<10ms)
                                             |
                                            R2
                                             |
                                        Cold Storage
                                        (< 100ms)
```

### Durable Object per Workbook

```
WorkbookDO (metadata, users, sharing)
  |
  +-- SheetsDO (cells, formulas, formatting)
  |     |-- SQLite: Cell data (encrypted)
  |     +-- R2: Large attachments, images
  |
  +-- CollabDO (real-time cursors, selections)
  |     +-- WebSocket connections
  |
  +-- HistoryDO (version control via gitx.do)
        +-- SQLite: Commits, branches
```

### Storage Tiers

| Tier | Storage | Use Case | Query Speed |
|------|---------|----------|-------------|
| **Hot** | SQLite | Active cells, recent edits | <10ms |
| **Warm** | R2 + Index | Historical versions, large files | <100ms |
| **Cold** | R2 Archive | Old versions, compliance | <1s |

## vs Traditional Spreadsheets

| Feature | Excel/Sheets | excel.do |
|---------|--------------|----------|
| **Interface** | Click menus | Natural language |
| **Formulas** | Memorize syntax | Describe intent |
| **Collaboration** | Merge conflicts | Git-native |
| **API Access** | Add-ons required | API-first |
| **Deployment** | Cloud lock-in | Your infrastructure |
| **AI** | Copilot add-on | Built-in from day one |
| **Version Control** | Manual saves | Automatic history |
| **Cost** | Per-user fees | Run your own |

## Use Cases

### Financial Reporting

```typescript
await excel`load GL_export.csv`
  .map(data => excel`map to chart of accounts`)
  .map(mapped => excel`generate trial balance`)
  .map(tb => excel`create P&L and Balance Sheet`)
  .map(statements => excel`format for board presentation`)
```

### Data Cleaning

```typescript
await excel`customer_list.xlsx`
  .map(data => excel`deduplicate by email`)
  .map(clean => excel`standardize addresses`)
  .map(std => excel`validate phone numbers`)
  .map(valid => excel`flag incomplete records`)
```

### Sales Analytics

```typescript
await excel`CRM export this quarter`
  .map(deals => excel`win rate by rep`)
  .map(rates => excel`pipeline velocity`)
  .map(velocity => excel`forecast accuracy`)
  .map(analysis => excel`sales dashboard`)
```

### Inventory Management

```typescript
await excel`warehouse_data.xlsx`
  .map(inv => excel`ABC analysis`)
  .map(abc => excel`reorder point calculations`)
  .map(reorder => excel`weeks of supply by SKU`)
  .map(wos => excel`stockout risk report`)
```

## Module Exports

```typescript
import { excel } from 'excel.do'              // Full SDK
import { excel } from 'excel.do/tiny'         // Minimal client (< 5KB)
import { excel } from 'excel.do/formulas'     // Formula-only operations
import { excel } from 'excel.do/import'       // XLSX/CSV import
import { excel } from 'excel.do/export'       // XLSX/CSV/PDF export
```

## Deployment Options

### Cloudflare Workers (Recommended)

```bash
npx create-dotdo excel
# Deploys in under 60 seconds
```

### Private Cloud

```bash
docker run -p 8787:8787 dotdo/excel
```

### On-Premises

```bash
./excel-do-install.sh --enterprise --sso
```

## Roadmap

### Core Engine
- [x] Cell data model with formatting
- [x] Formula parsing and evaluation
- [x] XLSX import/export
- [x] CSV import/export
- [x] A1 and R1C1 notation
- [x] Range operations
- [ ] Named ranges
- [ ] Conditional formatting engine
- [ ] Pivot table generation

### AI Features
- [x] Natural language queries
- [x] Smart suggestions
- [x] Data transformations
- [ ] Anomaly detection
- [ ] Forecasting with ML
- [ ] Chart recommendations

### Collaboration
- [x] Real-time editing via gitx.do
- [x] Version history
- [ ] Comments and annotations
- [ ] Conflict resolution UI
- [ ] Sharing permissions

### Integrations
- [x] mongo.do storage backend
- [x] gitx.do version control
- [ ] Google Sheets sync
- [ ] Airtable import
- [ ] API connectors (CRM, ERP)

## Contributing

excel.do is open source under the MIT license.

```bash
git clone https://github.com/dotdo/excel.do
cd excel.do
pnpm install
pnpm test
```

## License

MIT License - Spreadsheets for everyone.

---

<p align="center">
  <strong>Stop clicking. Start describing.</strong>
  <br />
  Edge-native. AI-first. Natural language.
  <br /><br />
  <a href="https://excel.do">Website</a> |
  <a href="https://docs.excel.do">Docs</a> |
  <a href="https://discord.gg/dotdo">Discord</a> |
  <a href="https://github.com/dotdo/excel.do">GitHub</a>
</p>
