# excel.do

**Spreadsheets at the Edge** - A mongo.do + gitx.do powered spreadsheet engine for Cloudflare Workers.

excel.do is a modern spreadsheet engine designed to run at the edge, providing full Excel compatibility with support for formulas, formatting, import/export, and collaborative features powered by mongo.do and gitx.do.

## Features

- **Edge-Native**: Built for Cloudflare Workers with zero cold starts
- **Excel Compatible**: Full XLSX import/export via SheetJS
- **Formula Engine**: Parse and evaluate Excel formulas with AST-based parsing
- **Cell Management**: Complete cell data model with formatting, validation, and metadata
- **Address Handling**: Support for A1 and R1C1 notation, ranges, and cross-sheet references
- **Storage Backends**: Multiple storage options including mongo.do and custom database implementations
- **Real-Time Collaboration**: WebSocket support for live updates powered by gitx.do
- **Type-Safe**: Full TypeScript support with comprehensive type definitions

## Installation

### npm

```bash
npm install excel.do
```

### With peer dependencies

If you want to use storage features with mongo.do or gitx.do:

```bash
npm install excel.do mongo.do gitx.do
```

## Quick Start

### Basic Usage

```typescript
import { createCell, createCellValue, colToIndex } from 'excel.do'

// Create a cell with a value
const cell = createCell({
  sheet: 'Sheet1',
  row: 1,
  col: 'A',
  value: 'Hello, World!',
})

console.log(cell)
// {
//   _id: 'Sheet1!A1',
//   sheet: 'Sheet1',
//   row: 1,
//   col: 'A',
//   colIndex: 0,
//   value: { v: 'Hello, World!', t: 'string' },
//   createdAt: Date,
//   updatedAt: Date
// }
```

### Working with Cell Addresses

```typescript
import { parseA1, colToIndex, indexToCol, toA1 } from 'excel.do/cell'

// Parse A1-style references
const ref = parseA1('Sheet1!$A$1')
// { original: 'Sheet1!$A$1', col: 'A', colIndex: 0, row: 1, sheet: 'Sheet1', absolute: { col: true, row: true } }

// Convert column letters to indices and back
const index = colToIndex('A')     // 0
const col = indexToCol(0)          // 'A'
const col2 = indexToCol(26)        // 'AA'

// Generate A1 notation
const a1 = toA1({ row: 5, col: 3 })
// 'C5'

const a1Abs = toA1({ row: 5, col: 3 }, { absolute: { col: true, row: true } })
// '$C$5'
```

### Range Operations

```typescript
import { parseRange, expandRange, rangeContains, rangeIntersects } from 'excel.do/cell'

// Parse ranges
const range = parseRange('A1:C10')
// {
//   original: 'A1:C10',
//   start: { col: 'A', colIndex: 0, row: 1, ... },
//   end: { col: 'C', colIndex: 2, row: 10, ... }
// }

// Expand range to individual cells
const cells = expandRange('A1:B2')
// ['A1', 'A2', 'B1', 'B2']

// Check if cell is in range
rangeContains('A1:C10', 'B5')  // true
rangeContains('A1:C10', 'D5')  // false

// Check range intersection
rangeIntersects('A1:C10', 'B5:D15')  // true
rangeIntersects('A1:C10', 'E1:F10')  // false
```

### Cell Values and Types

```typescript
import { createCellValue, inferValueType, formatCellValue } from 'excel.do/cell'

// Create typed cell values
const numValue = createCellValue(42)
// { v: 42, t: 'number' }

const dateValue = createCellValue(new Date('2024-01-15'))
// { v: Date('2024-01-15'), t: 'date' }

const formulaValue = createCellValue(100, '=A1*2')
// { v: 100, t: 'number', f: '=A1*2' }

const errorValue = createCellValue('#DIV/0!')
// { v: '#DIV/0!', t: 'error', e: '#DIV/0!' }

// Infer value types
inferValueType(42)      // 'number'
inferValueType(true)    // 'boolean'
inferValueType(new Date()) // 'date'
inferValueType('#N/A')  // 'error'

// Format values for display
formatCellValue({ v: 42, t: 'number' }, '0%')
// '4200%'

formatCellValue({ v: 42.50, t: 'number' }, '$#,##0.00')
// '$42.50'

formatCellValue({ v: new Date('2024-01-15'), t: 'date' }, 'yyyy-mm-dd')
// '2024-01-15'
```

### Formula Parsing

```typescript
import { parseFormula } from 'excel.do/formula'

// Parse Excel formulas into AST
const result = parseFormula('=SUM(A1:A10) + B1 * 2')
// {
//   ast: {
//     type: 'Program',
//     body: {
//       type: 'BinaryExpression',
//       operator: '+',
//       left: {
//         type: 'FunctionCall',
//         name: 'SUM',
//         arguments: [{ type: 'RangeReference', ... }]
//       },
//       right: { type: 'BinaryExpression', ... }
//     }
//   },
//   errors: [],
//   formula: '=SUM(A1:A10) + B1 * 2'
// }

// Check for parse errors
if (result.errors.length > 0) {
  console.error('Formula parse errors:', result.errors)
}
```

### Cell Formatting

```typescript
import { createCell, mergeCellFormat } from 'excel.do/cell'

// Create a cell with formatting
const cell = createCell({
  sheet: 'Sheet1',
  row: 1,
  col: 'A',
  value: 42.50,
  format: {
    font: {
      bold: true,
      size: 12,
      color: '#FF0000',
    },
    fill: {
      type: 'solid',
      color: '#FFFF00',
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle',
    },
    numberFormat: '$#,##0.00',
  },
})

// Merge format definitions
const base = { font: { bold: true, size: 12 } }
const override = { font: { color: '#FF0000' } }
const merged = mergeCellFormat(base, override)
// { font: { bold: true, size: 12, color: '#FF0000' } }
```

### Cell Cloning and Manipulation

```typescript
import { cloneCell, validateCell, isEmptyCell } from 'excel.do/cell'

const original = createCell({
  sheet: 'Sheet1',
  row: 1,
  col: 'A',
  value: 'Test',
})

// Clone a cell to a new location
const cloned = cloneCell(original, { row: 5, col: 'C' })
// Creates a new cell at Sheet1!C5 with same value but updated timestamps

// Validate cell data
const validation = validateCell(cell)
// { valid: true, errors: [] }

// Check if cell is empty
isEmptyCell(cell)  // false
```

### Import/Export XLSX

```typescript
import { parseWorkbook, parseSheet, rowsFromSheet } from 'excel.do/import'
import { toXLSX, toCSV, toWorkbook, toSheet, cellToSheetJS } from 'excel.do/export'

// Import and parse XLSX file
const buffer = await fs.readFile('data.xlsx')
const workbook = parseWorkbook(buffer)

// Export to XLSX file
const output = await toXLSX({
  sheets: workbook.sheets.map(sheet => ({
    name: sheet.name,
    cells: sheet.cells
  }))
})
await fs.writeFile('output.xlsx', output)

// Export sheet to CSV
const csv = toCSV({
  name: 'Sheet1',
  cells: workbook.sheets[0].cells
})
await fs.writeFile('output.csv', csv)

// Extract rows as documents
const rows = rowsFromSheet(worksheet, { headers: true })
```

## API Reference

### Cell Module (`excel.do/cell`)

#### Functions

- **`createCell(options: CreateCellOptions): Cell`**
  - Create a new cell with values, formatting, and metadata
  - Options: `sheet`, `row`, `col`, `value`, `formula`, `format`, `metadata`

- **`createCellValue(value: CellPrimitive, formula?: string): CellValue`**
  - Create a typed cell value from a primitive
  - Automatically infers type (string, number, boolean, date, error, empty)

- **`inferValueType(value: unknown): CellValueType`**
  - Determine the type of a cell value
  - Returns: 'boolean' | 'number' | 'string' | 'date' | 'error' | 'empty'

- **`getCellId(sheet: string, col: string, row: number): string`**
  - Generate a cell ID in format "Sheet!A1"
  - Handles special characters in sheet names

- **`cloneCell(cell: Cell, overrides?: Partial<CreateCellOptions>): Cell`**
  - Create a deep copy of a cell with optional position/value overrides

- **`validateCell(cell: Cell): ValidationResult`**
  - Validate cell data structure
  - Returns: `{ valid: boolean, errors: string[] }`

- **`isEmptyCell(cell: Cell): boolean`**
  - Check if a cell is empty (null, empty string, or no formula)

- **`formatCellValue(value: CellValue, numberFormat?: string): string`**
  - Format a cell value as a string for display
  - Supports formats: `0%`, `$#,##0.00`, `yyyy-mm-dd`

- **`mergeCellFormat(base?: CellFormat, override?: CellFormat): CellFormat`**
  - Deep merge two cell format objects

#### Address Functions

- **`colToIndex(col: string): number`**
  - Convert column letters to 0-based index
  - Example: 'A' → 0, 'Z' → 25, 'AA' → 26

- **`indexToCol(index: number): string`**
  - Convert 0-based index to column letters
  - Example: 0 → 'A', 25 → 'Z', 26 → 'AA'

- **`parseA1(ref: string): CellReference`**
  - Parse A1-style reference (e.g., 'Sheet1!$A$1')
  - Returns: `{ col, colIndex, row, sheet?, absolute: { col, row } }`

- **`parseR1C1(ref: string): R1C1Reference`**
  - Parse R1C1-style reference (e.g., 'R1C1', 'R[1]C[-1]')

- **`parseRange(ref: string): RangeReference`**
  - Parse range reference (e.g., 'A1:C10', 'Sheet1!A:A')
  - Returns: `{ original, start, end, sheet? }`

- **`parseReference(ref: string): ParsedReference`**
  - Auto-detect and parse any reference type
  - Returns: `{ type: 'cell' | 'range' | 'r1c1' | 'error', ref? }`

- **`isValidA1(ref: string): boolean`**
  - Check if string is a valid A1 reference

- **`isValidR1C1(ref: string): boolean`**
  - Check if string is a valid R1C1 reference

- **`toA1(pos: { row, col }, options?): string`**
  - Convert row/col to A1 notation
  - Options: `absolute: { row?, col? }`, `sheet?`

- **`toR1C1(a1: string, options?): string`**
  - Convert A1 notation to R1C1

- **`offsetReference(ref: CellReference, offset: { rowOffset?, colOffset? }): CellReference`**
  - Offset a cell reference by row/column deltas

- **`expandRange(rangeStr: string): string[]`**
  - Expand range to array of individual cell references
  - Example: 'A1:B2' → ['A1', 'A2', 'B1', 'B2']

- **`rangeContains(rangeStr: string, cellStr: string): boolean`**
  - Check if a cell is within a range

- **`rangeIntersects(range1Str: string, range2Str: string): boolean`**
  - Check if two ranges overlap

- **`mergeRanges(ranges: string[]): string[]`**
  - Merge overlapping or adjacent ranges

### Types

#### Cell Value Types

```typescript
type CellValueType = 'boolean' | 'number' | 'string' | 'date' | 'error' | 'empty'

interface CellValue {
  v: CellPrimitive           // Raw value
  t: CellValueType          // Type
  w?: string                // Formatted display text
  f?: string                // Formula
  e?: CellErrorValue        // Error value
}
```

#### Cell Object

```typescript
interface Cell {
  _id: string                    // Unique ID (e.g., "Sheet1!A1")
  sheet: string                  // Sheet name
  row: number                    // Row number (1-indexed)
  col: string                    // Column letters
  colIndex: number               // Column index (0-indexed)
  value: CellValue              // Cell value
  format?: CellFormat           // Formatting
  metadata?: CellMetadata       // Comments, validation, hyperlinks
  dependencies?: string[]       // Cells this depends on
  dependents?: string[]         // Cells that depend on this
  updatedAt: Date               // Last modified
  createdAt: Date               // Created timestamp
}
```

#### Cell Format

```typescript
interface CellFormat {
  font?: {
    name?: string
    size?: number
    bold?: boolean
    italic?: boolean
    underline?: boolean | 'single' | 'double'
    strikethrough?: boolean
    color?: Color
  }
  fill?: {
    type: 'solid' | 'pattern' | 'gradient'
    color?: Color
    patternType?: string
    fgColor?: Color
    bgColor?: Color
  }
  border?: {
    top?: Border
    right?: Border
    bottom?: Border
    left?: Border
  }
  alignment?: {
    horizontal?: 'left' | 'center' | 'right'
    vertical?: 'top' | 'center' | 'bottom'
    wrapText?: boolean
    shrinkToFit?: boolean
    textRotation?: number
    indent?: number
  }
  numberFormat?: string
  protection?: {
    locked?: boolean
    hidden?: boolean
  }
}
```

#### Address Types

```typescript
interface CellReference {
  original: string
  col: string
  colIndex: number
  row: number
  sheet?: string
  absolute: {
    col: boolean      // $ before column
    row: boolean      // $ before row
  }
}

interface RangeReference {
  original: string
  start: CellReference
  end: CellReference
  sheet?: string
}

interface R1C1Reference {
  original: string
  sheet?: string
  row: { value: number; relative: boolean }
  col: { value: number; relative: boolean }
}
```

### Import Module (`excel.do/import`)

#### Functions

- **`parseWorkbook(data: Buffer | Uint8Array | ArrayBuffer, options?: ImportOptions): ParsedWorkbook`**
  - Parse an XLSX buffer into workbook structure
  - Returns: `{ name, sheets: ParsedSheet[], metadata? }`

- **`parseSheet(worksheet: XLSX.WorkSheet, name: string, options?: ImportOptions): ParsedSheet`**
  - Parse a SheetJS worksheet into sheet structure
  - Returns: `{ name, cells: ImportedCell[], dimensions?, merges?, columnWidths?, rowHeights? }`

- **`rowsFromSheet(worksheet: XLSX.WorkSheet, options?: ImportOptions): RowDocument[]`**
  - Extract rows as documents from a worksheet
  - Supports header detection, custom headers, and data filtering
  - Returns: Array of row objects with column values

- **`cellFromSheetJS(sjsCell: XLSX.CellObject, sheet: string, col: string, row: number, options?: ImportOptions): ImportedCell`**
  - Convert a SheetJS cell to excel.do Cell type
  - Handles type inference, formatting, and metadata extraction

### Export Module (`excel.do/export`)

#### Functions

- **`toWorkbook(workbookData: WorkbookData, options?: ExportOptions): XLSX.WorkBook`**
  - Convert workbook data to a SheetJS WorkBook
  - Returns: SheetJS WorkBook with sheets and metadata

- **`toSheet(sheetData: SheetData, options?: ExportOptions): XLSX.WorkSheet`**
  - Convert sheet data to a SheetJS WorkSheet
  - Includes cell data, column widths, row heights, and merges

- **`toXLSX(workbookData: WorkbookData, options?: ExportOptions): Promise<Buffer>`**
  - Generate an XLSX buffer from workbook data
  - Async function that returns Buffer ready for file writing

- **`toCSV(sheetData: SheetData, options?: ExportOptions): string`**
  - Generate a CSV string from sheet data
  - Supports custom delimiters, headers, and formatting

- **`cellToSheetJS(cell: Cell, options?: ExportOptions): XLSX.CellObject`**
  - Convert an excel.do Cell to a SheetJS CellObject
  - Handles formatting, formulas, and metadata preservation

### Formula Module (`excel.do/formula`)

#### Functions

- **`parseFormula(formula: string, options?: ParserOptions): ParseResult`**
  - Parse an Excel formula into an AST
  - Returns: `{ ast, errors: ParseError[], formula }`

- **`tokenize(formula: string): Token[]`**
  - Tokenize a formula string into individual tokens

#### Types

```typescript
type TokenType =
  | 'NUMBER' | 'STRING' | 'BOOLEAN' | 'ERROR'
  | 'CELL_REF' | 'RANGE' | 'NAMED_RANGE'
  | 'OPERATOR_ADD' | 'OPERATOR_SUB' | 'OPERATOR_MUL' | 'OPERATOR_DIV'
  | 'OPERATOR_POW' | 'OPERATOR_CONCAT' | 'OPERATOR_EQ' | 'OPERATOR_NE'
  | 'OPERATOR_LT' | 'OPERATOR_GT' | 'OPERATOR_LTE' | 'OPERATOR_GTE'
  | 'LPAREN' | 'RPAREN' | 'COMMA' | 'COLON' | 'SEMICOLON'
  | 'FUNCTION' | 'WHITESPACE' | 'EOF'

interface Token {
  type: TokenType
  value: string
  start: number
  end: number
}

type ASTNodeType =
  | 'Program' | 'BinaryExpression' | 'UnaryExpression'
  | 'FunctionCall' | 'CellReference' | 'RangeReference'
  | 'Literal' | 'ArrayLiteral' | 'ErrorValue'

interface ParseResult {
  ast: ProgramNode
  errors: ParseError[]
  formula: string
}

interface ParseError {
  message: string
  position: number
  length: number
  code: ParseErrorCode
}
```

## Module Exports

excel.do provides multiple entry points for different use cases:

```typescript
// Main module - all exports
import * as excel from 'excel.do'

// Cell module
import { createCell, parseA1, expandRange } from 'excel.do/cell'

// Import module
import { parseWorkbook, parseSheet, rowsFromSheet, cellFromSheetJS } from 'excel.do/import'

// Export module
import { toXLSX, toCSV, toWorkbook, toSheet, cellToSheetJS } from 'excel.do/export'

// Formula module
import { parseFormula, tokenize } from 'excel.do/formula'
```

## Storage and Persistence

excel.do supports multiple storage backends:

- **mongo.do**: Full MongoDB integration for persistent cell storage
- **Custom Database**: Implement the `StorageBackend` interface
- **WebSocket**: Real-time collaboration support via gitx.do

## Testing

Run the test suite:

```bash
npm test              # Watch mode
npm run test:run      # Single run
npm run test:coverage # With coverage report
```

## Development

```bash
npm run dev       # Start Wrangler dev server
npm run build     # Build with tsup
npm run typecheck # Type check with tsc
npm run lint      # Lint with ESLint
npm run deploy    # Deploy to Cloudflare Workers
```

## TypeScript Support

excel.do is written in TypeScript with full type definitions. All types are exported from the main module:

```typescript
import type {
  Cell,
  CellValue,
  CellFormat,
  CellReference,
  RangeReference,
} from 'excel.do'
```

## Error Handling

Cell validation and formula parsing provide detailed error information:

```typescript
import { validateCell } from 'excel.do/cell'
import { parseFormula } from 'excel.do/formula'

// Cell validation
const validation = validateCell(cell)
if (!validation.valid) {
  console.error('Validation errors:', validation.errors)
}

// Formula parsing
const result = parseFormula('=INVALID(')
if (result.errors.length > 0) {
  result.errors.forEach(err => {
    console.error(`Error at position ${err.position}: ${err.message}`)
  })
}
```

## Browser Support

excel.do runs on:

- Cloudflare Workers (ES2022+)
- Node.js 18+
- Modern browsers with ES2022 support

## License

MIT

## Author

Nathan Clevenger

## Repository

https://github.com/nathanclevenger/excel.do

## Keywords

spreadsheet, excel, xlsx, cloudflare, workers, edge, mongodb, git, version-control, collaboration
