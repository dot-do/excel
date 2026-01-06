/**
 * XLSX Export Module for excel.do
 *
 * Provides functionality to export internal workbook/sheet data to SheetJS format
 * and generate XLSX/CSV output.
 */

import * as XLSX from 'xlsx'
import type { Cell, CellFormat, Color, RGBColor } from '../types'
import type { WorkBook, WorkSheet, CellObject } from 'xlsx'

// ============================================================================
// Export Types
// ============================================================================

/** Options for exporting data */
export interface ExportOptions {
  /** Include column headers in output */
  includeHeaders?: boolean
  /** Custom header names for columns */
  headerNames?: Record<string, string>
  /** Whether to include formatting/styles */
  includeFormatting?: boolean
  /** Column widths in characters */
  columnWidths?: Record<string, number>
  /** Default column width */
  defaultColumnWidth?: number
  /** Row heights */
  rowHeights?: Record<number, number>
  /** Sheet name for single-sheet export */
  sheetName?: string
  /** Date format for date cells */
  dateFormat?: string
  /** Number format for number cells */
  numberFormat?: string
  /** Whether to preserve formulas */
  preserveFormulas?: boolean
  /** CSV delimiter */
  csvDelimiter?: string
  /** CSV row separator */
  csvRowSeparator?: string
  /** Whether to quote all CSV fields */
  csvQuoteAll?: boolean
  /** Merge ranges (internal use) */
  merges?: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>
}

/** Data for a single sheet */
export interface SheetData {
  name: string
  cells: Cell[]
  columnWidths?: Record<string, number>
  rowHeights?: Record<number, number>
}

/** Data for a complete workbook */
export interface WorkbookData {
  sheets: SheetData[]
  activeSheet?: string
  metadata?: {
    title?: string
    author?: string
    created?: Date
    modified?: Date
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Convert hex color to RGB string for SheetJS
 */
function colorToRGB(color: Color | undefined): string | undefined {
  if (!color) return undefined

  if (typeof color === 'string') {
    // Remove # prefix if present
    const hex = color.replace(/^#/, '')
    return hex.toUpperCase()
  }

  // RGB object
  const rgb = color as RGBColor
  const r = Math.round(rgb.r).toString(16).padStart(2, '0')
  const g = Math.round(rgb.g).toString(16).padStart(2, '0')
  const b = Math.round(rgb.b).toString(16).padStart(2, '0')

  if (rgb.a !== undefined && rgb.a < 1) {
    const a = Math.round(rgb.a * 255).toString(16).padStart(2, '0')
    return `${a}${r}${g}${b}`.toUpperCase()
  }

  return `${r}${g}${b}`.toUpperCase()
}

/**
 * Convert column letter(s) to 0-based index
 */
function colToIndex(col: string): number {
  let index = 0
  const upper = col.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64)
  }
  return index - 1
}

/**
 * Convert 0-based index to column letter(s)
 */
function indexToCol(index: number): string {
  let col = ''
  let n = index + 1
  while (n > 0) {
    const remainder = (n - 1) % 26
    col = String.fromCharCode(65 + remainder) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}

// ============================================================================
// Cell Conversion
// ============================================================================

/**
 * Convert an excel.do Cell to a SheetJS CellObject.
 */
export function cellToSheetJS(cell: Cell, options?: ExportOptions): CellObject {
  const result: CellObject = {} as CellObject
  const value = cell.value

  // Determine cell type and value
  switch (value.t) {
    case 'string':
      result.t = 's'
      result.v = value.v as string
      break
    case 'number':
      result.t = 'n'
      result.v = value.v as number
      break
    case 'boolean':
      result.t = 'b'
      result.v = value.v as boolean
      break
    case 'date':
      result.t = 'd'
      result.v = value.v as Date
      break
    case 'error':
      result.t = 'e'
      result.v = value.v as string
      break
    case 'empty':
    default:
      result.t = 'z'
      // Empty cells have no value
      break
  }

  // Preserve formula (strip leading =)
  if (value.f) {
    result.f = value.f.startsWith('=') ? value.f.slice(1) : value.f
  }

  // Preserve formatted text
  if (value.w) {
    result.w = value.w
  }

  // Convert formatting if enabled
  if (options?.includeFormatting && cell.format) {
    result.s = convertFormat(cell.format)
  }

  // Apply number format from cell format
  if (options?.includeFormatting && cell.format?.numberFormat) {
    result.z = cell.format.numberFormat
  }

  // Apply date format from options
  if (options?.dateFormat && value.t === 'date') {
    result.z = options.dateFormat
  }

  // Apply number format from options
  if (options?.numberFormat && value.t === 'number' && !cell.format?.numberFormat) {
    result.z = options.numberFormat
  }

  // Convert hyperlink
  if (cell.metadata?.hyperlink) {
    result.l = {
      Target: cell.metadata.hyperlink.target,
    }
    if (cell.metadata.hyperlink.tooltip) {
      result.l.Tooltip = cell.metadata.hyperlink.tooltip
    }
  }

  // Convert comment
  if (cell.metadata?.comment) {
    result.c = [{
      t: cell.metadata.comment.text,
      a: cell.metadata.comment.author,
    }]
  }

  return result
}

/**
 * Convert CellFormat to SheetJS style object
 */
function convertFormat(format: CellFormat): any {
  const style: any = {}

  // Font conversion
  if (format.font) {
    style.font = {}
    if (format.font.bold) style.font.bold = true
    if (format.font.italic) style.font.italic = true
    if (format.font.strikethrough) style.font.strike = true
    if (format.font.size) style.font.sz = format.font.size
    if (format.font.name) style.font.name = format.font.name

    // Underline
    if (format.font.underline) {
      if (format.font.underline === 'double') {
        style.font.underline = 2
      } else {
        style.font.underline = true
      }
    }

    // Font color
    if (format.font.color) {
      style.font.color = { rgb: colorToRGB(format.font.color) }
    }
  }

  // Fill conversion
  if (format.fill) {
    style.fill = {}
    if (format.fill.type === 'solid' && format.fill.color) {
      style.fill.fgColor = { rgb: colorToRGB(format.fill.color) }
      style.fill.patternType = 'solid'
    } else if (format.fill.type === 'pattern') {
      if (format.fill.patternType) style.fill.patternType = format.fill.patternType
      if (format.fill.fgColor) style.fill.fgColor = { rgb: colorToRGB(format.fill.fgColor) }
      if (format.fill.bgColor) style.fill.bgColor = { rgb: colorToRGB(format.fill.bgColor) }
    }
  }

  // Border conversion
  if (format.border) {
    style.border = {}
    if (format.border.top) {
      style.border.top = {
        style: format.border.top.style,
        color: format.border.top.color ? { rgb: colorToRGB(format.border.top.color) } : undefined,
      }
    }
    if (format.border.right) {
      style.border.right = {
        style: format.border.right.style,
        color: format.border.right.color ? { rgb: colorToRGB(format.border.right.color) } : undefined,
      }
    }
    if (format.border.bottom) {
      style.border.bottom = {
        style: format.border.bottom.style,
        color: format.border.bottom.color ? { rgb: colorToRGB(format.border.bottom.color) } : undefined,
      }
    }
    if (format.border.left) {
      style.border.left = {
        style: format.border.left.style,
        color: format.border.left.color ? { rgb: colorToRGB(format.border.left.color) } : undefined,
      }
    }
  }

  // Alignment conversion
  if (format.alignment) {
    style.alignment = {}
    if (format.alignment.horizontal) style.alignment.horizontal = format.alignment.horizontal
    if (format.alignment.vertical) style.alignment.vertical = format.alignment.vertical
    if (format.alignment.wrapText) style.alignment.wrapText = true
    if (format.alignment.shrinkToFit) style.alignment.shrinkToFit = true
    if (format.alignment.textRotation !== undefined) style.alignment.textRotation = format.alignment.textRotation
    if (format.alignment.indent !== undefined) style.alignment.indent = format.alignment.indent
  }

  return style
}

// ============================================================================
// Sheet Conversion
// ============================================================================

/**
 * Convert sheet data to a SheetJS WorkSheet.
 */
export function toSheet(sheetData: SheetData, options?: ExportOptions): WorkSheet {
  const ws: WorkSheet = {}

  if (sheetData.cells.length === 0) {
    return ws
  }

  // Track range
  let minRow = Infinity, maxRow = -Infinity
  let minCol = Infinity, maxCol = -Infinity

  // Add cells to worksheet
  for (const cell of sheetData.cells) {
    const address = `${cell.col}${cell.row}`
    ws[address] = cellToSheetJS(cell, options)

    // Update range tracking
    minRow = Math.min(minRow, cell.row)
    maxRow = Math.max(maxRow, cell.row)
    minCol = Math.min(minCol, cell.colIndex)
    maxCol = Math.max(maxCol, cell.colIndex)
  }

  // Set range reference
  if (sheetData.cells.length > 0) {
    const startCol = indexToCol(minCol)
    const endCol = indexToCol(maxCol)
    ws['!ref'] = `${startCol}${minRow}:${endCol}${maxRow}`
  }

  // Set column widths
  const columnWidths = { ...options?.columnWidths, ...sheetData.columnWidths }
  const defaultWidth = options?.defaultColumnWidth

  if (Object.keys(columnWidths).length > 0 || defaultWidth) {
    const cols: XLSX.ColInfo[] = []

    // Find max column index from widths or cells
    let maxColIndex = maxCol
    for (const col of Object.keys(columnWidths)) {
      maxColIndex = Math.max(maxColIndex, colToIndex(col))
    }

    for (let i = 0; i <= maxColIndex; i++) {
      const colLetter = indexToCol(i)
      const width = columnWidths[colLetter] ?? defaultWidth
      if (width !== undefined) {
        cols[i] = { wch: width }
      }
    }

    ws['!cols'] = cols
  }

  // Set row heights
  const rowHeights = { ...options?.rowHeights, ...sheetData.rowHeights }
  if (Object.keys(rowHeights).length > 0) {
    const rows: XLSX.RowInfo[] = []

    for (const [rowStr, height] of Object.entries(rowHeights)) {
      const rowNum = parseInt(rowStr, 10)
      rows[rowNum - 1] = { hpt: height } // Row index is 0-based
    }

    ws['!rows'] = rows
  }

  // Set merged cells
  if (options?.merges && options.merges.length > 0) {
    ws['!merges'] = options.merges
  }

  return ws
}

// ============================================================================
// Workbook Conversion
// ============================================================================

/**
 * Convert workbook data to a SheetJS WorkBook.
 */
export function toWorkbook(workbookData: WorkbookData, options?: ExportOptions): WorkBook {
  const wb: WorkBook = {
    SheetNames: [],
    Sheets: {},
  }

  // Add each sheet
  for (const sheetData of workbookData.sheets) {
    wb.SheetNames.push(sheetData.name)
    wb.Sheets[sheetData.name] = toSheet(sheetData, options)
  }

  // Set workbook metadata
  if (workbookData.metadata) {
    wb.Props = {}
    if (workbookData.metadata.title) wb.Props.Title = workbookData.metadata.title
    if (workbookData.metadata.author) wb.Props.Author = workbookData.metadata.author
    if (workbookData.metadata.created) wb.Props.CreatedDate = workbookData.metadata.created
    if (workbookData.metadata.modified) wb.Props.ModifiedDate = workbookData.metadata.modified
  }

  // Set active sheet
  if (workbookData.activeSheet) {
    const activeIndex = wb.SheetNames.indexOf(workbookData.activeSheet)
    if (activeIndex >= 0) {
      wb.Workbook = {
        Sheets: wb.SheetNames.map(() => ({
          Hidden: 0,
        })),
      }
    }
  }

  return wb
}

// ============================================================================
// XLSX Export
// ============================================================================

/**
 * Generate an XLSX buffer from workbook data.
 */
export async function toXLSX(workbookData: WorkbookData, options?: ExportOptions): Promise<Buffer> {
  const wb = toWorkbook(workbookData, options)

  // Configure write options
  const writeOpts: XLSX.WritingOptions = {
    type: 'buffer',
    bookType: 'xlsx',
    cellStyles: options?.includeFormatting ?? false,
  }

  // Generate buffer
  const buffer = XLSX.write(wb, writeOpts)

  // Ensure we return a Buffer
  if (Buffer.isBuffer(buffer)) {
    return buffer
  }

  return Buffer.from(buffer)
}

// ============================================================================
// CSV Export
// ============================================================================

/**
 * Generate a CSV string from sheet data.
 */
export function toCSV(sheetData: SheetData, options?: ExportOptions): string {
  if (sheetData.cells.length === 0) {
    return ''
  }

  const delimiter = options?.csvDelimiter ?? ','
  const rowSeparator = options?.csvRowSeparator ?? '\n'
  const quoteAll = options?.csvQuoteAll ?? false

  // Group cells by row and find range
  const cellMap = new Map<string, Cell>()
  let minRow = Infinity, maxRow = -Infinity
  let minCol = Infinity, maxCol = -Infinity

  for (const cell of sheetData.cells) {
    const key = `${cell.row}:${cell.colIndex}`
    cellMap.set(key, cell)
    minRow = Math.min(minRow, cell.row)
    maxRow = Math.max(maxRow, cell.row)
    minCol = Math.min(minCol, cell.colIndex)
    maxCol = Math.max(maxCol, cell.colIndex)
  }

  const rows: string[] = []

  // Add header row if requested
  if (options?.includeHeaders && options.headerNames) {
    const headerCells: string[] = []
    for (let col = minCol; col <= maxCol; col++) {
      const colLetter = indexToCol(col)
      const header = options.headerNames[colLetter] ?? colLetter
      headerCells.push(escapeCSV(header, delimiter, quoteAll))
    }
    rows.push(headerCells.join(delimiter))
  }

  // Add data rows
  for (let row = minRow; row <= maxRow; row++) {
    const rowCells: string[] = []

    for (let col = minCol; col <= maxCol; col++) {
      const key = `${row}:${col}`
      const cell = cellMap.get(key)

      if (!cell) {
        rowCells.push('')
        continue
      }

      // Get formatted value
      const value = formatCSVValue(cell, options)
      rowCells.push(escapeCSV(value, delimiter, quoteAll))
    }

    rows.push(rowCells.join(delimiter))
  }

  return rows.join(rowSeparator)
}

/**
 * Format cell value for CSV output
 */
function formatCSVValue(cell: Cell, _options?: ExportOptions): string {
  const value = cell.value

  // Use formatted text if available
  if (value.w) {
    return value.w
  }

  // Format based on type
  switch (value.t) {
    case 'string':
      return value.v as string
    case 'number':
      return String(value.v)
    case 'boolean':
      return value.v ? 'TRUE' : 'FALSE'
    case 'date':
      const date = value.v as Date
      return date.toISOString().split('T')[0] // YYYY-MM-DD
    case 'error':
      return value.v as string
    case 'empty':
    default:
      return ''
  }
}

/**
 * Escape a value for CSV format
 */
function escapeCSV(value: string, delimiter: string, quoteAll: boolean): string {
  if (value === '') {
    return quoteAll ? '""' : ''
  }

  // Check if quoting is needed
  const needsQuoting = quoteAll ||
    value.includes(delimiter) ||
    value.includes('"') ||
    value.includes('\n') ||
    value.includes('\r')

  if (!needsQuoting) {
    return value
  }

  // Escape double quotes by doubling them
  const escaped = value.replace(/"/g, '""')
  return `"${escaped}"`
}
