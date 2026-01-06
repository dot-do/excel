/**
 * XLSX Import Implementation
 *
 * TDD GREEN: Full implementation to make tests pass.
 */

import * as XLSX from 'xlsx'
import type { Cell, CellFormat, CellMetadata, CellValue, CellValueType, CellErrorValue, CellPrimitive } from '../types'
import { colToIndex, indexToCol } from '../cell/address'

/** Import options */
export interface ImportOptions {
  sheets?: string[]
  range?: string
  headers?: boolean | 'auto' | string[]
  inferDates?: boolean
  inferNumbers?: boolean
  inferBooleans?: boolean
  preserveStyles?: boolean
  extractDependencies?: boolean
  skipEmptyRows?: boolean
  includeRowMeta?: boolean
}

/** Parsed workbook metadata */
export interface WorkbookMetadata {
  title?: string
  author?: string
  createdDate?: Date
  modifiedDate?: Date
}

/** Sheet dimensions */
export interface SheetDimensions {
  startRow: number
  endRow: number
  startCol: string
  endCol: string
}

/** Merge cell info */
export interface MergeInfo {
  start: { row: number; col: string }
  end: { row: number; col: string }
}

/** Parsed cell for import */
export interface ImportedCell extends Omit<Cell, 'createdAt' | 'updatedAt'> {
  dependencies?: string[]
}

/** Parsed sheet result */
export interface ParsedSheet {
  name: string
  cells: ImportedCell[]
  dimensions?: SheetDimensions
  merges?: MergeInfo[]
  columnWidths?: Record<string, number>
  rowHeights?: Record<number, number>
}

/** Parsed workbook result */
export interface ParsedWorkbook {
  name: string
  sheets: ParsedSheet[]
  metadata?: WorkbookMetadata
}

/** Row document from sheet */
export interface RowDocument {
  [key: string]: unknown
  _rowNumber?: number
  _range?: string
}

// Excel error values
const ERROR_VALUES: CellErrorValue[] = [
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?',
  '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'
]

// Date pattern matching
const DATE_PATTERNS = [
  /^\d{4}-\d{2}-\d{2}$/,                    // 2024-01-15
  /^\d{2}\/\d{2}\/\d{4}$/,                  // 01/15/2024
  /^[A-Za-z]+ \d{1,2}, \d{4}$/,             // January 15, 2024
]

/**
 * Infer value type from a primitive
 */
function inferType(value: unknown): CellValueType {
  if (value === null || value === undefined) return 'empty'
  if (typeof value === 'boolean') return 'boolean'
  if (typeof value === 'number') return 'number'
  if (value instanceof Date) return 'date'
  if (typeof value === 'string') {
    if (ERROR_VALUES.includes(value as CellErrorValue)) return 'error'
    return 'string'
  }
  return 'string'
}

/**
 * Try to infer a date from a string
 */
function tryInferDate(value: string): Date | null {
  for (const pattern of DATE_PATTERNS) {
    if (pattern.test(value)) {
      const parsed = new Date(value)
      if (!isNaN(parsed.getTime())) {
        return parsed
      }
    }
  }
  return null
}

/**
 * Try to infer a number from a string (including currency)
 */
function tryInferNumber(value: string): number | null {
  // Remove currency symbols and thousands separators
  const cleaned = value.replace(/[$,€£¥]/g, '').trim()
  const num = parseFloat(cleaned)
  if (!isNaN(num) && isFinite(num)) {
    return num
  }
  return null
}

/**
 * Try to infer a boolean from a string
 */
function tryInferBoolean(value: string): boolean | null {
  const lower = value.toLowerCase()
  if (lower === 'true' || lower === 'yes') return true
  if (lower === 'false' || lower === 'no') return false
  return null
}

/**
 * Generate a cell ID from sheet, column, and row
 */
function getCellId(sheet: string, col: string, row: number): string {
  const needsQuotes = /[\s']/.test(sheet)
  if (needsQuotes) {
    const escapedSheet = sheet.replace(/'/g, "''")
    return `'${escapedSheet}'!${col}${row}`
  }
  return `${sheet}!${col}${row}`
}

/**
 * Convert Excel serial date to JavaScript Date
 */
function excelSerialToDate(serial: number): Date {
  // Excel dates start from 1900-01-01 (serial 1)
  // but there's a bug where 1900 is treated as a leap year
  const baseDate = new Date(Date.UTC(1899, 11, 30))
  return new Date(baseDate.getTime() + serial * 86400000)
}

/**
 * Extract dependencies from a formula string
 */
function extractDependencies(formula: string, currentSheet: string): string[] {
  const deps: string[] = []

  // Match cell references like A1, $A$1, Sheet1!A1, 'Sheet Name'!A1
  const cellRefRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]+)\$?(\d+)/gi
  let match: RegExpExecArray | null

  while ((match = cellRefRegex.exec(formula)) !== null) {
    const sheetName = match[1] || match[2] || currentSheet
    const col = match[3].toUpperCase()
    const row = parseInt(match[4], 10)
    deps.push(`${sheetName}!${col}${row}`)
  }

  // Also extract range references
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/gi
  while ((match = rangeRegex.exec(formula)) !== null) {
    const sheetName = match[1] || match[2] || currentSheet
    deps.push(`${sheetName}!${match[3]}${match[4]}:${match[5]}${match[6]}`)
  }

  return [...new Set(deps)]
}

/**
 * Parse a range string into start/end coordinates
 */
function parseRangeString(rangeStr: string): { startRow: number; endRow: number; startCol: number; endCol: number } | null {
  // Handle row-only ranges like "2:5"
  const rowOnlyMatch = rangeStr.match(/^(\d+):(\d+)$/)
  if (rowOnlyMatch) {
    return {
      startRow: parseInt(rowOnlyMatch[1], 10),
      endRow: parseInt(rowOnlyMatch[2], 10),
      startCol: 0,
      endCol: 16383 // XFD
    }
  }

  // Handle column-only ranges like "A:C"
  const colOnlyMatch = rangeStr.match(/^([A-Z]+):([A-Z]+)$/i)
  if (colOnlyMatch) {
    return {
      startRow: 1,
      endRow: 1048576,
      startCol: colToIndex(colOnlyMatch[1]),
      endCol: colToIndex(colOnlyMatch[2])
    }
  }

  // Handle standard ranges like "A1:B10"
  const standardMatch = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i)
  if (standardMatch) {
    return {
      startRow: parseInt(standardMatch[2], 10),
      endRow: parseInt(standardMatch[4], 10),
      startCol: colToIndex(standardMatch[1]),
      endCol: colToIndex(standardMatch[3])
    }
  }

  return null
}

/** SheetJS cell style type (internal representation) */
interface SJSStyle {
  font?: {
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strike?: boolean
    sz?: number
    name?: string
    color?: { rgb?: string }
  }
  fill?: {
    patternType?: string
    fgColor?: { rgb?: string }
  }
  border?: {
    top?: { style?: string; color?: { rgb?: string } }
    bottom?: { style?: string; color?: { rgb?: string } }
    left?: { style?: string; color?: { rgb?: string } }
    right?: { style?: string; color?: { rgb?: string } }
  }
  alignment?: {
    horizontal?: string
    vertical?: string
    wrapText?: boolean
    textRotation?: number
  }
  protection?: {
    locked?: boolean
    hidden?: boolean
  }
}

/**
 * Convert SheetJS style to our CellFormat
 */
function convertStyle(sjsStyle: SJSStyle | undefined): CellFormat | undefined {
  if (!sjsStyle) return undefined

  const format: CellFormat = {}

  // Font styling
  if (sjsStyle.font) {
    format.font = {}
    if (sjsStyle.font.bold) format.font.bold = true
    if (sjsStyle.font.italic) format.font.italic = true
    if (sjsStyle.font.underline) format.font.underline = true
    if (sjsStyle.font.strike) format.font.strikethrough = true
    if (sjsStyle.font.sz) format.font.size = sjsStyle.font.sz
    if (sjsStyle.font.name) format.font.name = sjsStyle.font.name
    if (sjsStyle.font.color?.rgb) {
      format.font.color = `#${sjsStyle.font.color.rgb}`
    }
  }

  // Fill styling
  if (sjsStyle.fill) {
    const fillType = sjsStyle.fill.patternType === 'solid' ? 'solid'
      : sjsStyle.fill.patternType === 'gradient' ? 'gradient'
      : 'pattern'
    format.fill = { type: fillType as 'solid' | 'pattern' | 'gradient' }
    if (sjsStyle.fill.fgColor?.rgb) {
      format.fill.color = `#${sjsStyle.fill.fgColor.rgb}`
    }
  }

  // Border styling
  if (sjsStyle.border) {
    format.border = {}
    const borderStyleMap: Record<string, 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'none'> = {
      'thin': 'thin',
      'medium': 'medium',
      'thick': 'thick',
      'dashed': 'dashed',
      'dotted': 'dotted',
      'double': 'double',
      'none': 'none',
      'hair': 'thin',
      'mediumDashed': 'dashed',
      'dashDot': 'dashed',
      'mediumDashDot': 'dashed',
      'dashDotDot': 'dashed',
      'mediumDashDotDot': 'dashed',
      'slantDashDot': 'dashed',
    }
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const borderSide = sjsStyle.border[side]
      if (borderSide?.style) {
        format.border[side] = {
          style: borderStyleMap[borderSide.style] || 'thin',
          color: borderSide.color?.rgb ? `#${borderSide.color.rgb}` : undefined
        }
      }
    }
  }

  // Alignment
  if (sjsStyle.alignment) {
    format.alignment = {}
    if (sjsStyle.alignment.horizontal) format.alignment.horizontal = sjsStyle.alignment.horizontal as NonNullable<CellFormat['alignment']>['horizontal']
    if (sjsStyle.alignment.vertical) format.alignment.vertical = sjsStyle.alignment.vertical as NonNullable<CellFormat['alignment']>['vertical']
    if (sjsStyle.alignment.wrapText) format.alignment.wrapText = true
    if (sjsStyle.alignment.textRotation) format.alignment.textRotation = sjsStyle.alignment.textRotation
  }

  // Protection
  if (sjsStyle.protection) {
    format.protection = {}
    if (sjsStyle.protection.locked !== undefined) format.protection.locked = sjsStyle.protection.locked
    if (sjsStyle.protection.hidden !== undefined) format.protection.hidden = sjsStyle.protection.hidden
  }

  return Object.keys(format).length > 0 ? format : undefined
}

// XLSX file magic bytes (ZIP PK signature)
const XLSX_MAGIC = [0x50, 0x4B, 0x03, 0x04]

/**
 * Parse an XLSX buffer to internal workbook structure
 */
export function parseWorkbook(
  data: Buffer | Uint8Array | ArrayBuffer,
  options?: ImportOptions
): ParsedWorkbook {
  // Validate input is a valid ZIP file (XLSX is ZIP-based)
  const bytes = data instanceof ArrayBuffer
    ? new Uint8Array(data)
    : data instanceof Uint8Array
      ? data
      : new Uint8Array(data)

  if (bytes.length < 4 || !XLSX_MAGIC.every((b, i) => bytes[i] === b)) {
    throw new Error('Invalid XLSX file: not a valid ZIP/XLSX format')
  }

  // Read the workbook using SheetJS
  let wb: XLSX.WorkBook
  try {
    wb = XLSX.read(data, { type: 'buffer', cellStyles: true, cellDates: true })
    // Validate that we got a valid workbook with sheets
    if (!wb.SheetNames || wb.SheetNames.length === 0) {
      throw new Error('Invalid XLSX file: no sheets found')
    }
  } catch (e) {
    throw new Error(`Failed to parse XLSX file: ${(e as Error).message}`)
  }

  // Filter sheets if specified
  let sheetNames = wb.SheetNames
  if (options?.sheets && options.sheets.length > 0) {
    sheetNames = sheetNames.filter(name => options.sheets!.includes(name))
  }

  // Parse each sheet
  const sheets: ParsedSheet[] = sheetNames.map(name => {
    const ws = wb.Sheets[name]
    return parseSheet(ws, name, options)
  })

  // Extract metadata
  const metadata: WorkbookMetadata | undefined = wb.Props ? {
    title: wb.Props.Title,
    author: wb.Props.Author,
    createdDate: wb.Props.CreatedDate ? new Date(wb.Props.CreatedDate) : undefined,
    modifiedDate: wb.Props.ModifiedDate ? new Date(wb.Props.ModifiedDate) : undefined,
  } : undefined

  return {
    name: wb.Props?.Title || 'Workbook',
    sheets,
    metadata,
  }
}

/**
 * Parse a SheetJS worksheet to internal sheet structure
 */
export function parseSheet(
  worksheet: XLSX.WorkSheet,
  name: string,
  options?: ImportOptions
): ParsedSheet {
  const cells: ImportedCell[] = []
  let dimensions: SheetDimensions | undefined

  // Get sheet range
  const ref = worksheet['!ref']
  if (!ref) {
    return { name, cells }
  }

  // Parse the range
  const [startRef, endRef] = ref.split(':')
  const startMatch = startRef.match(/([A-Z]+)(\d+)/i)
  const endMatch = endRef?.match(/([A-Z]+)(\d+)/i)

  if (startMatch) {
    dimensions = {
      startRow: parseInt(startMatch[2], 10),
      endRow: endMatch ? parseInt(endMatch[2], 10) : parseInt(startMatch[2], 10),
      startCol: startMatch[1].toUpperCase(),
      endCol: endMatch ? endMatch[1].toUpperCase() : startMatch[1].toUpperCase(),
    }
  }

  // Parse range limits from options
  let rangeLimit: ReturnType<typeof parseRangeString> = null
  if (options?.range) {
    rangeLimit = parseRangeString(options.range)
  }

  // Iterate over cells
  for (const cellAddress of Object.keys(worksheet)) {
    if (cellAddress.startsWith('!')) continue // Skip metadata keys

    const sjsCell = worksheet[cellAddress] as XLSX.CellObject
    if (!sjsCell) continue

    // Parse address
    const addrMatch = cellAddress.match(/([A-Z]+)(\d+)/i)
    if (!addrMatch) continue

    const col = addrMatch[1].toUpperCase()
    const row = parseInt(addrMatch[2], 10)
    const colIdx = colToIndex(col)

    // Apply range filter
    if (rangeLimit) {
      if (row < rangeLimit.startRow || row > rangeLimit.endRow) continue
      if (colIdx < rangeLimit.startCol || colIdx > rangeLimit.endCol) continue
    }

    const cell = cellFromSheetJS(sjsCell, name, col, row, options)
    cells.push(cell)
  }

  // Update dimensions based on range filter
  if (rangeLimit && dimensions) {
    dimensions.startRow = Math.max(dimensions.startRow, rangeLimit.startRow)
    dimensions.endRow = Math.min(dimensions.endRow, rangeLimit.endRow)
    dimensions.startCol = indexToCol(Math.max(colToIndex(dimensions.startCol), rangeLimit.startCol))
    dimensions.endCol = indexToCol(Math.min(colToIndex(dimensions.endCol), rangeLimit.endCol))
  }

  // Parse merges
  const merges: MergeInfo[] | undefined = worksheet['!merges']?.map((merge: XLSX.Range) => ({
    start: { row: merge.s.r + 1, col: indexToCol(merge.s.c) },
    end: { row: merge.e.r + 1, col: indexToCol(merge.e.c) },
  }))

  // Parse column widths
  const columnWidths: Record<string, number> | undefined = worksheet['!cols']?.reduce((acc, col, idx) => {
    if (col?.wch) {
      acc[indexToCol(idx)] = col.wch
    }
    return acc
  }, {} as Record<string, number>)

  // Parse row heights
  const rowHeights: Record<number, number> | undefined = worksheet['!rows']?.reduce((acc, row, idx) => {
    if (row?.hpt) {
      acc[idx + 1] = row.hpt
    }
    return acc
  }, {} as Record<number, number>)

  return {
    name,
    cells,
    dimensions,
    merges: merges?.length ? merges : undefined,
    columnWidths: columnWidths && Object.keys(columnWidths).length ? columnWidths : undefined,
    rowHeights: rowHeights && Object.keys(rowHeights).length ? rowHeights : undefined,
  }
}

/**
 * Convert a SheetJS cell to excel.do Cell type
 */
export function cellFromSheetJS(
  sjsCell: XLSX.CellObject,
  sheet: string,
  col: string,
  row: number,
  options?: ImportOptions
): ImportedCell {
  const colIndex = colToIndex(col)
  const _id = getCellId(sheet, col, row)

  // Determine value and type
  let value: CellPrimitive = sjsCell.v as CellPrimitive
  let valueType: CellValueType

  // Handle different SheetJS cell types
  switch (sjsCell.t) {
    case 's': // String
      valueType = 'string'
      // Type inference
      if (options?.inferDates && typeof value === 'string') {
        const date = tryInferDate(value)
        if (date) {
          value = date
          valueType = 'date'
        }
      } else if (options?.inferNumbers && typeof value === 'string') {
        const num = tryInferNumber(value)
        if (num !== null) {
          value = num
          valueType = 'number'
        }
      } else if (options?.inferBooleans && typeof value === 'string') {
        const bool = tryInferBoolean(value)
        if (bool !== null) {
          value = bool
          valueType = 'boolean'
        }
      }
      break
    case 'n': // Number
      valueType = 'number'
      // Check if it's a date stored as serial number
      if (options?.inferDates && sjsCell.w && /\d{1,2}\/\d{1,2}\/\d{4}/.test(sjsCell.w)) {
        value = excelSerialToDate(value as number)
        valueType = 'date'
      }
      break
    case 'b': // Boolean
      valueType = 'boolean'
      break
    case 'd': // Date
      valueType = 'date'
      break
    case 'e': // Error
      valueType = 'error'
      break
    default:
      valueType = value === null || value === undefined ? 'empty' : inferType(value)
  }

  // Build cell value
  const cellValue: CellValue = {
    v: value ?? null,
    t: valueType,
  }

  // Add formatted text
  if (sjsCell.w) {
    cellValue.w = sjsCell.w
  }

  // Add formula with = prefix normalization
  if (sjsCell.f) {
    cellValue.f = sjsCell.f.startsWith('=') ? sjsCell.f : `=${sjsCell.f}`
  }

  // Add error value
  if (valueType === 'error' && typeof value === 'string') {
    cellValue.e = value as CellErrorValue
  }

  // Build cell metadata
  const metadata: CellMetadata = {}

  // Handle array formula
  if (sjsCell.F) {
    metadata.arrayFormulaRange = sjsCell.F
  }

  // Handle hyperlink
  if (sjsCell.l) {
    metadata.hyperlink = {
      target: sjsCell.l.Target,
      tooltip: sjsCell.l.Tooltip,
    }
  }

  // Handle comment
  if (sjsCell.c && sjsCell.c.length > 0) {
    metadata.comment = {
      text: sjsCell.c[0].t,
      author: sjsCell.c[0].a,
    }
  }

  // Build result
  const cell: ImportedCell = {
    _id,
    sheet,
    row,
    col,
    colIndex,
    value: cellValue,
  }

  // Add format if preserving styles
  if (options?.preserveStyles && sjsCell.s) {
    const format = convertStyle(sjsCell.s)
    if (format) cell.format = format
  }

  // Add number format
  if (options?.preserveStyles && sjsCell.z) {
    cell.format = cell.format || {}
    cell.format.numberFormat = String(sjsCell.z)
  }

  // Add metadata if populated
  if (Object.keys(metadata).length > 0) {
    cell.metadata = metadata
  }

  // Extract dependencies if requested
  if (options?.extractDependencies && cellValue.f) {
    cell.dependencies = extractDependencies(cellValue.f, sheet)
  }

  return cell
}

/**
 * Extract rows as documents from a worksheet
 */
export function rowsFromSheet(
  worksheet: XLSX.WorkSheet,
  options?: ImportOptions & { headers?: boolean | 'auto' | string[] }
): RowDocument[] {
  const ref = worksheet['!ref']
  if (!ref) return []

  // Parse range limits
  let rangeLimit: ReturnType<typeof parseRangeString> = null
  if (options?.range) {
    rangeLimit = parseRangeString(options.range)
  }

  // Get dimensions from ref
  const [startRef, endRef] = ref.split(':')
  const startMatch = startRef.match(/([A-Z]+)(\d+)/i)
  const endMatch = endRef?.match(/([A-Z]+)(\d+)/i)

  if (!startMatch) return []

  let startRow = rangeLimit?.startRow ?? parseInt(startMatch[2], 10)
  let endRow = rangeLimit?.endRow ?? (endMatch ? parseInt(endMatch[2], 10) : startRow)
  let startColIdx = rangeLimit?.startCol ?? colToIndex(startMatch[1])
  let endColIdx = rangeLimit?.endCol ?? (endMatch ? colToIndex(endMatch[1]) : startColIdx)

  // Determine headers
  let headers: string[]
  let dataStartRow: number

  const useHeaders = options?.headers

  if (Array.isArray(useHeaders)) {
    // Use provided headers
    headers = useHeaders
    dataStartRow = startRow
  } else if (useHeaders === true || (useHeaders === 'auto' && shouldUseFirstRowAsHeaders(worksheet, startRow, startColIdx, endColIdx))) {
    // Use first row as headers
    headers = []
    const headerCounts: Record<string, number> = {}

    for (let colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
      const col = indexToCol(colIdx)
      const cellAddr = `${col}${startRow}`
      const cell = worksheet[cellAddr] as XLSX.CellObject | undefined
      let headerName = cell?.v?.toString() || col

      // Handle blank headers
      if (!headerName || headerName.trim() === '') {
        headerName = col
      }

      // Handle duplicate headers
      if (headerCounts[headerName] !== undefined) {
        headerCounts[headerName]++
        headerName = `${headerName}_${headerCounts[headerName]}`
      } else {
        headerCounts[headerName] = 0
      }

      headers.push(headerName)
    }
    dataStartRow = startRow + 1
  } else {
    // Use column letters as headers
    headers = []
    for (let colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
      headers.push(indexToCol(colIdx))
    }
    dataStartRow = startRow
  }

  // Extract rows
  const rows: RowDocument[] = []

  for (let rowNum = dataStartRow; rowNum <= endRow; rowNum++) {
    const doc: RowDocument = {}
    let isEmpty = true

    for (let i = 0; i < headers.length; i++) {
      const colIdx = startColIdx + i
      const col = indexToCol(colIdx)
      const cellAddr = `${col}${rowNum}`
      const cell = worksheet[cellAddr] as XLSX.CellObject | undefined

      const value = cell?.v ?? null
      doc[headers[i]] = value

      if (value !== null && value !== undefined) {
        isEmpty = false
      }
    }

    // Skip empty rows if option is set
    if (options?.skipEmptyRows && isEmpty) {
      continue
    }

    // Add row metadata if requested
    if (options?.includeRowMeta) {
      doc._rowNumber = rowNum
      doc._range = `${indexToCol(startColIdx)}${rowNum}:${indexToCol(endColIdx)}${rowNum}`
    }

    rows.push(doc)
  }

  return rows
}

/**
 * Determine if first row should be used as headers (auto mode)
 */
function shouldUseFirstRowAsHeaders(
  worksheet: XLSX.WorkSheet,
  row: number,
  startColIdx: number,
  endColIdx: number
): boolean {
  // Check if all cells in first row are strings
  for (let colIdx = startColIdx; colIdx <= endColIdx; colIdx++) {
    const col = indexToCol(colIdx)
    const cellAddr = `${col}${row}`
    const cell = worksheet[cellAddr] as XLSX.CellObject | undefined

    // If any cell is not a string, don't use as headers
    if (cell && cell.t !== 's') {
      return false
    }
  }

  return true
}
