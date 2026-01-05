/**
 * Cell Address Utilities Implementation
 *
 * TDD GREEN: Parsing and manipulation of A1/R1C1 cell references.
 */

import type {
  CellReference,
  RangeReference,
  R1C1Reference,
  ParsedReference,
  ReferenceType,
} from '../types'

/** Excel maximum row */
const MAX_ROW = 1048576

/** Excel maximum column index (XFD = 16383) */
const MAX_COL_INDEX = 16383

/**
 * Convert column letter(s) to 0-based index
 * A=0, B=1, ..., Z=25, AA=26, AB=27, ...
 */
export function colToIndex(col: string): number {
  if (!col || col.length === 0) {
    throw new Error('Column name cannot be empty')
  }

  const upper = col.toUpperCase()

  // Validate: must be only letters
  if (!/^[A-Z]+$/.test(upper)) {
    throw new Error(`Invalid column name: ${col}`)
  }

  let index = 0
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64)
  }

  return index - 1 // Convert to 0-based
}

/**
 * Convert 0-based index to column letter(s)
 */
export function indexToCol(index: number): string {
  if (index < 0) {
    throw new Error('Column index cannot be negative')
  }

  let col = ''
  let n = index + 1 // Convert to 1-based for calculation

  while (n > 0) {
    const remainder = (n - 1) % 26
    col = String.fromCharCode(65 + remainder) + col
    n = Math.floor((n - 1) / 26)
  }

  return col
}

/** A1 reference regex - handles both quoted and unquoted sheet names */
const A1_REGEX = /^(?:(?:'([^']*(?:''[^']*)*)'|([A-Za-z0-9_]+))!)?(\$?)([A-Za-z]+)(\$?)(\d+)$/

/**
 * Parse an A1-style cell reference
 */
export function parseA1(ref: string): CellReference {
  if (!ref || ref.trim() === '') {
    throw new Error('Reference cannot be empty')
  }

  const match = ref.match(A1_REGEX)

  if (!match) {
    throw new Error(`Invalid A1 reference: ${ref}`)
  }

  const [, quotedSheet, unquotedSheet, colAbs, colLetters, rowAbs, rowNum] = match

  const col = colLetters.toUpperCase()
  const colIndex = colToIndex(col)
  const row = parseInt(rowNum, 10)

  // Validate bounds
  if (row < 1 || row > MAX_ROW) {
    throw new Error(`Row ${row} out of range (1-${MAX_ROW})`)
  }

  if (colIndex > MAX_COL_INDEX) {
    throw new Error(`Column ${col} exceeds maximum (XFD)`)
  }

  const result: CellReference = {
    original: ref,
    col,
    colIndex,
    row,
    absolute: {
      col: colAbs === '$',
      row: rowAbs === '$',
    },
  }

  if (quotedSheet) {
    // Unescape doubled single quotes
    result.sheet = quotedSheet.replace(/''/g, "'")
  } else if (unquotedSheet) {
    result.sheet = unquotedSheet
  }

  return result
}

/** R1C1 reference regex */
const R1C1_REGEX = /^(?:([^!]+)!)?R(\[?-?\d*\]?)C(\[?-?\d*\]?)$/i

/**
 * Parse an R1C1-style cell reference
 */
export function parseR1C1(ref: string): R1C1Reference {
  if (!ref || ref.trim() === '') {
    throw new Error('Reference cannot be empty')
  }

  const match = ref.match(R1C1_REGEX)

  if (!match) {
    throw new Error(`Invalid R1C1 reference: ${ref}`)
  }

  const [, sheet, rowPart, colPart] = match

  const parseComponent = (part: string): { value: number; relative: boolean } => {
    if (!part || part === '') {
      return { value: 0, relative: true }
    }

    if (part.startsWith('[') && part.endsWith(']')) {
      const inner = part.slice(1, -1)
      return {
        value: inner === '' ? 0 : parseInt(inner, 10),
        relative: true,
      }
    }

    return {
      value: parseInt(part, 10),
      relative: false,
    }
  }

  const result: R1C1Reference = {
    original: ref,
    row: parseComponent(rowPart),
    col: parseComponent(colPart),
  }

  if (sheet) {
    result.sheet = sheet.replace(/^'|'$/g, '').replace(/''/g, "'")
  }

  return result
}

/** Range regex - handles A1:B2 and full column/row references with quoted or unquoted sheet names */
const RANGE_REGEX = /^(?:(?:'([^']*(?:''[^']*)*)'|([A-Za-z0-9_]+))!)?(.+):(.+)$/

/**
 * Parse a range reference (A1:B2)
 */
export function parseRange(ref: string): RangeReference {
  if (!ref || ref.trim() === '') {
    throw new Error('Range reference cannot be empty')
  }

  const match = ref.match(RANGE_REGEX)

  if (!match) {
    throw new Error(`Invalid range reference: ${ref}`)
  }

  const [, quotedSheet, unquotedSheet, startPart, endPart] = match
  const sheetName = quotedSheet ? quotedSheet.replace(/''/g, "'") : unquotedSheet

  let start: CellReference
  let end: CellReference

  // Handle full column reference (A:A)
  if (/^[A-Za-z]+$/.test(startPart) && /^[A-Za-z]+$/.test(endPart)) {
    const startCol = startPart.toUpperCase()
    const endCol = endPart.toUpperCase()
    start = {
      original: startPart,
      col: startCol,
      colIndex: colToIndex(startCol),
      row: 1,
      absolute: { col: false, row: false },
    }
    end = {
      original: endPart,
      col: endCol,
      colIndex: colToIndex(endCol),
      row: MAX_ROW,
      absolute: { col: false, row: false },
    }
  }
  // Handle full row reference (1:1)
  else if (/^\d+$/.test(startPart) && /^\d+$/.test(endPart)) {
    const startRow = parseInt(startPart, 10)
    const endRow = parseInt(endPart, 10)
    start = {
      original: startPart,
      col: 'A',
      colIndex: 0,
      row: startRow,
      absolute: { col: false, row: false },
    }
    end = {
      original: endPart,
      col: 'XFD',
      colIndex: MAX_COL_INDEX,
      row: endRow,
      absolute: { col: false, row: false },
    }
  }
  // Standard cell range
  else {
    start = parseA1(startPart)
    end = parseA1(endPart)
  }

  // Normalize: ensure start <= end
  if (start.row > end.row) {
    [start.row, end.row] = [end.row, start.row]
  }
  if (start.colIndex > end.colIndex) {
    [start.col, end.col] = [end.col, start.col]
    ;[start.colIndex, end.colIndex] = [end.colIndex, start.colIndex]
  }

  const result: RangeReference = {
    original: ref,
    start,
    end,
  }

  if (sheetName) {
    result.sheet = sheetName
    result.start.sheet = sheetName
    result.end.sheet = sheetName
  }

  return result
}

/**
 * Parse any type of reference
 */
export function parseReference(ref: string): ParsedReference {
  try {
    if (ref.includes(':')) {
      return { type: 'range', ref: parseRange(ref) }
    }
    if (/^R(\[?-?\d*\]?)C(\[?-?\d*\]?)$/i.test(ref.replace(/^[^!]+!/, ''))) {
      return { type: 'r1c1', ref: parseR1C1(ref) }
    }
    return { type: 'cell', ref: parseA1(ref) }
  } catch (e) {
    return { type: 'error', message: String(e) }
  }
}

/**
 * Validate an A1-style reference string
 */
export function isValidA1(ref: string): boolean {
  try {
    parseA1(ref)
    return true
  } catch {
    return false
  }
}

/**
 * Validate an R1C1-style reference string
 */
export function isValidR1C1(ref: string): boolean {
  try {
    parseR1C1(ref)
    return true
  } catch {
    return false
  }
}

/**
 * Convert row/col to A1 notation
 */
export function toA1(
  pos: { row: number; col: number },
  options?: {
    absolute?: ReferenceType
    sheet?: string
  }
): string {
  const col = indexToCol(pos.col - 1) // Convert 1-based to 0-based
  const colPrefix = options?.absolute?.col ? '$' : ''
  const rowPrefix = options?.absolute?.row ? '$' : ''

  const cellRef = `${colPrefix}${col}${rowPrefix}${pos.row}`

  if (options?.sheet) {
    const needsQuotes = /[\s']/.test(options.sheet)
    if (needsQuotes) {
      const escaped = options.sheet.replace(/'/g, "''")
      return `'${escaped}'!${cellRef}`
    }
    return `${options.sheet}!${cellRef}`
  }

  return cellRef
}

/**
 * Convert A1 to R1C1 notation
 */
export function toR1C1(
  a1: string,
  options?: { baseCell?: { row: number; col: number } }
): string {
  const ref = parseA1(a1)

  if (!options?.baseCell) {
    return `R${ref.row}C${ref.colIndex + 1}`
  }

  const rowDiff = ref.row - options.baseCell.row
  const colDiff = ref.colIndex + 1 - options.baseCell.col

  if (rowDiff === 0 && colDiff === 0) {
    return 'RC'
  }

  const rowPart = rowDiff === 0 ? 'R' : `R[${rowDiff}]`
  const colPart = colDiff === 0 ? 'C' : `C[${colDiff}]`

  return `${rowPart}${colPart}`
}

/**
 * Offset a cell reference by row/col delta
 */
export function offsetReference(
  ref: CellReference,
  offset: { rowOffset?: number; colOffset?: number }
): CellReference {
  const newRow = ref.absolute.row ? ref.row : ref.row + (offset.rowOffset ?? 0)
  const newColIndex = ref.absolute.col
    ? ref.colIndex
    : ref.colIndex + (offset.colOffset ?? 0)

  if (newRow < 1) {
    throw new Error('Row offset results in row < 1')
  }
  if (newColIndex < 0) {
    throw new Error('Column offset results in negative column')
  }

  const newCol = indexToCol(newColIndex)

  return {
    ...ref,
    row: newRow,
    col: newCol,
    colIndex: newColIndex,
    original: `${ref.absolute.col ? '$' : ''}${newCol}${ref.absolute.row ? '$' : ''}${newRow}`,
  }
}

/**
 * Expand a range to an array of cell references
 */
export function expandRange(rangeStr: string): string[] {
  const range = parseRange(rangeStr)
  const cells: string[] = []

  for (let row = range.start.row; row <= range.end.row; row++) {
    for (let colIdx = range.start.colIndex; colIdx <= range.end.colIndex; colIdx++) {
      const col = indexToCol(colIdx)
      const cellRef = range.sheet ? `${range.sheet}!${col}${row}` : `${col}${row}`
      cells.push(cellRef)
    }
  }

  return cells
}

/**
 * Check if a cell is within a range
 */
export function rangeContains(rangeStr: string, cellStr: string): boolean {
  const range = parseRange(rangeStr)
  const cell = parseA1(cellStr)

  // Sheet must match if either specifies one
  const rangeSheet = range.sheet || null
  const cellSheet = cell.sheet || null

  // If sheets differ (and at least one is specified), not contained
  if (rangeSheet !== cellSheet) {
    return false
  }

  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.colIndex >= range.start.colIndex &&
    cell.colIndex <= range.end.colIndex
  )
}

/**
 * Check if two ranges intersect
 */
export function rangeIntersects(range1Str: string, range2Str: string): boolean {
  const r1 = parseRange(range1Str)
  const r2 = parseRange(range2Str)

  // No intersection if one is entirely to the left/right/above/below the other
  if (r1.end.colIndex < r2.start.colIndex || r2.end.colIndex < r1.start.colIndex) {
    return false
  }
  if (r1.end.row < r2.start.row || r2.end.row < r1.start.row) {
    return false
  }

  return true
}

/**
 * Merge overlapping/adjacent ranges
 */
export function mergeRanges(ranges: string[]): string[] {
  if (ranges.length <= 1) {
    return ranges
  }

  const parsed = ranges.map(parseRange)

  // Sort by start position
  parsed.sort((a, b) => {
    if (a.start.row !== b.start.row) return a.start.row - b.start.row
    return a.start.colIndex - b.start.colIndex
  })

  const merged: RangeReference[] = [parsed[0]]

  for (let i = 1; i < parsed.length; i++) {
    const current = parsed[i]
    const last = merged[merged.length - 1]

    // Check if current can be merged with last
    const canMerge =
      // Same columns, adjacent rows
      (current.start.colIndex === last.start.colIndex &&
        current.end.colIndex === last.end.colIndex &&
        current.start.row <= last.end.row + 1) ||
      // Overlapping
      rangeIntersects(
        `${last.start.col}${last.start.row}:${last.end.col}${last.end.row}`,
        `${current.start.col}${current.start.row}:${current.end.col}${current.end.row}`
      )

    if (canMerge) {
      // Extend last to include current
      last.end.row = Math.max(last.end.row, current.end.row)
      last.end.colIndex = Math.max(last.end.colIndex, current.end.colIndex)
      last.end.col = indexToCol(last.end.colIndex)
    } else {
      merged.push(current)
    }
  }

  return merged.map(
    (r) => `${r.start.col}${r.start.row}:${r.end.col}${r.end.row}`
  )
}
