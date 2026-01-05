/**
 * Cell Data Model Implementation
 *
 * TDD GREEN: Implement functions to make tests pass.
 */

import type {
  Cell,
  CellValue,
  CellValueType,
  CellFormat,
  CellPrimitive,
  CellErrorValue,
  CreateCellOptions,
  UpdateCellOptions,
} from '../types'
import { colToIndex } from './address'

/** All valid Excel error values */
const ERROR_VALUES: CellErrorValue[] = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
]

/** Excel maximum row number */
const MAX_ROW = 1048576

/** Excel maximum column (XFD = 16383) */
const MAX_COL = 16383

/**
 * Infer the type of a cell value
 */
export function inferValueType(value: unknown): CellValueType {
  if (value === null || value === undefined) {
    return 'empty'
  }
  if (typeof value === 'boolean') {
    return 'boolean'
  }
  if (typeof value === 'number') {
    return 'number'
  }
  if (value instanceof Date) {
    return 'date'
  }
  if (typeof value === 'string') {
    if (ERROR_VALUES.includes(value as CellErrorValue)) {
      return 'error'
    }
    return 'string'
  }
  return 'string'
}

/**
 * Create a CellValue from a primitive
 */
export function createCellValue(
  value: CellPrimitive | undefined,
  formula?: string
): CellValue {
  const v = value === undefined ? null : value
  const t = inferValueType(v)

  const cellValue: CellValue = { v, t }

  if (formula) {
    cellValue.f = formula
  }

  if (t === 'error' && typeof v === 'string') {
    cellValue.e = v as CellErrorValue
  }

  return cellValue
}

/**
 * Generate a cell ID from sheet, column, and row
 */
export function getCellId(sheet: string, col: string, row: number): string {
  // Check if sheet name needs quoting
  const needsQuotes = /[\s']/.test(sheet)

  if (needsQuotes) {
    // Escape single quotes by doubling them
    const escapedSheet = sheet.replace(/'/g, "''")
    return `'${escapedSheet}'!${col}${row}`
  }

  return `${sheet}!${col}${row}`
}

/**
 * Create a new Cell object
 */
export function createCell(options: CreateCellOptions): Cell {
  const { sheet, row, col, value, formula, format, metadata } = options

  const colIndex = colToIndex(col)
  const _id = getCellId(sheet, col, row)
  const now = new Date()

  const cellValue = formula
    ? createCellValue(value ?? null, formula)
    : createCellValue(value ?? null)

  const cell: Cell = {
    _id,
    sheet,
    row,
    col,
    colIndex,
    value: cellValue,
    createdAt: now,
    updatedAt: now,
  }

  if (format) {
    cell.format = format
  }

  if (metadata) {
    cell.metadata = metadata
  }

  return cell
}

/** Validation result */
export interface ValidationResult {
  valid: boolean
  errors: string[]
}

/**
 * Validate a cell object
 */
export function validateCell(cell: Cell): ValidationResult {
  const errors: string[] = []

  if (!cell.sheet || cell.sheet.trim() === '') {
    errors.push('Sheet name cannot be empty')
  }

  if (cell.row < 1) {
    errors.push('Row must be >= 1')
  }

  if (cell.row > MAX_ROW) {
    errors.push(`Row exceeds maximum (${MAX_ROW})`)
  }

  if (!/^[A-Za-z]+$/.test(cell.col)) {
    errors.push('Invalid column name')
  }

  if (cell.colIndex < 0 || cell.colIndex > MAX_COL) {
    errors.push(`Column index out of range (0-${MAX_COL})`)
  }

  return {
    valid: errors.length === 0,
    errors,
  }
}

/**
 * Check if a cell is empty
 */
export function isEmptyCell(cell: Cell): boolean {
  // Cell with formula is not empty
  if (cell.value.f) {
    return false
  }

  const v = cell.value.v

  // null is empty
  if (v === null) {
    return true
  }

  // Empty string is empty
  if (typeof v === 'string' && v === '') {
    return true
  }

  // Everything else (including 0, false) is not empty
  return false
}

/**
 * Create a deep clone of a cell
 */
export function cloneCell(
  cell: Cell,
  overrides?: Partial<CreateCellOptions>
): Cell {
  const newRow = overrides?.row ?? cell.row
  const newCol = overrides?.col ?? cell.col
  const newSheet = overrides?.sheet ?? cell.sheet

  const cloned: Cell = {
    ...cell,
    _id: getCellId(newSheet, newCol, newRow),
    sheet: newSheet,
    row: newRow,
    col: newCol,
    colIndex: colToIndex(newCol),
    value: { ...cell.value },
    createdAt: new Date(cell.createdAt),
    updatedAt: new Date(),
  }

  if (overrides?.value !== undefined) {
    cloned.value = createCellValue(overrides.value)
  }

  if (cell.format) {
    cloned.format = JSON.parse(JSON.stringify(cell.format))
  }

  if (cell.metadata) {
    cloned.metadata = JSON.parse(JSON.stringify(cell.metadata))
  }

  if (cell.dependencies) {
    cloned.dependencies = [...cell.dependencies]
  }

  if (cell.dependents) {
    cloned.dependents = [...cell.dependents]
  }

  return cloned
}

/**
 * Deep merge two cell formats
 */
export function mergeCellFormat(
  base: CellFormat | undefined,
  override: CellFormat | undefined
): CellFormat {
  if (!base && !override) {
    return {}
  }
  if (!base) {
    return { ...override }
  }
  if (!override) {
    return { ...base }
  }

  const merged: CellFormat = {}

  // Merge font
  if (base.font || override.font) {
    merged.font = { ...base.font, ...override.font }
  }

  // Merge fill
  if (base.fill || override.fill) {
    merged.fill = { ...base.fill, ...override.fill }
  }

  // Merge border
  if (base.border || override.border) {
    merged.border = { ...base.border, ...override.border }
  }

  // Merge alignment
  if (base.alignment || override.alignment) {
    merged.alignment = { ...base.alignment, ...override.alignment }
  }

  // Number format (override wins)
  if (override.numberFormat !== undefined) {
    merged.numberFormat = override.numberFormat
  } else if (base.numberFormat !== undefined) {
    merged.numberFormat = base.numberFormat
  }

  // Protection
  if (base.protection || override.protection) {
    merged.protection = { ...base.protection, ...override.protection }
  }

  return merged
}

/**
 * Format a cell value as a string for display
 */
export function formatCellValue(
  value: CellValue,
  numberFormat?: string
): string {
  if (value.t === 'empty' || value.v === null) {
    return ''
  }

  if (value.t === 'error') {
    return value.e || String(value.v)
  }

  if (value.t === 'boolean') {
    return value.v ? 'TRUE' : 'FALSE'
  }

  if (value.t === 'string') {
    return String(value.v)
  }

  if (value.t === 'date' && value.v instanceof Date) {
    if (numberFormat === 'yyyy-mm-dd') {
      const d = value.v
      // Use UTC methods to avoid timezone issues
      const year = d.getUTCFullYear()
      const month = String(d.getUTCMonth() + 1).padStart(2, '0')
      const day = String(d.getUTCDate()).padStart(2, '0')
      return `${year}-${month}-${day}`
    }
    return value.v.toLocaleDateString()
  }

  if (value.t === 'number' && typeof value.v === 'number') {
    if (!numberFormat) {
      return String(value.v)
    }

    // Handle percentage format
    if (numberFormat === '0%') {
      return `${Math.round(value.v * 100)}%`
    }

    // Handle currency format
    if (numberFormat === '$#,##0.00') {
      return `$${value.v.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })}`
    }

    return String(value.v)
  }

  return String(value.v)
}
