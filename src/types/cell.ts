/**
 * Cell Types for excel.do
 *
 * Core type definitions for spreadsheet cells, values, and formatting.
 */

/** Cell value types matching Excel/SheetJS semantics */
export type CellValueType = 'boolean' | 'number' | 'string' | 'date' | 'error' | 'empty'

/** Error values that can appear in cells */
export type CellErrorValue =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA'
  | '#SPILL!'
  | '#CALC!'

/** Primitive cell values */
export type CellPrimitive = string | number | boolean | Date | null

/** Cell value with type information */
export interface CellValue {
  /** The raw value */
  v: CellPrimitive
  /** The value type */
  t: CellValueType
  /** Formatted text representation */
  w?: string
  /** Formula if cell contains one */
  f?: string
  /** Error value if type is 'error' */
  e?: CellErrorValue
}

/** Horizontal alignment options */
export type HorizontalAlignment = 'left' | 'center' | 'right' | 'fill' | 'justify'

/** Vertical alignment options */
export type VerticalAlignment = 'top' | 'center' | 'bottom'

/** Border style options */
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'none'

/** RGB color representation */
export interface RGBColor {
  r: number
  g: number
  b: number
  a?: number
}

/** Color can be hex string or RGB object */
export type Color = string | RGBColor

/** Border definition for a single edge */
export interface Border {
  style: BorderStyle
  color?: Color
}

/** All four borders of a cell */
export interface CellBorders {
  top?: Border
  right?: Border
  bottom?: Border
  left?: Border
  diagonal?: Border
}

/** Font styling */
export interface CellFont {
  name?: string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean | 'single' | 'double'
  strikethrough?: boolean
  color?: Color
}

/** Fill/background styling */
export interface CellFill {
  type: 'solid' | 'pattern' | 'gradient'
  color?: Color
  patternType?: string
  fgColor?: Color
  bgColor?: Color
}

/** Text alignment */
export interface CellAlignment {
  horizontal?: HorizontalAlignment
  vertical?: VerticalAlignment
  wrapText?: boolean
  shrinkToFit?: boolean
  textRotation?: number
  indent?: number
}

/** Complete cell format/style */
export interface CellFormat {
  font?: CellFont
  fill?: CellFill
  border?: CellBorders
  alignment?: CellAlignment
  numberFormat?: string
  protection?: {
    locked?: boolean
    hidden?: boolean
  }
}

/** Cell metadata */
export interface CellMetadata {
  /** Cell comment/note */
  comment?: {
    text: string
    author?: string
    visible?: boolean
  }
  /** Hyperlink */
  hyperlink?: {
    target: string
    tooltip?: string
  }
  /** Data validation */
  validation?: {
    type: 'list' | 'number' | 'date' | 'text' | 'custom'
    formula?: string
    values?: string[]
    min?: number | Date
    max?: number | Date
    errorMessage?: string
    errorTitle?: string
  }
}

/** Complete cell object stored in mongo.do */
export interface Cell {
  /** Unique cell identifier: "SheetName!A1" */
  _id: string
  /** Sheet this cell belongs to */
  sheet: string
  /** Row number (1-indexed) */
  row: number
  /** Column letter(s) */
  col: string
  /** Column index (0-indexed) */
  colIndex: number
  /** Cell value */
  value: CellValue
  /** Cell formatting */
  format?: CellFormat
  /** Cell metadata */
  metadata?: CellMetadata
  /** Cells this cell depends on (for formulas) */
  dependencies?: string[]
  /** Cells that depend on this cell */
  dependents?: string[]
  /** Last modified timestamp */
  updatedAt: Date
  /** Created timestamp */
  createdAt: Date
}

/** Options for creating a new cell */
export interface CreateCellOptions {
  sheet: string
  row: number
  col: string
  value?: CellPrimitive
  formula?: string
  format?: CellFormat
  metadata?: CellMetadata
}

/** Options for updating a cell */
export interface UpdateCellOptions {
  value?: CellPrimitive
  formula?: string
  format?: Partial<CellFormat>
  metadata?: Partial<CellMetadata>
}
