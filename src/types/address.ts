/**
 * Cell Address Types for excel.do
 *
 * Types for parsing and manipulating cell references like A1, $A$1, R1C1.
 */

/** Cell reference style */
export type ReferenceStyle = 'A1' | 'R1C1'

/** Whether a reference component is absolute or relative */
export interface ReferenceType {
  row: boolean  // true = absolute ($), false = relative
  col: boolean  // true = absolute ($), false = relative
}

/** Parsed A1-style cell reference */
export interface CellReference {
  /** Original string (e.g., "$A$1") */
  original: string
  /** Column letter(s) */
  col: string
  /** Column index (0-based) */
  colIndex: number
  /** Row number (1-based) */
  row: number
  /** Whether components are absolute */
  absolute: ReferenceType
  /** Sheet name if specified (e.g., "Sheet1!A1") */
  sheet?: string
}

/** Parsed R1C1-style cell reference */
export interface R1C1Reference {
  /** Original string (e.g., "R1C1" or "R[1]C[2]") */
  original: string
  /** Row component */
  row: {
    value: number
    relative: boolean
  }
  /** Column component */
  col: {
    value: number
    relative: boolean
  }
  /** Sheet name if specified */
  sheet?: string
}

/** Range reference (e.g., A1:B10) */
export interface RangeReference {
  /** Original string */
  original: string
  /** Start cell */
  start: CellReference
  /** End cell */
  end: CellReference
  /** Sheet name if specified */
  sheet?: string
}

/** Named range definition */
export interface NamedRange {
  /** Name of the range */
  name: string
  /** Reference formula (e.g., "Sheet1!$A$1:$B$10") */
  reference: string
  /** Parsed start cell */
  start: CellReference
  /** Parsed end cell */
  end: CellReference
  /** Scope: workbook-level or sheet-specific */
  scope: 'workbook' | string
  /** Comment/description */
  comment?: string
}

/** Result of parsing any cell reference */
export type ParsedReference =
  | { type: 'cell'; ref: CellReference }
  | { type: 'range'; ref: RangeReference }
  | { type: 'named'; ref: NamedRange }
  | { type: 'r1c1'; ref: R1C1Reference }
  | { type: 'error'; message: string }

/** Options for converting references */
export interface ConvertOptions {
  /** Base cell for relative reference conversion */
  baseCell?: { row: number; col: number }
  /** Target reference style */
  targetStyle?: ReferenceStyle
  /** Whether to make absolute */
  makeAbsolute?: boolean
}
