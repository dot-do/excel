/**
 * XLSX Export Tests (RED)
 *
 * TDD: These tests define expected behavior for SheetJS XLSX export functionality.
 * All tests should FAIL initially until implementation is complete.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import {
  toWorkbook,
  toSheet,
  cellToSheetJS,
  toXLSX,
  toCSV,
} from './xlsx'
import type {
  Cell,
  CellValue,
  CellFormat,
  CellFont,
  CellFill,
  CellBorders,
  CellAlignment,
} from '../types'
import type { WorkBook, WorkSheet, CellObject } from 'xlsx'

// ============================================================================
// Test Helpers
// ============================================================================

function createTestCell(overrides: Partial<Cell> = {}): Cell {
  const now = new Date()
  return {
    _id: 'Sheet1!A1',
    sheet: 'Sheet1',
    row: 1,
    col: 'A',
    colIndex: 0,
    value: { v: 'test', t: 'string' },
    createdAt: now,
    updatedAt: now,
    ...overrides,
  }
}

function createTestCellValue(value: unknown, formula?: string): CellValue {
  if (value === null || value === undefined) {
    return { v: null, t: 'empty', f: formula }
  }
  if (typeof value === 'string') {
    if (value.startsWith('#') && value.endsWith('!')) {
      return { v: value, t: 'error', e: value as any, f: formula }
    }
    return { v: value, t: 'string', f: formula }
  }
  if (typeof value === 'number') {
    return { v: value, t: 'number', f: formula }
  }
  if (typeof value === 'boolean') {
    return { v: value, t: 'boolean', f: formula }
  }
  if (value instanceof Date) {
    return { v: value, t: 'date', f: formula }
  }
  return { v: null, t: 'empty', f: formula }
}

// ============================================================================
// Types for export options
// ============================================================================

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
}

export interface SheetData {
  name: string
  cells: Cell[]
  columnWidths?: Record<string, number>
  rowHeights?: Record<number, number>
}

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
// cellToSheetJS Tests
// ============================================================================

describe('XLSX Export', () => {
  describe('cellToSheetJS', () => {
    describe('value conversion', () => {
      it('should convert string cell to SheetJS cell', () => {
        const cell = createTestCell({
          value: createTestCellValue('Hello World'),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('s')
        expect(result.v).toBe('Hello World')
      })

      it('should convert number cell to SheetJS cell', () => {
        const cell = createTestCell({
          value: createTestCellValue(42.5),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('n')
        expect(result.v).toBe(42.5)
      })

      it('should convert boolean cell to SheetJS cell', () => {
        const cell = createTestCell({
          value: createTestCellValue(true),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('b')
        expect(result.v).toBe(true)
      })

      it('should convert false boolean to SheetJS cell', () => {
        const cell = createTestCell({
          value: createTestCellValue(false),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('b')
        expect(result.v).toBe(false)
      })

      it('should convert date cell to SheetJS cell', () => {
        const date = new Date('2024-06-15T10:30:00Z')
        const cell = createTestCell({
          value: createTestCellValue(date),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('d')
        expect(result.v).toEqual(date)
      })

      it('should convert empty cell to SheetJS cell', () => {
        const cell = createTestCell({
          value: createTestCellValue(null),
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('z')
        expect(result.v).toBeUndefined()
      })

      it('should convert error cell to SheetJS cell', () => {
        const cell = createTestCell({
          value: { v: '#DIV/0!', t: 'error', e: '#DIV/0!' },
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('e')
        expect(result.v).toBe('#DIV/0!')
      })

      it('should handle #N/A error', () => {
        const cell = createTestCell({
          value: { v: '#N/A', t: 'error', e: '#N/A' },
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('e')
        expect(result.v).toBe('#N/A')
      })

      it('should handle #REF! error', () => {
        const cell = createTestCell({
          value: { v: '#REF!', t: 'error', e: '#REF!' },
        })
        const result = cellToSheetJS(cell)

        expect(result.t).toBe('e')
        expect(result.v).toBe('#REF!')
      })
    })

    describe('formula preservation', () => {
      it('should preserve formula in cell', () => {
        const cell = createTestCell({
          value: { v: 100, t: 'number', f: '=A1+B1' },
        })
        const result = cellToSheetJS(cell)

        expect(result.f).toBe('A1+B1') // SheetJS strips leading =
        expect(result.v).toBe(100)
      })

      it('should preserve formula with cell references', () => {
        const cell = createTestCell({
          value: { v: 42, t: 'number', f: '=SUM(A1:A10)' },
        })
        const result = cellToSheetJS(cell)

        expect(result.f).toBe('SUM(A1:A10)')
      })

      it('should preserve formula with absolute references', () => {
        const cell = createTestCell({
          value: { v: 10, t: 'number', f: '=$A$1*B2' },
        })
        const result = cellToSheetJS(cell)

        expect(result.f).toBe('$A$1*B2')
      })

      it('should preserve formula with sheet references', () => {
        const cell = createTestCell({
          value: { v: 50, t: 'number', f: "=Sheet2!A1+'Other Sheet'!B2" },
        })
        const result = cellToSheetJS(cell)

        expect(result.f).toBe("Sheet2!A1+'Other Sheet'!B2")
      })

      it('should handle formula-only cell with no computed value', () => {
        const cell = createTestCell({
          value: { v: null, t: 'empty', f: '=NOW()' },
        })
        const result = cellToSheetJS(cell)

        expect(result.f).toBe('NOW()')
        expect(result.t).toBe('z')
      })
    })

    describe('formatted text', () => {
      it('should include formatted text representation', () => {
        const cell = createTestCell({
          value: { v: 0.75, t: 'number', w: '75%' },
        })
        const result = cellToSheetJS(cell)

        expect(result.w).toBe('75%')
      })

      it('should include formatted currency', () => {
        const cell = createTestCell({
          value: { v: 1234.56, t: 'number', w: '$1,234.56' },
        })
        const result = cellToSheetJS(cell)

        expect(result.w).toBe('$1,234.56')
      })
    })

    describe('style conversion', () => {
      it('should convert font styling to SheetJS style', () => {
        const cell = createTestCell({
          value: createTestCellValue('Bold Text'),
          format: {
            font: {
              bold: true,
              italic: true,
              size: 14,
              name: 'Arial',
              color: '#FF0000',
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s).toBeDefined()
        expect(result.s?.font?.bold).toBe(true)
        expect(result.s?.font?.italic).toBe(true)
        expect(result.s?.font?.sz).toBe(14)
        expect(result.s?.font?.name).toBe('Arial')
        expect(result.s?.font?.color?.rgb).toBe('FF0000')
      })

      it('should convert underline style', () => {
        const cell = createTestCell({
          value: createTestCellValue('Underlined'),
          format: {
            font: { underline: true },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.font?.underline).toBe(true)
      })

      it('should convert double underline style', () => {
        const cell = createTestCell({
          value: createTestCellValue('Double Underlined'),
          format: {
            font: { underline: 'double' },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.font?.underline).toBe(2)
      })

      it('should convert strikethrough style', () => {
        const cell = createTestCell({
          value: createTestCellValue('Strikethrough'),
          format: {
            font: { strikethrough: true },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.font?.strike).toBe(true)
      })

      it('should convert solid fill to SheetJS style', () => {
        const cell = createTestCell({
          value: createTestCellValue('Yellow Background'),
          format: {
            fill: {
              type: 'solid',
              color: '#FFFF00',
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.fill?.fgColor?.rgb).toBe('FFFF00')
        expect(result.s?.fill?.patternType).toBe('solid')
      })

      it('should convert pattern fill', () => {
        const cell = createTestCell({
          value: createTestCellValue('Pattern'),
          format: {
            fill: {
              type: 'pattern',
              patternType: 'gray125',
              fgColor: '#000000',
              bgColor: '#FFFFFF',
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.fill?.patternType).toBe('gray125')
        expect(result.s?.fill?.fgColor?.rgb).toBe('000000')
        expect(result.s?.fill?.bgColor?.rgb).toBe('FFFFFF')
      })

      it('should convert border styling', () => {
        const cell = createTestCell({
          value: createTestCellValue('Bordered'),
          format: {
            border: {
              top: { style: 'thin', color: '#000000' },
              right: { style: 'medium', color: '#FF0000' },
              bottom: { style: 'thick', color: '#00FF00' },
              left: { style: 'dashed', color: '#0000FF' },
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.border?.top?.style).toBe('thin')
        expect(result.s?.border?.top?.color?.rgb).toBe('000000')
        expect(result.s?.border?.right?.style).toBe('medium')
        expect(result.s?.border?.bottom?.style).toBe('thick')
        expect(result.s?.border?.left?.style).toBe('dashed')
      })

      it('should convert alignment', () => {
        const cell = createTestCell({
          value: createTestCellValue('Centered'),
          format: {
            alignment: {
              horizontal: 'center',
              vertical: 'center',
              wrapText: true,
              textRotation: 45,
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.alignment?.horizontal).toBe('center')
        expect(result.s?.alignment?.vertical).toBe('center')
        expect(result.s?.alignment?.wrapText).toBe(true)
        expect(result.s?.alignment?.textRotation).toBe(45)
      })

      it('should convert shrinkToFit alignment', () => {
        const cell = createTestCell({
          value: createTestCellValue('Shrink'),
          format: {
            alignment: { shrinkToFit: true },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.alignment?.shrinkToFit).toBe(true)
      })

      it('should convert indent alignment', () => {
        const cell = createTestCell({
          value: createTestCellValue('Indented'),
          format: {
            alignment: { indent: 2 },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.alignment?.indent).toBe(2)
      })

      it('should convert number format', () => {
        const cell = createTestCell({
          value: createTestCellValue(0.5),
          format: {
            numberFormat: '0.00%',
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.z).toBe('0.00%')
      })

      it('should skip styles when includeFormatting is false', () => {
        const cell = createTestCell({
          value: createTestCellValue('No Style'),
          format: {
            font: { bold: true },
            fill: { type: 'solid', color: '#FF0000' },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: false })

        expect(result.s).toBeUndefined()
      })

      it('should handle RGB color objects', () => {
        const cell = createTestCell({
          value: createTestCellValue('RGB Color'),
          format: {
            font: {
              color: { r: 255, g: 128, b: 64 },
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        expect(result.s?.font?.color?.rgb).toBe('FF8040')
      })

      it('should handle RGB color with alpha', () => {
        const cell = createTestCell({
          value: createTestCellValue('RGBA Color'),
          format: {
            font: {
              color: { r: 255, g: 0, b: 0, a: 0.5 },
            },
          },
        })
        const result = cellToSheetJS(cell, { includeFormatting: true })

        // Alpha channel should be converted to ARGB format
        expect(result.s?.font?.color?.rgb).toMatch(/^[0-9A-F]{6,8}$/i)
      })
    })

    describe('hyperlink conversion', () => {
      it('should include hyperlink in cell', () => {
        const cell = createTestCell({
          value: createTestCellValue('Click here'),
          metadata: {
            hyperlink: {
              target: 'https://example.com',
              tooltip: 'Visit Example',
            },
          },
        })
        const result = cellToSheetJS(cell)

        expect(result.l).toBeDefined()
        expect(result.l?.Target).toBe('https://example.com')
        expect(result.l?.Tooltip).toBe('Visit Example')
      })

      it('should handle hyperlink without tooltip', () => {
        const cell = createTestCell({
          value: createTestCellValue('Link'),
          metadata: {
            hyperlink: {
              target: 'https://example.com',
            },
          },
        })
        const result = cellToSheetJS(cell)

        expect(result.l?.Target).toBe('https://example.com')
        expect(result.l?.Tooltip).toBeUndefined()
      })
    })

    describe('comment conversion', () => {
      it('should include comment in cell', () => {
        const cell = createTestCell({
          value: createTestCellValue('With Comment'),
          metadata: {
            comment: {
              text: 'This is a note',
              author: 'John Doe',
            },
          },
        })
        const result = cellToSheetJS(cell)

        expect(result.c).toBeDefined()
        expect(result.c?.[0]?.t).toBe('This is a note')
        expect(result.c?.[0]?.a).toBe('John Doe')
      })

      it('should handle comment without author', () => {
        const cell = createTestCell({
          value: createTestCellValue('Comment'),
          metadata: {
            comment: {
              text: 'Anonymous note',
            },
          },
        })
        const result = cellToSheetJS(cell)

        expect(result.c?.[0]?.t).toBe('Anonymous note')
        expect(result.c?.[0]?.a).toBeUndefined()
      })
    })
  })

  // ============================================================================
  // toSheet Tests
  // ============================================================================

  describe('toSheet', () => {
    describe('basic conversion', () => {
      it('should create empty worksheet', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [],
        }
        const result = toSheet(sheetData)

        expect(result).toBeDefined()
        expect(result['!ref']).toBeUndefined() // Empty sheet has no ref
      })

      it('should create worksheet from single cell', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Hello'),
            }),
          ],
        }
        const result = toSheet(sheetData)

        expect(result['!ref']).toBe('A1:A1')
        expect(result['A1']).toBeDefined()
        expect(result['A1'].v).toBe('Hello')
      })

      it('should create worksheet from multiple cells', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('A1'),
            }),
            createTestCell({
              _id: 'Sheet1!B2',
              row: 2,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue('B2'),
            }),
            createTestCell({
              _id: 'Sheet1!C3',
              row: 3,
              col: 'C',
              colIndex: 2,
              value: createTestCellValue('C3'),
            }),
          ],
        }
        const result = toSheet(sheetData)

        expect(result['!ref']).toBe('A1:C3')
        expect(result['A1'].v).toBe('A1')
        expect(result['B2'].v).toBe('B2')
        expect(result['C3'].v).toBe('C3')
      })

      it('should handle sparse cells correctly', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Start'),
            }),
            createTestCell({
              _id: 'Sheet1!Z100',
              row: 100,
              col: 'Z',
              colIndex: 25,
              value: createTestCellValue('End'),
            }),
          ],
        }
        const result = toSheet(sheetData)

        expect(result['!ref']).toBe('A1:Z100')
        expect(result['A1'].v).toBe('Start')
        expect(result['Z100'].v).toBe('End')
        expect(result['B2']).toBeUndefined() // Sparse - no cell at B2
      })

      it('should handle multi-letter columns', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!AA1',
              row: 1,
              col: 'AA',
              colIndex: 26,
              value: createTestCellValue('Column AA'),
            }),
          ],
        }
        const result = toSheet(sheetData)

        expect(result['AA1']).toBeDefined()
        expect(result['AA1'].v).toBe('Column AA')
      })
    })

    describe('column widths', () => {
      it('should set column widths', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Wide'),
            }),
          ],
          columnWidths: { A: 20, B: 15 },
        }
        const result = toSheet(sheetData)

        expect(result['!cols']).toBeDefined()
        expect(result['!cols']?.[0]?.wch).toBe(20)
        expect(result['!cols']?.[1]?.wch).toBe(15)
      })

      it('should handle column widths with options', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Test'),
            }),
          ],
        }
        const options: ExportOptions = {
          columnWidths: { A: 25, C: 30 },
          defaultColumnWidth: 10,
        }
        const result = toSheet(sheetData, options)

        expect(result['!cols']?.[0]?.wch).toBe(25)
        expect(result['!cols']?.[2]?.wch).toBe(30)
      })
    })

    describe('row heights', () => {
      it('should set row heights', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Tall'),
            }),
          ],
          rowHeights: { 1: 30, 2: 40 },
        }
        const result = toSheet(sheetData)

        expect(result['!rows']).toBeDefined()
        expect(result['!rows']?.[0]?.hpt).toBe(30) // Row 1 = index 0
        expect(result['!rows']?.[1]?.hpt).toBe(40) // Row 2 = index 1
      })
    })

    describe('formatting', () => {
      it('should include cell formatting when enabled', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Bold'),
              format: { font: { bold: true } },
            }),
          ],
        }
        const options: ExportOptions = { includeFormatting: true }
        const result = toSheet(sheetData, options)

        expect(result['A1'].s?.font?.bold).toBe(true)
      })

      it('should skip cell formatting when disabled', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Bold'),
              format: { font: { bold: true } },
            }),
          ],
        }
        const options: ExportOptions = { includeFormatting: false }
        const result = toSheet(sheetData, options)

        expect(result['A1'].s).toBeUndefined()
      })
    })

    describe('merged cells', () => {
      it('should handle merged cell ranges', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Merged'),
            }),
          ],
        }
        // Merges are typically stored in sheet metadata
        const merges = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }] // A1:B2
        const result = toSheet(sheetData, { merges } as any)

        expect(result['!merges']).toBeDefined()
        expect(result['!merges']?.[0]).toEqual({ s: { r: 0, c: 0 }, e: { r: 1, c: 1 } })
      })
    })
  })

  // ============================================================================
  // toWorkbook Tests
  // ============================================================================

  describe('toWorkbook', () => {
    describe('single sheet workbook', () => {
      it('should create workbook with one sheet', () => {
        const workbookData: WorkbookData = {
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                createTestCell({
                  _id: 'Sheet1!A1',
                  row: 1,
                  col: 'A',
                  colIndex: 0,
                  value: createTestCellValue('Hello'),
                }),
              ],
            },
          ],
        }
        const result = toWorkbook(workbookData)

        expect(result.SheetNames).toContain('Sheet1')
        expect(result.Sheets['Sheet1']).toBeDefined()
        expect(result.Sheets['Sheet1']['A1'].v).toBe('Hello')
      })

      it('should set active sheet', () => {
        const workbookData: WorkbookData = {
          sheets: [
            {
              name: 'Sheet1',
              cells: [],
            },
          ],
          activeSheet: 'Sheet1',
        }
        const result = toWorkbook(workbookData)

        expect(result.Workbook?.Sheets?.[0]?.Hidden).toBeFalsy()
      })
    })

    describe('multi-sheet workbook', () => {
      it('should create workbook with multiple sheets', () => {
        const workbookData: WorkbookData = {
          sheets: [
            {
              name: 'Data',
              cells: [
                createTestCell({
                  _id: 'Data!A1',
                  sheet: 'Data',
                  row: 1,
                  col: 'A',
                  colIndex: 0,
                  value: createTestCellValue('Data Sheet'),
                }),
              ],
            },
            {
              name: 'Summary',
              cells: [
                createTestCell({
                  _id: 'Summary!A1',
                  sheet: 'Summary',
                  row: 1,
                  col: 'A',
                  colIndex: 0,
                  value: createTestCellValue('Summary Sheet'),
                }),
              ],
            },
          ],
        }
        const result = toWorkbook(workbookData)

        expect(result.SheetNames).toHaveLength(2)
        expect(result.SheetNames).toContain('Data')
        expect(result.SheetNames).toContain('Summary')
        expect(result.Sheets['Data']['A1'].v).toBe('Data Sheet')
        expect(result.Sheets['Summary']['A1'].v).toBe('Summary Sheet')
      })

      it('should preserve sheet order', () => {
        const workbookData: WorkbookData = {
          sheets: [
            { name: 'First', cells: [] },
            { name: 'Second', cells: [] },
            { name: 'Third', cells: [] },
          ],
        }
        const result = toWorkbook(workbookData)

        expect(result.SheetNames[0]).toBe('First')
        expect(result.SheetNames[1]).toBe('Second')
        expect(result.SheetNames[2]).toBe('Third')
      })
    })

    describe('workbook metadata', () => {
      it('should include workbook properties', () => {
        const workbookData: WorkbookData = {
          sheets: [{ name: 'Sheet1', cells: [] }],
          metadata: {
            title: 'Test Workbook',
            author: 'Test Author',
            created: new Date('2024-01-01'),
            modified: new Date('2024-06-15'),
          },
        }
        const result = toWorkbook(workbookData)

        expect(result.Props?.Title).toBe('Test Workbook')
        expect(result.Props?.Author).toBe('Test Author')
        expect(result.Props?.CreatedDate).toEqual(new Date('2024-01-01'))
        expect(result.Props?.ModifiedDate).toEqual(new Date('2024-06-15'))
      })
    })

    describe('export options', () => {
      it('should apply formatting options to all sheets', () => {
        const workbookData: WorkbookData = {
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                createTestCell({
                  _id: 'Sheet1!A1',
                  row: 1,
                  col: 'A',
                  colIndex: 0,
                  value: createTestCellValue('Styled'),
                  format: { font: { bold: true } },
                }),
              ],
            },
          ],
        }
        const options: ExportOptions = { includeFormatting: true }
        const result = toWorkbook(workbookData, options)

        expect(result.Sheets['Sheet1']['A1'].s?.font?.bold).toBe(true)
      })

      it('should apply column widths from options', () => {
        const workbookData: WorkbookData = {
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                createTestCell({
                  _id: 'Sheet1!A1',
                  row: 1,
                  col: 'A',
                  colIndex: 0,
                  value: createTestCellValue('Test'),
                }),
              ],
            },
          ],
        }
        const options: ExportOptions = { columnWidths: { A: 50 } }
        const result = toWorkbook(workbookData, options)

        expect(result.Sheets['Sheet1']['!cols']?.[0]?.wch).toBe(50)
      })
    })
  })

  // ============================================================================
  // toXLSX Tests
  // ============================================================================

  describe('toXLSX', () => {
    it('should generate XLSX buffer from workbook data', async () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue('Export Test'),
              }),
            ],
          },
        ],
      }
      const result = await toXLSX(workbookData)

      expect(result).toBeInstanceOf(Buffer)
      expect(result.length).toBeGreaterThan(0)
      // XLSX files start with PK (zip signature)
      expect(result[0]).toBe(0x50) // 'P'
      expect(result[1]).toBe(0x4b) // 'K'
    })

    it('should generate XLSX with multiple sheets', async () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Data',
            cells: [
              createTestCell({
                _id: 'Data!A1',
                sheet: 'Data',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue(100),
              }),
            ],
          },
          {
            name: 'Summary',
            cells: [
              createTestCell({
                _id: 'Summary!A1',
                sheet: 'Summary',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: { v: 100, t: 'number', f: '=Data!A1' },
              }),
            ],
          },
        ],
      }
      const result = await toXLSX(workbookData)

      expect(result).toBeInstanceOf(Buffer)
      expect(result.length).toBeGreaterThan(0)
    })

    it('should preserve formulas in XLSX output', async () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue(10),
              }),
              createTestCell({
                _id: 'Sheet1!A2',
                row: 2,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue(20),
              }),
              createTestCell({
                _id: 'Sheet1!A3',
                row: 3,
                col: 'A',
                colIndex: 0,
                value: { v: 30, t: 'number', f: '=SUM(A1:A2)' },
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = { preserveFormulas: true }
      const result = await toXLSX(workbookData, options)

      expect(result).toBeInstanceOf(Buffer)
      // The result should contain the formula (we can't easily verify without parsing)
    })

    it('should include formatting in XLSX when enabled', async () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue('Bold'),
                format: {
                  font: { bold: true, size: 16 },
                  fill: { type: 'solid', color: '#FFFF00' },
                },
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = { includeFormatting: true }
      const result = await toXLSX(workbookData, options)

      expect(result).toBeInstanceOf(Buffer)
      expect(result.length).toBeGreaterThan(0)
    })

    it('should handle empty workbook', async () => {
      const workbookData: WorkbookData = {
        sheets: [{ name: 'Empty', cells: [] }],
      }
      const result = await toXLSX(workbookData)

      expect(result).toBeInstanceOf(Buffer)
      expect(result.length).toBeGreaterThan(0)
    })

    it('should handle large dataset', async () => {
      const cells: Cell[] = []
      for (let row = 1; row <= 1000; row++) {
        for (let col = 0; col < 10; col++) {
          const colLetter = String.fromCharCode(65 + col)
          cells.push(
            createTestCell({
              _id: `Sheet1!${colLetter}${row}`,
              row,
              col: colLetter,
              colIndex: col,
              value: createTestCellValue(`R${row}C${col}`),
            })
          )
        }
      }
      const workbookData: WorkbookData = {
        sheets: [{ name: 'Sheet1', cells }],
      }
      const result = await toXLSX(workbookData)

      expect(result).toBeInstanceOf(Buffer)
      expect(result.length).toBeGreaterThan(0)
    })
  })

  // ============================================================================
  // toCSV Tests
  // ============================================================================

  describe('toCSV', () => {
    describe('basic conversion', () => {
      it('should convert single cell to CSV', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Hello'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('Hello')
      })

      it('should convert row to CSV', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('A'),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue('B'),
            }),
            createTestCell({
              _id: 'Sheet1!C1',
              row: 1,
              col: 'C',
              colIndex: 2,
              value: createTestCellValue('C'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('A,B,C')
      })

      it('should convert multiple rows to CSV', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Name'),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue('Age'),
            }),
            createTestCell({
              _id: 'Sheet1!A2',
              row: 2,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('John'),
            }),
            createTestCell({
              _id: 'Sheet1!B2',
              row: 2,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue(30),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('Name,Age\nJohn,30')
      })

      it('should handle empty cells with commas', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('First'),
            }),
            createTestCell({
              _id: 'Sheet1!C1',
              row: 1,
              col: 'C',
              colIndex: 2,
              value: createTestCellValue('Third'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('First,,Third')
      })
    })

    describe('value formatting', () => {
      it('should format numbers correctly', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue(42.5),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('42.5')
      })

      it('should format booleans as TRUE/FALSE', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue(true),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue(false),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('TRUE,FALSE')
      })

      it('should format dates in ISO format by default', () => {
        const date = new Date('2024-06-15T10:30:00Z')
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue(date),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toMatch(/2024-06-15/)
      })

      it('should use formatted text when available', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: { v: 0.75, t: 'number', w: '75%' },
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('75%')
      })
    })

    describe('escaping', () => {
      it('should quote strings containing commas', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Hello, World'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('"Hello, World"')
      })

      it('should quote strings containing quotes and escape inner quotes', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Say "Hello"'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('"Say ""Hello"""')
      })

      it('should quote strings containing newlines', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Line1\nLine2'),
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('"Line1\nLine2"')
      })
    })

    describe('options', () => {
      it('should use custom delimiter', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('A'),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue('B'),
            }),
          ],
        }
        const options: ExportOptions = { csvDelimiter: ';' }
        const result = toCSV(sheetData, options)

        expect(result).toBe('A;B')
      })

      it('should use custom row separator', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Row1'),
            }),
            createTestCell({
              _id: 'Sheet1!A2',
              row: 2,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Row2'),
            }),
          ],
        }
        const options: ExportOptions = { csvRowSeparator: '\r\n' }
        const result = toCSV(sheetData, options)

        expect(result).toBe('Row1\r\nRow2')
      })

      it('should quote all fields when csvQuoteAll is true', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('Simple'),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue(123),
            }),
          ],
        }
        const options: ExportOptions = { csvQuoteAll: true }
        const result = toCSV(sheetData, options)

        expect(result).toBe('"Simple","123"')
      })

      it('should add header row when includeHeaders is true', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: createTestCellValue('John'),
            }),
            createTestCell({
              _id: 'Sheet1!B1',
              row: 1,
              col: 'B',
              colIndex: 1,
              value: createTestCellValue(30),
            }),
          ],
        }
        const options: ExportOptions = {
          includeHeaders: true,
          headerNames: { A: 'Name', B: 'Age' },
        }
        const result = toCSV(sheetData, options)

        expect(result).toBe('Name,Age\nJohn,30')
      })
    })

    describe('formula handling', () => {
      it('should export formula results, not formulas', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [
            createTestCell({
              _id: 'Sheet1!A1',
              row: 1,
              col: 'A',
              colIndex: 0,
              value: { v: 100, t: 'number', f: '=50+50' },
            }),
          ],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('100')
        expect(result).not.toContain('=')
      })
    })

    describe('empty workbook', () => {
      it('should return empty string for empty sheet', () => {
        const sheetData: SheetData = {
          name: 'Sheet1',
          cells: [],
        }
        const result = toCSV(sheetData)

        expect(result).toBe('')
      })
    })
  })

  // ============================================================================
  // Export Options Integration Tests
  // ============================================================================

  describe('Export Options Integration', () => {
    it('should apply default column width to all columns', () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue('Test'),
              }),
              createTestCell({
                _id: 'Sheet1!B1',
                row: 1,
                col: 'B',
                colIndex: 1,
                value: createTestCellValue('Test'),
              }),
              createTestCell({
                _id: 'Sheet1!C1',
                row: 1,
                col: 'C',
                colIndex: 2,
                value: createTestCellValue('Test'),
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = { defaultColumnWidth: 15 }
      const result = toWorkbook(workbookData, options)

      const cols = result.Sheets['Sheet1']['!cols']
      expect(cols?.[0]?.wch).toBe(15)
      expect(cols?.[1]?.wch).toBe(15)
      expect(cols?.[2]?.wch).toBe(15)
    })

    it('should override default column width with specific widths', () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue('Test'),
              }),
              createTestCell({
                _id: 'Sheet1!B1',
                row: 1,
                col: 'B',
                colIndex: 1,
                value: createTestCellValue('Test'),
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = {
        defaultColumnWidth: 10,
        columnWidths: { A: 25 },
      }
      const result = toWorkbook(workbookData, options)

      const cols = result.Sheets['Sheet1']['!cols']
      expect(cols?.[0]?.wch).toBe(25) // Specific width
      expect(cols?.[1]?.wch).toBe(10) // Default width
    })

    it('should apply date format to date cells', () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue(new Date('2024-01-15')),
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = {
        dateFormat: 'yyyy-mm-dd',
        includeFormatting: true,
      }
      const result = toWorkbook(workbookData, options)

      expect(result.Sheets['Sheet1']['A1'].z).toBe('yyyy-mm-dd')
    })

    it('should apply number format to number cells', () => {
      const workbookData: WorkbookData = {
        sheets: [
          {
            name: 'Sheet1',
            cells: [
              createTestCell({
                _id: 'Sheet1!A1',
                row: 1,
                col: 'A',
                colIndex: 0,
                value: createTestCellValue(1234.5),
              }),
            ],
          },
        ],
      }
      const options: ExportOptions = {
        numberFormat: '#,##0.00',
        includeFormatting: true,
      }
      const result = toWorkbook(workbookData, options)

      expect(result.Sheets['Sheet1']['A1'].z).toBe('#,##0.00')
    })
  })
})
