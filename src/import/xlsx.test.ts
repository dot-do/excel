/**
 * XLSX Import Tests (RED)
 *
 * TDD: These tests define expected behavior for SheetJS XLSX import functionality.
 * All tests should FAIL initially until implementation is complete.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import * as XLSX from 'xlsx'
import {
  parseWorkbook,
  parseSheet,
  cellFromSheetJS,
  rowsFromSheet,
} from './xlsx'
import type {
  Cell,
  CellValue,
  CellFormat,
  CellValueType,
} from '../types'
import type {
  ImportOptions,
  ParsedWorkbook,
  ParsedSheet,
  RowDocument,
} from './xlsx'

// Helper to create a mock SheetJS workbook
function createMockWorkbook(sheets: Record<string, XLSX.WorkSheet>): XLSX.WorkBook {
  return {
    SheetNames: Object.keys(sheets),
    Sheets: sheets,
  }
}

// Helper to create a mock SheetJS worksheet
function createMockSheet(data: (string | number | boolean | null | Date)[][]): XLSX.WorkSheet {
  return XLSX.utils.aoa_to_sheet(data)
}

describe('XLSX Import', () => {
  describe('parseWorkbook', () => {
    it('should parse an XLSX buffer to internal workbook structure', () => {
      // Create a simple workbook buffer
      const wb = XLSX.utils.book_new()
      const ws = XLSX.utils.aoa_to_sheet([
        ['Name', 'Age'],
        ['Alice', 30],
        ['Bob', 25],
      ])
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

      const result = parseWorkbook(buffer)

      expect(result.name).toBeDefined()
      expect(result.sheets).toHaveLength(1)
      expect(result.sheets[0].name).toBe('Sheet1')
    })

    it('should parse multiple sheets', () => {
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['A']]), 'Sheet1')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['B']]), 'Sheet2')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['C']]), 'Data')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

      const result = parseWorkbook(buffer)

      expect(result.sheets).toHaveLength(3)
      expect(result.sheets.map(s => s.name)).toEqual(['Sheet1', 'Sheet2', 'Data'])
    })

    it('should extract workbook metadata', () => {
      const wb = XLSX.utils.book_new()
      wb.Props = {
        Title: 'Test Workbook',
        Author: 'Test Author',
        CreatedDate: new Date('2024-01-15'),
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['A']]), 'Sheet1')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

      const result = parseWorkbook(buffer)

      expect(result.metadata?.title).toBe('Test Workbook')
      expect(result.metadata?.author).toBe('Test Author')
    })

    it('should handle empty workbook', () => {
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Empty')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

      const result = parseWorkbook(buffer)

      expect(result.sheets).toHaveLength(1)
      expect(result.sheets[0].cells).toHaveLength(0)
    })

    it('should respect sheet filtering option', () => {
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['A']]), 'Include')
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['B']]), 'Exclude')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })

      const result = parseWorkbook(buffer, { sheets: ['Include'] })

      expect(result.sheets).toHaveLength(1)
      expect(result.sheets[0].name).toBe('Include')
    })

    it('should throw on invalid buffer', () => {
      const invalidBuffer = Buffer.from('not a valid xlsx file')

      expect(() => parseWorkbook(invalidBuffer)).toThrow()
    })

    it('should parse Uint8Array input', () => {
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Test']]), 'Sheet1')
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
      const uint8Array = new Uint8Array(buffer)

      const result = parseWorkbook(uint8Array)

      expect(result.sheets).toHaveLength(1)
    })
  })

  describe('parseSheet', () => {
    it('should parse a SheetJS worksheet to internal sheet structure', () => {
      const ws = createMockSheet([
        ['Name', 'Age', 'City'],
        ['Alice', 30, 'NYC'],
        ['Bob', 25, 'LA'],
      ])

      const result = parseSheet(ws, 'TestSheet')

      expect(result.name).toBe('TestSheet')
      expect(result.cells.length).toBeGreaterThan(0)
    })

    it('should extract sheet dimensions', () => {
      const ws = createMockSheet([
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ])

      const result = parseSheet(ws, 'Sheet1')

      expect(result.dimensions).toBeDefined()
      expect(result.dimensions.startRow).toBe(1)
      expect(result.dimensions.endRow).toBe(2)
      expect(result.dimensions.startCol).toBe('A')
      expect(result.dimensions.endCol).toBe('C')
    })

    it('should handle sparse sheets with gaps', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 'A1', t: 's' },
        C3: { v: 'C3', t: 's' },
        '!ref': 'A1:C3',
      }

      const result = parseSheet(ws, 'Sparse')

      expect(result.cells).toHaveLength(2)
      expect(result.cells.find(c => c.col === 'A' && c.row === 1)).toBeDefined()
      expect(result.cells.find(c => c.col === 'C' && c.row === 3)).toBeDefined()
    })

    it('should respect range limiting option', () => {
      const ws = createMockSheet([
        ['A', 'B', 'C', 'D'],
        ['1', '2', '3', '4'],
        ['5', '6', '7', '8'],
        ['9', '10', '11', '12'],
      ])

      const result = parseSheet(ws, 'Sheet1', { range: 'A1:B2' })

      expect(result.cells.length).toBe(4) // Only A1, B1, A2, B2
      const cols = [...new Set(result.cells.map(c => c.col))]
      const rows = [...new Set(result.cells.map(c => c.row))]
      expect(cols).toEqual(['A', 'B'])
      expect(rows).toEqual([1, 2])
    })

    it('should handle sheet with merged cells', () => {
      const ws = createMockSheet([
        ['Merged', '', 'Normal'],
        ['Data', 'More', 'Data'],
      ])
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]

      const result = parseSheet(ws, 'Sheet1')

      expect(result.merges).toBeDefined()
      expect(result.merges).toHaveLength(1)
      expect(result.merges[0]).toEqual({
        start: { row: 1, col: 'A' },
        end: { row: 1, col: 'B' },
      })
    })

    it('should preserve column widths', () => {
      const ws = createMockSheet([['A', 'B', 'C']])
      ws['!cols'] = [
        { wch: 20 },
        { wch: 15 },
        { wch: 10 },
      ]

      const result = parseSheet(ws, 'Sheet1')

      expect(result.columnWidths).toBeDefined()
      expect(result.columnWidths?.A).toBe(20)
      expect(result.columnWidths?.B).toBe(15)
      expect(result.columnWidths?.C).toBe(10)
    })

    it('should preserve row heights', () => {
      const ws = createMockSheet([['Row1'], ['Row2'], ['Row3']])
      ws['!rows'] = [
        { hpt: 30 },
        { hpt: 20 },
        { hpt: 15 },
      ]

      const result = parseSheet(ws, 'Sheet1')

      expect(result.rowHeights).toBeDefined()
      expect(result.rowHeights?.[1]).toBe(30)
      expect(result.rowHeights?.[2]).toBe(20)
      expect(result.rowHeights?.[3]).toBe(15)
    })
  })

  describe('cellFromSheetJS', () => {
    it('should convert a string cell', () => {
      const sjsCell: XLSX.CellObject = { v: 'Hello', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'A', 1)

      expect(result._id).toBe('Sheet1!A1')
      expect(result.sheet).toBe('Sheet1')
      expect(result.col).toBe('A')
      expect(result.row).toBe(1)
      expect(result.colIndex).toBe(0)
      expect(result.value.v).toBe('Hello')
      expect(result.value.t).toBe('string')
    })

    it('should convert a number cell', () => {
      const sjsCell: XLSX.CellObject = { v: 42.5, t: 'n' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'B', 2)

      expect(result.value.v).toBe(42.5)
      expect(result.value.t).toBe('number')
    })

    it('should convert a boolean cell', () => {
      const sjsCell: XLSX.CellObject = { v: true, t: 'b' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'C', 3)

      expect(result.value.v).toBe(true)
      expect(result.value.t).toBe('boolean')
    })

    it('should convert a date cell', () => {
      const date = new Date('2024-06-15')
      const sjsCell: XLSX.CellObject = { v: date, t: 'd' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'D', 4)

      expect(result.value.v).toEqual(date)
      expect(result.value.t).toBe('date')
    })

    it('should handle date stored as serial number', () => {
      // Excel serial date for 2024-01-15 (days since 1900-01-01)
      const sjsCell: XLSX.CellObject = { v: 45306, t: 'n', w: '1/15/2024' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'E', 5, { inferDates: true })

      expect(result.value.t).toBe('date')
      expect(result.value.v).toBeInstanceOf(Date)
    })

    it('should convert an error cell', () => {
      const sjsCell: XLSX.CellObject = { v: '#DIV/0!', t: 'e' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'F', 6)

      expect(result.value.t).toBe('error')
      expect(result.value.e).toBe('#DIV/0!')
    })

    it('should preserve formula', () => {
      const sjsCell: XLSX.CellObject = { v: 100, t: 'n', f: 'A1+B1' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'G', 7)

      expect(result.value.v).toBe(100)
      expect(result.value.f).toBe('=A1+B1') // Normalized with = prefix
    })

    it('should preserve array formula', () => {
      const sjsCell: XLSX.CellObject = { v: 10, t: 'n', f: 'SUM(A1:A10)', F: 'A1:B10' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'H', 8)

      expect(result.value.f).toBe('=SUM(A1:A10)')
      expect(result.metadata?.arrayFormulaRange).toBe('A1:B10')
    })

    it('should preserve formatted text', () => {
      const sjsCell: XLSX.CellObject = { v: 0.75, t: 'n', w: '75%' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'I', 9)

      expect(result.value.w).toBe('75%')
    })

    it('should handle hyperlink', () => {
      const sjsCell: XLSX.CellObject = {
        v: 'Click here',
        t: 's',
        l: { Target: 'https://example.com', Tooltip: 'Visit site' },
      }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'J', 10)

      expect(result.metadata?.hyperlink?.target).toBe('https://example.com')
      expect(result.metadata?.hyperlink?.tooltip).toBe('Visit site')
    })

    it('should handle comment/note', () => {
      const sjsCell: XLSX.CellObject = {
        v: 'Cell with note',
        t: 's',
        c: [{ a: 'Author', t: 'This is a comment' }],
      }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'K', 11)

      expect(result.metadata?.comment?.text).toBe('This is a comment')
      expect(result.metadata?.comment?.author).toBe('Author')
    })

    it('should handle multi-letter columns', () => {
      const sjsCell: XLSX.CellObject = { v: 'AA column', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'AA', 1)

      expect(result.col).toBe('AA')
      expect(result.colIndex).toBe(26)
    })

    it('should handle sheet names with spaces', () => {
      const sjsCell: XLSX.CellObject = { v: 'Test', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'My Sheet', 'A', 1)

      expect(result._id).toBe("'My Sheet'!A1")
      expect(result.sheet).toBe('My Sheet')
    })
  })

  describe('rowsFromSheet', () => {
    it('should extract rows as documents using first row as headers', () => {
      const ws = createMockSheet([
        ['Name', 'Age', 'City'],
        ['Alice', 30, 'NYC'],
        ['Bob', 25, 'LA'],
      ])

      const result = rowsFromSheet(ws, { headers: true })

      expect(result).toHaveLength(2)
      expect(result[0]).toEqual({ Name: 'Alice', Age: 30, City: 'NYC' })
      expect(result[1]).toEqual({ Name: 'Bob', Age: 25, City: 'LA' })
    })

    it('should use column letters as keys when headers is false', () => {
      const ws = createMockSheet([
        ['Alice', 30, 'NYC'],
        ['Bob', 25, 'LA'],
      ])

      const result = rowsFromSheet(ws, { headers: false })

      expect(result).toHaveLength(2)
      expect(result[0]).toEqual({ A: 'Alice', B: 30, C: 'NYC' })
      expect(result[1]).toEqual({ A: 'Bob', B: 25, C: 'LA' })
    })

    it('should use custom headers array', () => {
      const ws = createMockSheet([
        ['Alice', 30, 'NYC'],
        ['Bob', 25, 'LA'],
      ])

      const result = rowsFromSheet(ws, { headers: ['name', 'age', 'city'] })

      expect(result[0]).toEqual({ name: 'Alice', age: 30, city: 'NYC' })
    })

    it('should respect range option for row extraction', () => {
      const ws = createMockSheet([
        ['Header1', 'Header2'],
        ['Skip1', 'Skip2'],
        ['Name', 'Value'],
        ['A', 1],
        ['B', 2],
      ])

      const result = rowsFromSheet(ws, { headers: true, range: 'A3:B5' })

      expect(result).toHaveLength(2)
      expect(result[0]).toEqual({ Name: 'A', Value: 1 })
    })

    it('should handle empty cells in rows', () => {
      const ws = createMockSheet([
        ['Name', 'Age', 'City'],
        ['Alice', null, 'NYC'],
        ['Bob', 25, null],
      ])

      const result = rowsFromSheet(ws, { headers: true })

      expect(result[0]).toEqual({ Name: 'Alice', Age: null, City: 'NYC' })
      expect(result[1]).toEqual({ Name: 'Bob', Age: 25, City: null })
    })

    it('should skip empty rows when option is set', () => {
      const ws = createMockSheet([
        ['Name', 'Age'],
        ['Alice', 30],
        [null, null],
        ['Bob', 25],
      ])

      const result = rowsFromSheet(ws, { headers: true, skipEmptyRows: true })

      expect(result).toHaveLength(2)
      expect(result[0].Name).toBe('Alice')
      expect(result[1].Name).toBe('Bob')
    })

    it('should include row metadata when requested', () => {
      const ws = createMockSheet([
        ['Name', 'Age'],
        ['Alice', 30],
      ])

      const result = rowsFromSheet(ws, { headers: true, includeRowMeta: true })

      expect(result[0]._rowNumber).toBe(2)
      expect(result[0]._range).toBe('A2:B2')
    })

    it('should handle duplicate header names', () => {
      const ws = createMockSheet([
        ['Name', 'Value', 'Value'],
        ['Alice', 1, 2],
      ])

      const result = rowsFromSheet(ws, { headers: true })

      // Should rename duplicates to Value_1, Value_2
      expect(result[0]).toHaveProperty('Name')
      expect(result[0]).toHaveProperty('Value')
      expect(result[0]).toHaveProperty('Value_1')
    })

    it('should handle blank header cells', () => {
      const ws = createMockSheet([
        ['Name', '', 'Age'],
        ['Alice', 'Extra', 30],
      ])

      const result = rowsFromSheet(ws, { headers: true })

      // Should use column letter for blank headers
      expect(result[0]).toHaveProperty('Name')
      expect(result[0]).toHaveProperty('B')
      expect(result[0]).toHaveProperty('Age')
    })
  })

  describe('Import Options', () => {
    describe('headers option', () => {
      it('should auto-detect headers when set to "auto"', () => {
        // First row looks like headers (all strings, unique)
        const ws = createMockSheet([
          ['ID', 'Name', 'Score'],
          [1, 'Alice', 95],
          [2, 'Bob', 87],
        ])

        const result = rowsFromSheet(ws, { headers: 'auto' })

        expect(result[0]).toHaveProperty('ID')
        expect(result[0]).toHaveProperty('Name')
      })

      it('should not use headers when first row contains numbers in auto mode', () => {
        const ws = createMockSheet([
          [1, 2, 3],
          [4, 5, 6],
        ])

        const result = rowsFromSheet(ws, { headers: 'auto' })

        // Should use column letters since first row has numbers
        expect(result[0]).toHaveProperty('A')
        expect(result).toHaveLength(2) // Both rows included
      })
    })

    describe('range option', () => {
      it('should limit import to specified range', () => {
        const ws = createMockSheet([
          ['A1', 'B1', 'C1', 'D1'],
          ['A2', 'B2', 'C2', 'D2'],
          ['A3', 'B3', 'C3', 'D3'],
          ['A4', 'B4', 'C4', 'D4'],
        ])

        const result = parseSheet(ws, 'Sheet1', { range: 'B2:C3' })

        const cellAddresses = result.cells.map(c => `${c.col}${c.row}`)
        expect(cellAddresses).toEqual(['B2', 'C2', 'B3', 'C3'])
      })

      it('should support row-only range (e.g., "2:5")', () => {
        const ws = createMockSheet([
          ['Header1', 'Header2'],
          ['R2C1', 'R2C2'],
          ['R3C1', 'R3C2'],
          ['R4C1', 'R4C2'],
          ['R5C1', 'R5C2'],
        ])

        const result = parseSheet(ws, 'Sheet1', { range: '2:4' })

        const rows = [...new Set(result.cells.map(c => c.row))]
        expect(rows).toEqual([2, 3, 4])
      })

      it('should support column-only range (e.g., "A:C")', () => {
        const ws = createMockSheet([
          ['A', 'B', 'C', 'D', 'E'],
        ])

        const result = parseSheet(ws, 'Sheet1', { range: 'A:C' })

        const cols = result.cells.map(c => c.col)
        expect(cols).toEqual(['A', 'B', 'C'])
      })
    })

    describe('type inference options', () => {
      it('should infer dates from formatted strings when enabled', () => {
        const ws: XLSX.WorkSheet = {
          A1: { v: '2024-01-15', t: 's' },
          A2: { v: '01/15/2024', t: 's' },
          A3: { v: 'January 15, 2024', t: 's' },
          '!ref': 'A1:A3',
        }

        const result = parseSheet(ws, 'Sheet1', { inferDates: true })

        expect(result.cells[0].value.t).toBe('date')
        expect(result.cells[1].value.t).toBe('date')
        expect(result.cells[2].value.t).toBe('date')
      })

      it('should keep dates as strings when inference is disabled', () => {
        const ws: XLSX.WorkSheet = {
          A1: { v: '2024-01-15', t: 's' },
          '!ref': 'A1:A1',
        }

        const result = parseSheet(ws, 'Sheet1', { inferDates: false })

        expect(result.cells[0].value.t).toBe('string')
      })

      it('should infer numbers from numeric strings when enabled', () => {
        const ws: XLSX.WorkSheet = {
          A1: { v: '42', t: 's' },
          A2: { v: '3.14', t: 's' },
          A3: { v: '-100', t: 's' },
          '!ref': 'A1:A3',
        }

        const result = parseSheet(ws, 'Sheet1', { inferNumbers: true })

        expect(result.cells[0].value.t).toBe('number')
        expect(result.cells[0].value.v).toBe(42)
        expect(result.cells[1].value.v).toBe(3.14)
        expect(result.cells[2].value.v).toBe(-100)
      })

      it('should handle currency strings with inference', () => {
        const ws: XLSX.WorkSheet = {
          A1: { v: '$1,234.56', t: 's' },
          '!ref': 'A1:A1',
        }

        const result = parseSheet(ws, 'Sheet1', { inferNumbers: true })

        expect(result.cells[0].value.t).toBe('number')
        expect(result.cells[0].value.v).toBe(1234.56)
      })

      it('should infer booleans from strings when enabled', () => {
        const ws: XLSX.WorkSheet = {
          A1: { v: 'true', t: 's' },
          A2: { v: 'FALSE', t: 's' },
          A3: { v: 'yes', t: 's' },
          A4: { v: 'no', t: 's' },
          '!ref': 'A1:A4',
        }

        const result = parseSheet(ws, 'Sheet1', { inferBooleans: true })

        expect(result.cells[0].value.t).toBe('boolean')
        expect(result.cells[0].value.v).toBe(true)
        expect(result.cells[1].value.v).toBe(false)
        expect(result.cells[2].value.v).toBe(true)
        expect(result.cells[3].value.v).toBe(false)
      })
    })
  })

  describe('Formula Preservation', () => {
    it('should preserve simple formula', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 10, t: 'n' },
        B1: { v: 20, t: 'n' },
        C1: { v: 30, t: 'n', f: 'A1+B1' },
        '!ref': 'A1:C1',
      }

      const result = parseSheet(ws, 'Sheet1')

      const c1 = result.cells.find(c => c.col === 'C' && c.row === 1)
      expect(c1?.value.f).toBe('=A1+B1')
      expect(c1?.value.v).toBe(30)
    })

    it('should preserve formula with functions', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 100, t: 'n', f: 'SUM(B1:B10)' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1')

      expect(result.cells[0].value.f).toBe('=SUM(B1:B10)')
    })

    it('should extract formula dependencies', () => {
      const ws: XLSX.WorkSheet = {
        C1: { v: 30, t: 'n', f: 'A1+B1' },
        '!ref': 'C1:C1',
      }

      const result = parseSheet(ws, 'Sheet1', { extractDependencies: true })

      const c1 = result.cells.find(c => c.col === 'C')
      expect(c1?.dependencies).toContain('Sheet1!A1')
      expect(c1?.dependencies).toContain('Sheet1!B1')
    })

    it('should handle cross-sheet references', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 100, t: 'n', f: 'Sheet2!A1+Sheet3!B2' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { extractDependencies: true })

      expect(result.cells[0].dependencies).toContain('Sheet2!A1')
      expect(result.cells[0].dependencies).toContain('Sheet3!B2')
    })

    it('should handle range references in formulas', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 100, t: 'n', f: 'SUM(B1:B10)' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { extractDependencies: true })

      // Should expand range to individual cells or store range reference
      expect(result.cells[0].dependencies).toBeDefined()
      expect(result.cells[0].dependencies?.length).toBeGreaterThan(0)
    })

    it('should preserve named range references in formulas', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 500, t: 'n', f: 'SUM(SalesData)' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1')

      expect(result.cells[0].value.f).toBe('=SUM(SalesData)')
    })

    it('should handle array formulas (CSE formulas)', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 1, t: 'n', f: 'TRANSPOSE(B1:D1)', F: 'A1:A3' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1')

      const a1 = result.cells[0]
      expect(a1.value.f).toBe('=TRANSPOSE(B1:D1)')
      expect(a1.metadata?.arrayFormulaRange).toBe('A1:A3')
    })
  })

  describe('Style/Format Preservation', () => {
    it('should preserve number format', () => {
      const ws: XLSX.WorkSheet = {
        A1: { v: 0.75, t: 'n', z: '0.00%' },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      expect(result.cells[0].format?.numberFormat).toBe('0.00%')
    })

    it('should preserve font styling', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Bold Text',
          t: 's',
          s: {
            font: {
              bold: true,
              sz: 14,
              name: 'Arial',
              color: { rgb: 'FF0000' },
            },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      const format = result.cells[0].format
      expect(format?.font?.bold).toBe(true)
      expect(format?.font?.size).toBe(14)
      expect(format?.font?.name).toBe('Arial')
      expect(format?.font?.color).toBe('#FF0000')
    })

    it('should preserve fill/background color', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Colored',
          t: 's',
          s: {
            fill: {
              patternType: 'solid',
              fgColor: { rgb: '00FF00' },
            },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      const format = result.cells[0].format
      expect(format?.fill?.type).toBe('solid')
      expect(format?.fill?.color).toBe('#00FF00')
    })

    it('should preserve border styles', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Bordered',
          t: 's',
          s: {
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thick', color: { rgb: '000000' } },
              left: { style: 'dashed', color: { rgb: 'FF0000' } },
              right: { style: 'dotted', color: { rgb: '0000FF' } },
            },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      const border = result.cells[0].format?.border
      expect(border?.top?.style).toBe('thin')
      expect(border?.bottom?.style).toBe('thick')
      expect(border?.left?.style).toBe('dashed')
      expect(border?.right?.style).toBe('dotted')
    })

    it('should preserve alignment settings', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Aligned',
          t: 's',
          s: {
            alignment: {
              horizontal: 'center',
              vertical: 'top',
              wrapText: true,
              textRotation: 45,
            },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      const alignment = result.cells[0].format?.alignment
      expect(alignment?.horizontal).toBe('center')
      expect(alignment?.vertical).toBe('top')
      expect(alignment?.wrapText).toBe(true)
      expect(alignment?.textRotation).toBe(45)
    })

    it('should ignore styles when preserveStyles is false', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Styled',
          t: 's',
          s: {
            font: { bold: true },
            fill: { patternType: 'solid', fgColor: { rgb: 'FF0000' } },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: false })

      expect(result.cells[0].format).toBeUndefined()
    })

    it('should preserve text underline and strikethrough', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Underline',
          t: 's',
          s: { font: { underline: true } },
        },
        A2: {
          v: 'Strike',
          t: 's',
          s: { font: { strike: true } },
        },
        '!ref': 'A1:A2',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      expect(result.cells[0].format?.font?.underline).toBe(true)
      expect(result.cells[1].format?.font?.strikethrough).toBe(true)
    })

    it('should preserve cell protection settings', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Protected',
          t: 's',
          s: {
            protection: { locked: true, hidden: true },
          },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      expect(result.cells[0].format?.protection?.locked).toBe(true)
      expect(result.cells[0].format?.protection?.hidden).toBe(true)
    })

    it('should handle italic font', () => {
      const ws: XLSX.WorkSheet = {
        A1: {
          v: 'Italic',
          t: 's',
          s: { font: { italic: true } },
        },
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      expect(result.cells[0].format?.font?.italic).toBe(true)
    })
  })

  describe('Edge Cases', () => {
    it('should handle very large column indices', () => {
      const sjsCell: XLSX.CellObject = { v: 'XFD', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'XFD', 1)

      expect(result.col).toBe('XFD')
      expect(result.colIndex).toBe(16383) // Last Excel column
    })

    it('should handle maximum row number', () => {
      const sjsCell: XLSX.CellObject = { v: 'Max row', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'A', 1048576)

      expect(result.row).toBe(1048576)
    })

    it('should handle unicode content', () => {
      const sjsCell: XLSX.CellObject = { v: 'Hello', t: 's' }

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'A', 1)

      expect(result.value.v).toBe('Hello')
    })

    it('should handle cells with rich text', () => {
      const sjsCell: XLSX.CellObject = {
        v: 'Rich Text',
        t: 's',
        r: '<r><t>Rich </t></r><r><rPr><b/></rPr><t>Text</t></r>',
      } as XLSX.CellObject

      const result = cellFromSheetJS(sjsCell, 'Sheet1', 'A', 1)

      expect(result.value.v).toBe('Rich Text')
      // Rich text info may be preserved in metadata
    })

    it('should handle empty worksheet gracefully', () => {
      const ws: XLSX.WorkSheet = {}

      const result = parseSheet(ws, 'Empty')

      expect(result.cells).toHaveLength(0)
      expect(result.dimensions).toBeUndefined()
    })

    it('should handle worksheet with only formatting (no values)', () => {
      const ws: XLSX.WorkSheet = {
        A1: { s: { font: { bold: true } } } as XLSX.CellObject,
        '!ref': 'A1:A1',
      }

      const result = parseSheet(ws, 'Sheet1', { preserveStyles: true })

      // Cell exists but has empty value
      expect(result.cells).toHaveLength(1)
      expect(result.cells[0].value.t).toBe('empty')
    })
  })
})
