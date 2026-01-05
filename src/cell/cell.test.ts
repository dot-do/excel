/**
 * Cell Data Model Tests (RED)
 *
 * TDD: These tests define expected behavior before implementation.
 * All tests should FAIL initially.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import {
  createCell,
  createCellValue,
  getCellId,
  inferValueType,
  formatCellValue,
  validateCell,
  isEmptyCell,
  cloneCell,
  mergeCellFormat,
} from './cell'
import type { Cell, CellValue, CellFormat, CreateCellOptions } from '../types'

describe('Cell Data Model', () => {
  describe('createCellValue', () => {
    it('should create a string cell value', () => {
      const value = createCellValue('hello')
      expect(value.v).toBe('hello')
      expect(value.t).toBe('string')
    })

    it('should create a number cell value', () => {
      const value = createCellValue(42)
      expect(value.v).toBe(42)
      expect(value.t).toBe('number')
    })

    it('should create a boolean cell value', () => {
      const value = createCellValue(true)
      expect(value.v).toBe(true)
      expect(value.t).toBe('boolean')
    })

    it('should create a date cell value', () => {
      const date = new Date('2024-01-15')
      const value = createCellValue(date)
      expect(value.v).toEqual(date)
      expect(value.t).toBe('date')
    })

    it('should create an empty cell value for null', () => {
      const value = createCellValue(null)
      expect(value.v).toBe(null)
      expect(value.t).toBe('empty')
    })

    it('should create an empty cell value for undefined', () => {
      const value = createCellValue(undefined)
      expect(value.v).toBe(null)
      expect(value.t).toBe('empty')
    })

    it('should preserve formula in cell value', () => {
      const value = createCellValue(100, '=A1+B1')
      expect(value.v).toBe(100)
      expect(value.f).toBe('=A1+B1')
    })

    it('should handle error values', () => {
      const value = createCellValue('#DIV/0!')
      expect(value.t).toBe('error')
      expect(value.e).toBe('#DIV/0!')
    })
  })

  describe('inferValueType', () => {
    it('should infer string type', () => {
      expect(inferValueType('hello')).toBe('string')
      expect(inferValueType('')).toBe('string')
    })

    it('should infer number type', () => {
      expect(inferValueType(42)).toBe('number')
      expect(inferValueType(3.14)).toBe('number')
      expect(inferValueType(-100)).toBe('number')
      expect(inferValueType(0)).toBe('number')
    })

    it('should infer boolean type', () => {
      expect(inferValueType(true)).toBe('boolean')
      expect(inferValueType(false)).toBe('boolean')
    })

    it('should infer date type', () => {
      expect(inferValueType(new Date())).toBe('date')
    })

    it('should infer empty type for null/undefined', () => {
      expect(inferValueType(null)).toBe('empty')
      expect(inferValueType(undefined)).toBe('empty')
    })

    it('should infer error type for error strings', () => {
      expect(inferValueType('#REF!')).toBe('error')
      expect(inferValueType('#N/A')).toBe('error')
      expect(inferValueType('#VALUE!')).toBe('error')
    })
  })

  describe('getCellId', () => {
    it('should generate cell ID from sheet, col, and row', () => {
      expect(getCellId('Sheet1', 'A', 1)).toBe('Sheet1!A1')
      expect(getCellId('Sheet1', 'B', 2)).toBe('Sheet1!B2')
      expect(getCellId('Data', 'AA', 100)).toBe('Data!AA100')
    })

    it('should handle sheet names with spaces', () => {
      expect(getCellId('My Sheet', 'A', 1)).toBe("'My Sheet'!A1")
    })

    it('should handle sheet names with special characters', () => {
      expect(getCellId("Sheet's Data", 'A', 1)).toBe("'Sheet''s Data'!A1")
    })
  })

  describe('createCell', () => {
    it('should create a cell with string value', () => {
      const options: CreateCellOptions = {
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Hello',
      }
      const cell = createCell(options)

      expect(cell._id).toBe('Sheet1!A1')
      expect(cell.sheet).toBe('Sheet1')
      expect(cell.row).toBe(1)
      expect(cell.col).toBe('A')
      expect(cell.colIndex).toBe(0)
      expect(cell.value.v).toBe('Hello')
      expect(cell.value.t).toBe('string')
      expect(cell.createdAt).toBeInstanceOf(Date)
      expect(cell.updatedAt).toBeInstanceOf(Date)
    })

    it('should create a cell with number value', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'B',
        value: 42,
      })

      expect(cell.value.v).toBe(42)
      expect(cell.value.t).toBe('number')
      expect(cell.colIndex).toBe(1)
    })

    it('should create a cell with formula', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 3,
        col: 'C',
        formula: '=A1+B1',
      })

      expect(cell.value.f).toBe('=A1+B1')
      expect(cell.value.t).toBe('empty') // No computed value yet
    })

    it('should create a cell with formatting', () => {
      const format: CellFormat = {
        font: { bold: true, size: 14 },
        fill: { type: 'solid', color: '#FF0000' },
      }
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Bold Red',
        format,
      })

      expect(cell.format).toEqual(format)
    })

    it('should create a cell with metadata', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'With comment',
        metadata: {
          comment: { text: 'This is a note', author: 'User' },
        },
      })

      expect(cell.metadata?.comment?.text).toBe('This is a note')
    })

    it('should handle multi-letter columns', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'AA',
        value: 'Column AA',
      })

      expect(cell.col).toBe('AA')
      expect(cell.colIndex).toBe(26) // AA is the 27th column (0-indexed = 26)
    })

    it('should handle column AZ', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'AZ',
        value: 'Column AZ',
      })

      expect(cell.colIndex).toBe(51) // AZ = 51 (0-indexed)
    })

    it('should handle column XFD (Excel max)', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'XFD',
        value: 'Max column',
      })

      expect(cell.colIndex).toBe(16383) // XFD = 16383 (0-indexed)
    })
  })

  describe('validateCell', () => {
    it('should return valid for a properly formed cell', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Test',
      })
      const result = validateCell(cell)
      expect(result.valid).toBe(true)
      expect(result.errors).toHaveLength(0)
    })

    it('should reject invalid row numbers', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 0,
        col: 'A',
        value: 'Test',
      })
      cell.row = 0 // Force invalid row
      const result = validateCell(cell)
      expect(result.valid).toBe(false)
      expect(result.errors).toContain('Row must be >= 1')
    })

    it('should reject row numbers exceeding Excel limit', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1048577, // Excel max is 1048576
        col: 'A',
        value: 'Test',
      })
      const result = validateCell(cell)
      expect(result.valid).toBe(false)
      expect(result.errors).toContain('Row exceeds maximum (1048576)')
    })

    it('should reject invalid column names', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Test',
      })
      cell.col = '123' // Force invalid column
      const result = validateCell(cell)
      expect(result.valid).toBe(false)
      expect(result.errors).toContain('Invalid column name')
    })

    it('should reject empty sheet names', () => {
      const cell = createCell({
        sheet: '',
        row: 1,
        col: 'A',
        value: 'Test',
      })
      const result = validateCell(cell)
      expect(result.valid).toBe(false)
      expect(result.errors).toContain('Sheet name cannot be empty')
    })
  })

  describe('isEmptyCell', () => {
    it('should return true for null value', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: null,
      })
      expect(isEmptyCell(cell)).toBe(true)
    })

    it('should return true for empty string', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: '',
      })
      expect(isEmptyCell(cell)).toBe(true)
    })

    it('should return false for non-empty value', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Hello',
      })
      expect(isEmptyCell(cell)).toBe(false)
    })

    it('should return false for zero', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 0,
      })
      expect(isEmptyCell(cell)).toBe(false)
    })

    it('should return false for false boolean', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: false,
      })
      expect(isEmptyCell(cell)).toBe(false)
    })

    it('should return false if cell has formula even with empty value', () => {
      const cell = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        formula: '=A2',
      })
      expect(isEmptyCell(cell)).toBe(false)
    })
  })

  describe('cloneCell', () => {
    it('should create a deep copy of a cell', () => {
      const original = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Test',
        format: { font: { bold: true } },
      })
      const cloned = cloneCell(original)

      expect(cloned).toEqual(original)
      expect(cloned).not.toBe(original)
      expect(cloned.format).not.toBe(original.format)
    })

    it('should allow overriding properties during clone', () => {
      const original = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Original',
      })
      const cloned = cloneCell(original, { value: 'Cloned' })

      expect(cloned.value.v).toBe('Cloned')
      expect(original.value.v).toBe('Original')
    })

    it('should update cell ID when row/col changes', () => {
      const original = createCell({
        sheet: 'Sheet1',
        row: 1,
        col: 'A',
        value: 'Test',
      })
      const cloned = cloneCell(original, { row: 2, col: 'B' })

      expect(cloned._id).toBe('Sheet1!B2')
      expect(cloned.row).toBe(2)
      expect(cloned.col).toBe('B')
    })
  })

  describe('mergeCellFormat', () => {
    it('should merge two formats', () => {
      const base: CellFormat = {
        font: { bold: true, size: 12 },
      }
      const override: CellFormat = {
        font: { italic: true },
        fill: { type: 'solid', color: '#FF0000' },
      }
      const merged = mergeCellFormat(base, override)

      expect(merged.font?.bold).toBe(true)
      expect(merged.font?.italic).toBe(true)
      expect(merged.font?.size).toBe(12)
      expect(merged.fill?.color).toBe('#FF0000')
    })

    it('should override nested properties', () => {
      const base: CellFormat = {
        font: { size: 12, color: '#000000' },
      }
      const override: CellFormat = {
        font: { size: 14 },
      }
      const merged = mergeCellFormat(base, override)

      expect(merged.font?.size).toBe(14)
      expect(merged.font?.color).toBe('#000000')
    })

    it('should handle undefined inputs', () => {
      const format: CellFormat = { font: { bold: true } }

      expect(mergeCellFormat(undefined, format)).toEqual(format)
      expect(mergeCellFormat(format, undefined)).toEqual(format)
      expect(mergeCellFormat(undefined, undefined)).toEqual({})
    })
  })

  describe('formatCellValue', () => {
    it('should format number with default format', () => {
      const value: CellValue = { v: 1234.567, t: 'number' }
      expect(formatCellValue(value)).toBe('1234.567')
    })

    it('should format number with custom format', () => {
      const value: CellValue = { v: 0.75, t: 'number' }
      expect(formatCellValue(value, '0%')).toBe('75%')
    })

    it('should format currency', () => {
      const value: CellValue = { v: 1234.5, t: 'number' }
      expect(formatCellValue(value, '$#,##0.00')).toBe('$1,234.50')
    })

    it('should format date with default format', () => {
      const date = new Date('2024-01-15')
      const value: CellValue = { v: date, t: 'date' }
      expect(formatCellValue(value)).toMatch(/2024/)
    })

    it('should format date with custom format', () => {
      const date = new Date('2024-01-15')
      const value: CellValue = { v: date, t: 'date' }
      expect(formatCellValue(value, 'yyyy-mm-dd')).toBe('2024-01-15')
    })

    it('should return boolean as TRUE/FALSE', () => {
      expect(formatCellValue({ v: true, t: 'boolean' })).toBe('TRUE')
      expect(formatCellValue({ v: false, t: 'boolean' })).toBe('FALSE')
    })

    it('should return error value as-is', () => {
      const value: CellValue = { v: '#DIV/0!', t: 'error', e: '#DIV/0!' }
      expect(formatCellValue(value)).toBe('#DIV/0!')
    })

    it('should return empty string for empty cells', () => {
      const value: CellValue = { v: null, t: 'empty' }
      expect(formatCellValue(value)).toBe('')
    })
  })
})
