/**
 * Cell Address Utilities Tests (RED)
 *
 * TDD: These tests define expected behavior for A1/R1C1 notation parsing.
 * All tests should FAIL initially.
 */

import { describe, it, expect } from 'vitest'
import {
  parseA1,
  parseR1C1,
  parseRange,
  parseReference,
  colToIndex,
  indexToCol,
  isValidA1,
  isValidR1C1,
  toA1,
  toR1C1,
  offsetReference,
  expandRange,
  rangeContains,
  rangeIntersects,
  mergeRanges,
} from './address'
import type { CellReference, RangeReference, R1C1Reference } from '../types'

describe('Cell Address Utilities', () => {
  describe('colToIndex', () => {
    it('should convert single letter columns', () => {
      expect(colToIndex('A')).toBe(0)
      expect(colToIndex('B')).toBe(1)
      expect(colToIndex('Z')).toBe(25)
    })

    it('should convert double letter columns', () => {
      expect(colToIndex('AA')).toBe(26)
      expect(colToIndex('AB')).toBe(27)
      expect(colToIndex('AZ')).toBe(51)
      expect(colToIndex('BA')).toBe(52)
      expect(colToIndex('ZZ')).toBe(701)
    })

    it('should convert triple letter columns', () => {
      expect(colToIndex('AAA')).toBe(702)
      expect(colToIndex('XFD')).toBe(16383) // Excel max column
    })

    it('should be case-insensitive', () => {
      expect(colToIndex('a')).toBe(0)
      expect(colToIndex('aa')).toBe(26)
      expect(colToIndex('Aa')).toBe(26)
    })

    it('should throw for invalid column names', () => {
      expect(() => colToIndex('')).toThrow()
      expect(() => colToIndex('123')).toThrow()
      expect(() => colToIndex('A1')).toThrow()
    })
  })

  describe('indexToCol', () => {
    it('should convert single digit indices', () => {
      expect(indexToCol(0)).toBe('A')
      expect(indexToCol(1)).toBe('B')
      expect(indexToCol(25)).toBe('Z')
    })

    it('should convert to double letter columns', () => {
      expect(indexToCol(26)).toBe('AA')
      expect(indexToCol(27)).toBe('AB')
      expect(indexToCol(51)).toBe('AZ')
      expect(indexToCol(52)).toBe('BA')
      expect(indexToCol(701)).toBe('ZZ')
    })

    it('should convert to triple letter columns', () => {
      expect(indexToCol(702)).toBe('AAA')
      expect(indexToCol(16383)).toBe('XFD')
    })

    it('should throw for negative indices', () => {
      expect(() => indexToCol(-1)).toThrow()
    })

    it('should be inverse of colToIndex', () => {
      for (let i = 0; i < 100; i++) {
        expect(colToIndex(indexToCol(i))).toBe(i)
      }
    })
  })

  describe('parseA1', () => {
    it('should parse simple cell reference', () => {
      const ref = parseA1('A1')
      expect(ref.col).toBe('A')
      expect(ref.colIndex).toBe(0)
      expect(ref.row).toBe(1)
      expect(ref.absolute.row).toBe(false)
      expect(ref.absolute.col).toBe(false)
    })

    it('should parse with absolute column', () => {
      const ref = parseA1('$A1')
      expect(ref.col).toBe('A')
      expect(ref.absolute.col).toBe(true)
      expect(ref.absolute.row).toBe(false)
    })

    it('should parse with absolute row', () => {
      const ref = parseA1('A$1')
      expect(ref.absolute.col).toBe(false)
      expect(ref.absolute.row).toBe(true)
    })

    it('should parse fully absolute reference', () => {
      const ref = parseA1('$A$1')
      expect(ref.absolute.col).toBe(true)
      expect(ref.absolute.row).toBe(true)
    })

    it('should parse multi-letter columns', () => {
      const ref = parseA1('AA100')
      expect(ref.col).toBe('AA')
      expect(ref.colIndex).toBe(26)
      expect(ref.row).toBe(100)
    })

    it('should parse with sheet reference', () => {
      const ref = parseA1('Sheet1!A1')
      expect(ref.sheet).toBe('Sheet1')
      expect(ref.col).toBe('A')
      expect(ref.row).toBe(1)
    })

    it('should parse quoted sheet names', () => {
      const ref = parseA1("'My Sheet'!A1")
      expect(ref.sheet).toBe('My Sheet')
    })

    it('should parse sheet names with escaped quotes', () => {
      const ref = parseA1("'Sheet''s Data'!A1")
      expect(ref.sheet).toBe("Sheet's Data")
    })

    it('should preserve original string', () => {
      const ref = parseA1('$B$2')
      expect(ref.original).toBe('$B$2')
    })

    it('should throw for invalid reference', () => {
      expect(() => parseA1('')).toThrow()
      expect(() => parseA1('1A')).toThrow()
      expect(() => parseA1('A')).toThrow()
      expect(() => parseA1('1')).toThrow()
    })
  })

  describe('parseR1C1', () => {
    it('should parse absolute R1C1 reference', () => {
      const ref = parseR1C1('R1C1')
      expect(ref.row.value).toBe(1)
      expect(ref.row.relative).toBe(false)
      expect(ref.col.value).toBe(1)
      expect(ref.col.relative).toBe(false)
    })

    it('should parse relative R1C1 reference', () => {
      const ref = parseR1C1('R[1]C[2]')
      expect(ref.row.value).toBe(1)
      expect(ref.row.relative).toBe(true)
      expect(ref.col.value).toBe(2)
      expect(ref.col.relative).toBe(true)
    })

    it('should parse negative relative offsets', () => {
      const ref = parseR1C1('R[-1]C[-2]')
      expect(ref.row.value).toBe(-1)
      expect(ref.col.value).toBe(-2)
    })

    it('should parse mixed absolute/relative', () => {
      const ref = parseR1C1('R1C[2]')
      expect(ref.row.relative).toBe(false)
      expect(ref.col.relative).toBe(true)
    })

    it('should parse with sheet reference', () => {
      const ref = parseR1C1('Sheet1!R1C1')
      expect(ref.sheet).toBe('Sheet1')
    })

    it('should handle RC (same row/col) notation', () => {
      const ref = parseR1C1('RC')
      expect(ref.row.value).toBe(0)
      expect(ref.row.relative).toBe(true)
      expect(ref.col.value).toBe(0)
      expect(ref.col.relative).toBe(true)
    })
  })

  describe('parseRange', () => {
    it('should parse simple range', () => {
      const range = parseRange('A1:B2')
      expect(range.start.col).toBe('A')
      expect(range.start.row).toBe(1)
      expect(range.end.col).toBe('B')
      expect(range.end.row).toBe(2)
    })

    it('should parse range with absolute references', () => {
      const range = parseRange('$A$1:$B$2')
      expect(range.start.absolute.col).toBe(true)
      expect(range.start.absolute.row).toBe(true)
      expect(range.end.absolute.col).toBe(true)
      expect(range.end.absolute.row).toBe(true)
    })

    it('should parse range with sheet reference', () => {
      const range = parseRange('Sheet1!A1:B2')
      expect(range.sheet).toBe('Sheet1')
    })

    it('should normalize range (start <= end)', () => {
      const range = parseRange('B2:A1')
      expect(range.start.col).toBe('A')
      expect(range.start.row).toBe(1)
      expect(range.end.col).toBe('B')
      expect(range.end.row).toBe(2)
    })

    it('should parse full column reference', () => {
      const range = parseRange('A:A')
      expect(range.start.col).toBe('A')
      expect(range.start.row).toBe(1)
      expect(range.end.col).toBe('A')
      expect(range.end.row).toBe(1048576) // Excel max row
    })

    it('should parse full row reference', () => {
      const range = parseRange('1:1')
      expect(range.start.col).toBe('A')
      expect(range.start.row).toBe(1)
      expect(range.end.col).toBe('XFD')
      expect(range.end.row).toBe(1)
    })
  })

  describe('isValidA1', () => {
    it('should return true for valid references', () => {
      expect(isValidA1('A1')).toBe(true)
      expect(isValidA1('$A$1')).toBe(true)
      expect(isValidA1('AA100')).toBe(true)
      expect(isValidA1('XFD1048576')).toBe(true)
    })

    it('should return false for invalid references', () => {
      expect(isValidA1('')).toBe(false)
      expect(isValidA1('1A')).toBe(false)
      expect(isValidA1('A0')).toBe(false)
      expect(isValidA1('XFE1')).toBe(false) // Beyond XFD
      expect(isValidA1('A1048577')).toBe(false) // Beyond max row
    })
  })

  describe('toA1', () => {
    it('should convert R1C1 to A1', () => {
      expect(toA1({ row: 1, col: 1 })).toBe('A1')
      expect(toA1({ row: 2, col: 2 })).toBe('B2')
      expect(toA1({ row: 100, col: 27 })).toBe('AA100')
    })

    it('should convert with absolute markers', () => {
      expect(toA1({ row: 1, col: 1 }, { absolute: { row: true, col: true } })).toBe('$A$1')
      expect(toA1({ row: 1, col: 1 }, { absolute: { row: false, col: true } })).toBe('$A1')
    })

    it('should include sheet name', () => {
      expect(toA1({ row: 1, col: 1 }, { sheet: 'Sheet1' })).toBe('Sheet1!A1')
      expect(toA1({ row: 1, col: 1 }, { sheet: 'My Sheet' })).toBe("'My Sheet'!A1")
    })
  })

  describe('toR1C1', () => {
    it('should convert A1 to R1C1', () => {
      expect(toR1C1('A1')).toBe('R1C1')
      expect(toR1C1('B2')).toBe('R2C2')
    })

    it('should convert with relative offset', () => {
      expect(toR1C1('A1', { baseCell: { row: 2, col: 2 } })).toBe('R[-1]C[-1]')
      expect(toR1C1('C3', { baseCell: { row: 1, col: 1 } })).toBe('R[2]C[2]')
    })

    it('should handle same row/col as RC', () => {
      expect(toR1C1('B2', { baseCell: { row: 2, col: 2 } })).toBe('RC')
    })
  })

  describe('offsetReference', () => {
    it('should offset row', () => {
      const ref = parseA1('A1')
      const offset = offsetReference(ref, { rowOffset: 1 })
      expect(offset.row).toBe(2)
      expect(offset.col).toBe('A')
    })

    it('should offset column', () => {
      const ref = parseA1('A1')
      const offset = offsetReference(ref, { colOffset: 1 })
      expect(offset.row).toBe(1)
      expect(offset.col).toBe('B')
    })

    it('should offset both', () => {
      const ref = parseA1('A1')
      const offset = offsetReference(ref, { rowOffset: 5, colOffset: 2 })
      expect(offset.row).toBe(6)
      expect(offset.col).toBe('C')
    })

    it('should not offset absolute components', () => {
      const ref = parseA1('$A$1')
      const offset = offsetReference(ref, { rowOffset: 1, colOffset: 1 })
      expect(offset.row).toBe(1) // Absolute, no change
      expect(offset.col).toBe('A')
    })

    it('should throw if offset goes out of bounds', () => {
      const ref = parseA1('A1')
      expect(() => offsetReference(ref, { rowOffset: -1 })).toThrow()
      expect(() => offsetReference(ref, { colOffset: -1 })).toThrow()
    })
  })

  describe('expandRange', () => {
    it('should expand simple range to cell array', () => {
      const cells = expandRange('A1:B2')
      expect(cells).toHaveLength(4)
      expect(cells).toContain('A1')
      expect(cells).toContain('A2')
      expect(cells).toContain('B1')
      expect(cells).toContain('B2')
    })

    it('should expand single cell range', () => {
      const cells = expandRange('A1:A1')
      expect(cells).toHaveLength(1)
      expect(cells).toContain('A1')
    })

    it('should expand column range', () => {
      const cells = expandRange('A1:A3')
      expect(cells).toEqual(['A1', 'A2', 'A3'])
    })

    it('should expand row range', () => {
      const cells = expandRange('A1:C1')
      expect(cells).toEqual(['A1', 'B1', 'C1'])
    })

    it('should include sheet name if present', () => {
      const cells = expandRange('Sheet1!A1:A2')
      expect(cells).toEqual(['Sheet1!A1', 'Sheet1!A2'])
    })
  })

  describe('rangeContains', () => {
    it('should return true if cell is in range', () => {
      expect(rangeContains('A1:C3', 'B2')).toBe(true)
      expect(rangeContains('A1:C3', 'A1')).toBe(true)
      expect(rangeContains('A1:C3', 'C3')).toBe(true)
    })

    it('should return false if cell is outside range', () => {
      expect(rangeContains('A1:C3', 'D1')).toBe(false)
      expect(rangeContains('A1:C3', 'A4')).toBe(false)
    })

    it('should handle sheet references', () => {
      expect(rangeContains('Sheet1!A1:C3', 'Sheet1!B2')).toBe(true)
      expect(rangeContains('Sheet1!A1:C3', 'Sheet2!B2')).toBe(false)
    })
  })

  describe('rangeIntersects', () => {
    it('should return true for overlapping ranges', () => {
      expect(rangeIntersects('A1:C3', 'B2:D4')).toBe(true)
      expect(rangeIntersects('A1:C3', 'C3:E5')).toBe(true)
    })

    it('should return false for non-overlapping ranges', () => {
      expect(rangeIntersects('A1:B2', 'C3:D4')).toBe(false)
      expect(rangeIntersects('A1:B2', 'A3:B4')).toBe(false)
    })

    it('should handle adjacent ranges (no intersection)', () => {
      expect(rangeIntersects('A1:A2', 'B1:B2')).toBe(false)
    })
  })

  describe('mergeRanges', () => {
    it('should merge adjacent ranges', () => {
      const merged = mergeRanges(['A1:A2', 'A3:A4'])
      expect(merged).toEqual(['A1:A4'])
    })

    it('should merge overlapping ranges', () => {
      const merged = mergeRanges(['A1:B2', 'B2:C3'])
      expect(merged).toHaveLength(1)
    })

    it('should keep non-overlapping ranges separate', () => {
      const merged = mergeRanges(['A1:A2', 'C1:C2'])
      expect(merged).toHaveLength(2)
    })
  })
})
