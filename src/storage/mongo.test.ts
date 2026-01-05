/**
 * MongoDB Storage Layer Tests for excel.do
 *
 * TDD RED: Tests for mongo.do storage layer.
 * These tests should FAIL until implementation is complete.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import {
  CellStore,
  SheetStore,
  WorkbookStore,
  StorageManager,
  NotImplementedError,
  createIndex,
  createCellIndexes,
  parseCellRef,
  parseRange,
  generateCellId,
  generateId,
} from './mongo'
import type {
  MongoClient,
  MongoDatabase,
  MongoCollection,
  MongoSession,
  MongoFilter,
  MongoUpdate,
  FindOptions,
  InsertResult,
  InsertManyResult,
  UpdateResult,
  DeleteResult,
  IndexSpec,
  IndexOptions,
  IndexInfo,
  Sheet,
  Workbook,
  CellData,
  CellQuery,
  Range,
} from './types'
import type { Cell, CellPrimitive } from '../types'

// ============================================================================
// Mock mongo.do Client
// ============================================================================

function createMockCollection<T>(): MongoCollection<T> {
  const data = new Map<string, T>()

  return {
    findOne: vi.fn(async (filter: MongoFilter): Promise<T | null> => {
      const id = (filter as { _id?: string })._id
      if (id) {
        return data.get(id) || null
      }
      // Simple filter matching for tests
      for (const doc of data.values()) {
        let matches = true
        for (const [key, value] of Object.entries(filter)) {
          if ((doc as Record<string, unknown>)[key] !== value) {
            matches = false
            break
          }
        }
        if (matches) return doc
      }
      return null
    }),

    find: vi.fn(async (filter: MongoFilter, _options?: FindOptions): Promise<T[]> => {
      const results: T[] = []
      for (const doc of data.values()) {
        let matches = true
        for (const [key, value] of Object.entries(filter)) {
          const docValue = (doc as Record<string, unknown>)[key]
          if (docValue !== value) {
            // Handle special operators like $gte, $lte
            if (typeof value === 'object' && value !== null) {
              const ops = value as Record<string, unknown>
              if ('$gte' in ops && typeof docValue === 'number' && docValue < (ops.$gte as number)) {
                matches = false
              }
              if ('$lte' in ops && typeof docValue === 'number' && docValue > (ops.$lte as number)) {
                matches = false
              }
            } else {
              matches = false
            }
          }
        }
        if (matches) results.push(doc)
      }
      return results
    }),

    insertOne: vi.fn(async (doc: T): Promise<InsertResult> => {
      const id = (doc as { _id?: string })._id || `mock-${Date.now()}-${Math.random()}`
      ;(doc as { _id: string })._id = id
      data.set(id, doc)
      return { insertedId: id, acknowledged: true }
    }),

    insertMany: vi.fn(async (docs: T[]): Promise<InsertManyResult> => {
      const ids: string[] = []
      for (const doc of docs) {
        const id = (doc as { _id?: string })._id || `mock-${Date.now()}-${Math.random()}`
        ;(doc as { _id: string })._id = id
        data.set(id, doc)
        ids.push(id)
      }
      return { insertedIds: ids, insertedCount: docs.length, acknowledged: true }
    }),

    updateOne: vi.fn(async (filter: MongoFilter, update: MongoUpdate<T>): Promise<UpdateResult> => {
      const id = (filter as { _id?: string })._id
      const doc = id ? data.get(id) : null
      if (!doc) {
        return { matchedCount: 0, modifiedCount: 0, acknowledged: true }
      }
      if (update.$set) {
        Object.assign(doc as object, update.$set)
      }
      return { matchedCount: 1, modifiedCount: 1, acknowledged: true }
    }),

    updateMany: vi.fn(async (_filter: MongoFilter, _update: MongoUpdate<T>): Promise<UpdateResult> => {
      return { matchedCount: 0, modifiedCount: 0, acknowledged: true }
    }),

    deleteOne: vi.fn(async (filter: MongoFilter): Promise<DeleteResult> => {
      const id = (filter as { _id?: string })._id
      if (id && data.has(id)) {
        data.delete(id)
        return { deletedCount: 1, acknowledged: true }
      }
      return { deletedCount: 0, acknowledged: true }
    }),

    deleteMany: vi.fn(async (filter: MongoFilter): Promise<DeleteResult> => {
      let count = 0
      for (const [id, doc] of data.entries()) {
        let matches = true
        for (const [key, value] of Object.entries(filter)) {
          if ((doc as Record<string, unknown>)[key] !== value) {
            matches = false
            break
          }
        }
        if (matches) {
          data.delete(id)
          count++
        }
      }
      return { deletedCount: count, acknowledged: true }
    }),

    countDocuments: vi.fn(async (filter?: MongoFilter): Promise<number> => {
      if (!filter || Object.keys(filter).length === 0) {
        return data.size
      }
      let count = 0
      for (const doc of data.values()) {
        let matches = true
        for (const [key, value] of Object.entries(filter)) {
          if ((doc as Record<string, unknown>)[key] !== value) {
            matches = false
            break
          }
        }
        if (matches) count++
      }
      return count
    }),

    createIndex: vi.fn(async (_spec: IndexSpec, options?: IndexOptions): Promise<string> => {
      return options?.name || 'index_1'
    }),

    dropIndex: vi.fn(async (_name: string): Promise<void> => {}),

    listIndexes: vi.fn(async (): Promise<IndexInfo[]> => {
      return [{ name: '_id_', key: { _id: 1 } }]
    }),
  }
}

function createMockDatabase(): MongoDatabase {
  const collections = new Map<string, MongoCollection<unknown>>()

  return {
    collection: <T>(name: string): MongoCollection<T> => {
      if (!collections.has(name)) {
        collections.set(name, createMockCollection<T>())
      }
      return collections.get(name) as MongoCollection<T>
    },
  }
}

function createMockSession(): MongoSession {
  return {
    startTransaction: vi.fn(),
    commitTransaction: vi.fn(async () => {}),
    abortTransaction: vi.fn(async () => {}),
    endSession: vi.fn(async () => {}),
  }
}

function createMockClient(): MongoClient {
  const db = createMockDatabase()
  const session = createMockSession()

  return {
    db: vi.fn((_name?: string) => db),
    startSession: vi.fn(async () => session),
    close: vi.fn(async () => {}),
  }
}

// ============================================================================
// CellStore Tests
// ============================================================================

describe('CellStore', () => {
  let mockClient: MongoClient
  let mockDb: MongoDatabase
  let cellStore: CellStore

  beforeEach(() => {
    mockClient = createMockClient()
    mockDb = mockClient.db('test')
    cellStore = new CellStore(mockDb)
  })

  describe('getCell', () => {
    it('should return null for non-existent cell', async () => {
      const result = await cellStore.getCell('sheet1', 'A1')
      expect(result).toBeNull()
    })

    it('should return cell data for existing cell', async () => {
      // First set a cell
      await cellStore.setCell('sheet1', 'A1', 42)

      const result = await cellStore.getCell('sheet1', 'A1')
      expect(result).not.toBeNull()
      expect(result?.value.v).toBe(42)
      expect(result?.value.t).toBe('number')
      expect(result?.sheet).toBe('sheet1')
      expect(result?.col).toBe('A')
      expect(result?.row).toBe(1)
    })

    it('should handle cell references with multiple column letters', async () => {
      await cellStore.setCell('sheet1', 'AA100', 'test')

      const result = await cellStore.getCell('sheet1', 'AA100')
      expect(result).not.toBeNull()
      expect(result?.col).toBe('AA')
      expect(result?.row).toBe(100)
    })
  })

  describe('setCell', () => {
    it('should create a new cell with primitive value', async () => {
      const result = await cellStore.setCell('sheet1', 'A1', 42)

      expect(result._id).toBeDefined()
      expect(result.sheet).toBe('sheet1')
      expect(result.col).toBe('A')
      expect(result.row).toBe(1)
      expect(result.value.v).toBe(42)
      expect(result.value.t).toBe('number')
      expect(result.createdAt).toBeInstanceOf(Date)
      expect(result.updatedAt).toBeInstanceOf(Date)
    })

    it('should create a cell with string value', async () => {
      const result = await cellStore.setCell('sheet1', 'B2', 'Hello')

      expect(result.value.v).toBe('Hello')
      expect(result.value.t).toBe('string')
    })

    it('should create a cell with boolean value', async () => {
      const result = await cellStore.setCell('sheet1', 'C3', true)

      expect(result.value.v).toBe(true)
      expect(result.value.t).toBe('boolean')
    })

    it('should create a cell with date value', async () => {
      const date = new Date('2024-01-15')
      const result = await cellStore.setCell('sheet1', 'D4', date)

      expect(result.value.v).toEqual(date)
      expect(result.value.t).toBe('date')
    })

    it('should create a cell with null value (empty)', async () => {
      const result = await cellStore.setCell('sheet1', 'E5', null)

      expect(result.value.v).toBeNull()
      expect(result.value.t).toBe('empty')
    })

    it('should create a cell with CellData including formula', async () => {
      const cellData: CellData = {
        ref: 'F6',
        formula: '=SUM(A1:A5)',
        value: 100,
      }
      const result = await cellStore.setCell('sheet1', 'F6', cellData)

      expect(result.value.v).toBe(100)
      expect(result.value.f).toBe('=SUM(A1:A5)')
    })

    it('should create a cell with CellData including format', async () => {
      const cellData: CellData = {
        ref: 'G7',
        value: 1234.56,
        format: {
          numberFormat: '$#,##0.00',
          font: { bold: true },
        },
      }
      const result = await cellStore.setCell('sheet1', 'G7', cellData)

      expect(result.value.v).toBe(1234.56)
      expect(result.format?.numberFormat).toBe('$#,##0.00')
      expect(result.format?.font?.bold).toBe(true)
    })

    it('should update existing cell', async () => {
      await cellStore.setCell('sheet1', 'A1', 10)
      const updated = await cellStore.setCell('sheet1', 'A1', 20)

      expect(updated.value.v).toBe(20)

      const fetched = await cellStore.getCell('sheet1', 'A1')
      expect(fetched?.value.v).toBe(20)
    })

    it('should calculate correct colIndex', async () => {
      const resultA = await cellStore.setCell('sheet1', 'A1', 1)
      expect(resultA.colIndex).toBe(0)

      const resultZ = await cellStore.setCell('sheet1', 'Z1', 1)
      expect(resultZ.colIndex).toBe(25)

      const resultAA = await cellStore.setCell('sheet1', 'AA1', 1)
      expect(resultAA.colIndex).toBe(26)
    })
  })

  describe('getCells', () => {
    beforeEach(async () => {
      // Set up test data
      await cellStore.setCell('sheet1', 'A1', 1)
      await cellStore.setCell('sheet1', 'A2', 2)
      await cellStore.setCell('sheet1', 'B1', 3)
      await cellStore.setCell('sheet1', 'B2', 4)
      await cellStore.setCell('sheet1', 'C3', 5)
    })

    it('should return cells in range specified as string', async () => {
      const cells = await cellStore.getCells('sheet1', 'A1:B2')

      expect(cells).toHaveLength(4)
      expect(cells.map((c) => c.value.v)).toContain(1)
      expect(cells.map((c) => c.value.v)).toContain(2)
      expect(cells.map((c) => c.value.v)).toContain(3)
      expect(cells.map((c) => c.value.v)).toContain(4)
    })

    it('should return cells in range specified as Range object', async () => {
      const range: Range = {
        startCol: 'A',
        startRow: 1,
        endCol: 'B',
        endRow: 2,
      }
      const cells = await cellStore.getCells('sheet1', range)

      expect(cells).toHaveLength(4)
    })

    it('should return empty array for range with no cells', async () => {
      const cells = await cellStore.getCells('sheet1', 'D1:E2')

      expect(cells).toHaveLength(0)
    })

    it('should return cells in correct order (row-major)', async () => {
      const cells = await cellStore.getCells('sheet1', 'A1:B2')

      // Cells should be ordered by row, then by column
      expect(cells[0].row).toBeLessThanOrEqual(cells[1].row)
    })

    it('should handle single cell range', async () => {
      const cells = await cellStore.getCells('sheet1', 'A1:A1')

      expect(cells).toHaveLength(1)
      expect(cells[0].value.v).toBe(1)
    })

    it('should handle entire column range', async () => {
      const cells = await cellStore.getCells('sheet1', 'A1:A100')

      expect(cells).toHaveLength(2) // A1 and A2
    })

    it('should handle entire row range', async () => {
      const cells = await cellStore.getCells('sheet1', 'A1:Z1')

      expect(cells).toHaveLength(2) // A1 and B1
    })
  })

  describe('setCells', () => {
    it('should set multiple cells at once', async () => {
      const cellsData: CellData[] = [
        { ref: 'A1', value: 10 },
        { ref: 'A2', value: 20 },
        { ref: 'A3', value: 30 },
      ]

      const results = await cellStore.setCells('sheet1', cellsData)

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(10)
      expect(results[1].value.v).toBe(20)
      expect(results[2].value.v).toBe(30)
    })

    it('should handle cells with formulas in bulk', async () => {
      const cellsData: CellData[] = [
        { ref: 'A1', value: 10 },
        { ref: 'A2', value: 20 },
        { ref: 'A3', formula: '=SUM(A1:A2)', value: 30 },
      ]

      const results = await cellStore.setCells('sheet1', cellsData)

      expect(results[2].value.f).toBe('=SUM(A1:A2)')
    })

    it('should handle cells with different value types', async () => {
      const cellsData: CellData[] = [
        { ref: 'A1', value: 42 },
        { ref: 'A2', value: 'text' },
        { ref: 'A3', value: true },
        { ref: 'A4', value: null },
      ]

      const results = await cellStore.setCells('sheet1', cellsData)

      expect(results[0].value.t).toBe('number')
      expect(results[1].value.t).toBe('string')
      expect(results[2].value.t).toBe('boolean')
      expect(results[3].value.t).toBe('empty')
    })

    it('should return empty array for empty input', async () => {
      const results = await cellStore.setCells('sheet1', [])

      expect(results).toHaveLength(0)
    })

    it('should update existing cells in bulk', async () => {
      await cellStore.setCell('sheet1', 'A1', 1)
      await cellStore.setCell('sheet1', 'A2', 2)

      const cellsData: CellData[] = [
        { ref: 'A1', value: 100 },
        { ref: 'A2', value: 200 },
      ]

      await cellStore.setCells('sheet1', cellsData)

      const a1 = await cellStore.getCell('sheet1', 'A1')
      const a2 = await cellStore.getCell('sheet1', 'A2')

      expect(a1?.value.v).toBe(100)
      expect(a2?.value.v).toBe(200)
    })
  })

  describe('deleteCell', () => {
    it('should delete an existing cell and return true', async () => {
      await cellStore.setCell('sheet1', 'A1', 42)

      const result = await cellStore.deleteCell('sheet1', 'A1')

      expect(result).toBe(true)

      const cell = await cellStore.getCell('sheet1', 'A1')
      expect(cell).toBeNull()
    })

    it('should return false when deleting non-existent cell', async () => {
      const result = await cellStore.deleteCell('sheet1', 'Z99')

      expect(result).toBe(false)
    })
  })

  describe('deleteCells', () => {
    beforeEach(async () => {
      await cellStore.setCell('sheet1', 'A1', 1)
      await cellStore.setCell('sheet1', 'A2', 2)
      await cellStore.setCell('sheet1', 'B1', 3)
      await cellStore.setCell('sheet1', 'B2', 4)
    })

    it('should delete cells in range and return count', async () => {
      const count = await cellStore.deleteCells('sheet1', 'A1:B2')

      expect(count).toBe(4)
    })

    it('should return 0 for empty range', async () => {
      const count = await cellStore.deleteCells('sheet1', 'C1:D2')

      expect(count).toBe(0)
    })
  })

  describe('findCells', () => {
    beforeEach(async () => {
      await cellStore.setCell('sheet1', 'A1', 10)
      await cellStore.setCell('sheet1', 'A2', 20)
      await cellStore.setCell('sheet1', 'B1', 'hello')
      await cellStore.setCell('sheet1', 'B2', 'world')
      await cellStore.setCell('sheet1', 'C1', { ref: 'C1', formula: '=SUM(A1:A2)', value: 30 })
    })

    it('should find cells by exact value', async () => {
      const query: CellQuery = { value: 10 }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(1)
      expect(cells[0].value.v).toBe(10)
    })

    it('should find cells by value type', async () => {
      const query: CellQuery = { valueType: 'string' }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(2)
      expect(cells.every((c) => c.value.t === 'string')).toBe(true)
    })

    it('should find cells with formulas', async () => {
      const query: CellQuery = { hasFormula: true }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(1)
      expect(cells[0].value.f).toBe('=SUM(A1:A2)')
    })

    it('should find cells by formula pattern', async () => {
      const query: CellQuery = { formulaPattern: 'SUM' }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(1)
      expect(cells[0].value.f).toContain('SUM')
    })

    it('should find cells in row range', async () => {
      const query: CellQuery = { rowRange: { min: 1, max: 1 } }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(3) // A1, B1, C1
      expect(cells.every((c) => c.row === 1)).toBe(true)
    })

    it('should find cells in column range', async () => {
      const query: CellQuery = { colRange: { minIndex: 0, maxIndex: 0 } }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(2) // A1, A2
      expect(cells.every((c) => c.col === 'A')).toBe(true)
    })

    it('should combine multiple query conditions', async () => {
      const query: CellQuery = {
        valueType: 'number',
        rowRange: { min: 1, max: 1 },
      }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(1)
      expect(cells[0].value.v).toBe(10)
    })

    it('should return empty array when no matches', async () => {
      const query: CellQuery = { value: 'nonexistent' }
      const cells = await cellStore.findCells('sheet1', query)

      expect(cells).toHaveLength(0)
    })
  })

  describe('countCells', () => {
    beforeEach(async () => {
      await cellStore.setCell('sheet1', 'A1', 10)
      await cellStore.setCell('sheet1', 'A2', 20)
      await cellStore.setCell('sheet1', 'B1', 'text')
    })

    it('should count all cells in sheet', async () => {
      const count = await cellStore.countCells('sheet1')

      expect(count).toBe(3)
    })

    it('should count cells matching query', async () => {
      const count = await cellStore.countCells('sheet1', { valueType: 'number' })

      expect(count).toBe(2)
    })

    it('should return 0 for empty sheet', async () => {
      const count = await cellStore.countCells('emptySheet')

      expect(count).toBe(0)
    })
  })
})

// ============================================================================
// SheetStore Tests
// ============================================================================

describe('SheetStore', () => {
  let mockClient: MongoClient
  let mockDb: MongoDatabase
  let sheetStore: SheetStore

  beforeEach(() => {
    mockClient = createMockClient()
    mockDb = mockClient.db('test')
    sheetStore = new SheetStore(mockDb)
  })

  describe('createSheet', () => {
    it('should create a new sheet with required fields', async () => {
      const sheet = await sheetStore.createSheet('workbook1', 'Sheet1')

      expect(sheet._id).toBeDefined()
      expect(sheet.workbookId).toBe('workbook1')
      expect(sheet.name).toBe('Sheet1')
      expect(sheet.index).toBe(0)
      expect(sheet.createdAt).toBeInstanceOf(Date)
      expect(sheet.updatedAt).toBeInstanceOf(Date)
    })

    it('should create a sheet with custom index', async () => {
      const sheet = await sheetStore.createSheet('workbook1', 'Sheet2', { index: 5 })

      expect(sheet.index).toBe(5)
    })

    it('should create a sheet with default dimensions', async () => {
      const sheet = await sheetStore.createSheet('workbook1', 'Sheet1', {
        defaultColWidth: 100,
        defaultRowHeight: 25,
      })

      expect(sheet.defaultColWidth).toBe(100)
      expect(sheet.defaultRowHeight).toBe(25)
    })

    it('should auto-increment index for subsequent sheets', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet1')
      const sheet2 = await sheetStore.createSheet('workbook1', 'Sheet2')

      expect(sheet2.index).toBe(1)
    })
  })

  describe('getSheet', () => {
    it('should return null for non-existent sheet', async () => {
      const sheet = await sheetStore.getSheet('nonexistent')

      expect(sheet).toBeNull()
    })

    it('should return sheet by ID', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const sheet = await sheetStore.getSheet(created._id)

      expect(sheet).not.toBeNull()
      expect(sheet?._id).toBe(created._id)
      expect(sheet?.name).toBe('Sheet1')
    })
  })

  describe('getSheetByName', () => {
    it('should return null for non-existent sheet name', async () => {
      const sheet = await sheetStore.getSheetByName('workbook1', 'NonExistent')

      expect(sheet).toBeNull()
    })

    it('should return sheet by name within workbook', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet1')
      await sheetStore.createSheet('workbook1', 'Sheet2')

      const sheet = await sheetStore.getSheetByName('workbook1', 'Sheet2')

      expect(sheet).not.toBeNull()
      expect(sheet?.name).toBe('Sheet2')
    })

    it('should not return sheet from different workbook', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet1')
      await sheetStore.createSheet('workbook2', 'Sheet1')

      const sheet = await sheetStore.getSheetByName('workbook1', 'Sheet1')

      expect(sheet?.workbookId).toBe('workbook1')
    })
  })

  describe('updateSheet', () => {
    it('should return null for non-existent sheet', async () => {
      const result = await sheetStore.updateSheet('nonexistent', { name: 'NewName' })

      expect(result).toBeNull()
    })

    it('should update sheet name', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, { name: 'RenamedSheet' })

      expect(updated?.name).toBe('RenamedSheet')
    })

    it('should update sheet index', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, { index: 10 })

      expect(updated?.index).toBe(10)
    })

    it('should update column widths', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, {
        colWidths: { 0: 150, 1: 200 },
      })

      expect(updated?.colWidths).toEqual({ 0: 150, 1: 200 })
    })

    it('should update row heights', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, {
        rowHeights: { 1: 30, 2: 40 },
      })

      expect(updated?.rowHeights).toEqual({ 1: 30, 2: 40 })
    })

    it('should update freeze panes', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, {
        freeze: { rows: 1, cols: 2 },
      })

      expect(updated?.freeze).toEqual({ rows: 1, cols: 2 })
    })

    it('should update hidden columns and rows', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const updated = await sheetStore.updateSheet(created._id, {
        hiddenCols: [0, 2],
        hiddenRows: [5, 10],
      })

      expect(updated?.hiddenCols).toEqual([0, 2])
      expect(updated?.hiddenRows).toEqual([5, 10])
    })

    it('should update updatedAt timestamp', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')
      const originalUpdatedAt = created.updatedAt

      // Small delay to ensure different timestamp
      await new Promise((resolve) => setTimeout(resolve, 10))

      const updated = await sheetStore.updateSheet(created._id, { name: 'NewName' })

      expect(updated?.updatedAt.getTime()).toBeGreaterThan(originalUpdatedAt.getTime())
    })
  })

  describe('deleteSheet', () => {
    it('should return false for non-existent sheet', async () => {
      const result = await sheetStore.deleteSheet('nonexistent')

      expect(result).toBe(false)
    })

    it('should delete existing sheet and return true', async () => {
      const created = await sheetStore.createSheet('workbook1', 'Sheet1')

      const result = await sheetStore.deleteSheet(created._id)

      expect(result).toBe(true)

      const sheet = await sheetStore.getSheet(created._id)
      expect(sheet).toBeNull()
    })
  })

  describe('listSheets', () => {
    it('should return empty array for workbook with no sheets', async () => {
      const sheets = await sheetStore.listSheets('emptyWorkbook')

      expect(sheets).toHaveLength(0)
    })

    it('should return all sheets in workbook', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet1')
      await sheetStore.createSheet('workbook1', 'Sheet2')
      await sheetStore.createSheet('workbook1', 'Sheet3')

      const sheets = await sheetStore.listSheets('workbook1')

      expect(sheets).toHaveLength(3)
    })

    it('should return sheets ordered by index', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet3', { index: 2 })
      await sheetStore.createSheet('workbook1', 'Sheet1', { index: 0 })
      await sheetStore.createSheet('workbook1', 'Sheet2', { index: 1 })

      const sheets = await sheetStore.listSheets('workbook1')

      expect(sheets[0].name).toBe('Sheet1')
      expect(sheets[1].name).toBe('Sheet2')
      expect(sheets[2].name).toBe('Sheet3')
    })

    it('should only return sheets from specified workbook', async () => {
      await sheetStore.createSheet('workbook1', 'Sheet1')
      await sheetStore.createSheet('workbook2', 'Sheet2')

      const sheets = await sheetStore.listSheets('workbook1')

      expect(sheets).toHaveLength(1)
      expect(sheets[0].workbookId).toBe('workbook1')
    })
  })

  describe('reorderSheets', () => {
    it('should reorder sheets by updating indexes', async () => {
      const sheet1 = await sheetStore.createSheet('workbook1', 'Sheet1')
      const sheet2 = await sheetStore.createSheet('workbook1', 'Sheet2')
      const sheet3 = await sheetStore.createSheet('workbook1', 'Sheet3')

      // Reorder to: Sheet3, Sheet1, Sheet2
      await sheetStore.reorderSheets('workbook1', [sheet3._id, sheet1._id, sheet2._id])

      const sheets = await sheetStore.listSheets('workbook1')

      expect(sheets[0].name).toBe('Sheet3')
      expect(sheets[1].name).toBe('Sheet1')
      expect(sheets[2].name).toBe('Sheet2')
    })
  })
})

// ============================================================================
// WorkbookStore Tests
// ============================================================================

describe('WorkbookStore', () => {
  let mockClient: MongoClient
  let mockDb: MongoDatabase
  let workbookStore: WorkbookStore

  beforeEach(() => {
    mockClient = createMockClient()
    mockDb = mockClient.db('test')
    workbookStore = new WorkbookStore(mockDb)
  })

  describe('createWorkbook', () => {
    it('should create a new workbook with required fields', async () => {
      const workbook = await workbookStore.createWorkbook('My Workbook')

      expect(workbook._id).toBeDefined()
      expect(workbook.name).toBe('My Workbook')
      expect(workbook.createdAt).toBeInstanceOf(Date)
      expect(workbook.updatedAt).toBeInstanceOf(Date)
    })

    it('should create a workbook with properties', async () => {
      const workbook = await workbookStore.createWorkbook('My Workbook', {
        properties: {
          title: 'Sales Report',
          author: 'John Doe',
          subject: 'Q1 2024 Sales',
        },
      })

      expect(workbook.properties?.title).toBe('Sales Report')
      expect(workbook.properties?.author).toBe('John Doe')
      expect(workbook.properties?.subject).toBe('Q1 2024 Sales')
    })

    it('should initialize with activeSheet as 0', async () => {
      const workbook = await workbookStore.createWorkbook('My Workbook')

      expect(workbook.activeSheet).toBe(0)
    })
  })

  describe('getWorkbook', () => {
    it('should return null for non-existent workbook', async () => {
      const workbook = await workbookStore.getWorkbook('nonexistent')

      expect(workbook).toBeNull()
    })

    it('should return workbook by ID', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const workbook = await workbookStore.getWorkbook(created._id)

      expect(workbook).not.toBeNull()
      expect(workbook?._id).toBe(created._id)
      expect(workbook?.name).toBe('My Workbook')
    })
  })

  describe('updateWorkbook', () => {
    it('should return null for non-existent workbook', async () => {
      const result = await workbookStore.updateWorkbook('nonexistent', { name: 'NewName' })

      expect(result).toBeNull()
    })

    it('should update workbook name', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const updated = await workbookStore.updateWorkbook(created._id, { name: 'Renamed Workbook' })

      expect(updated?.name).toBe('Renamed Workbook')
    })

    it('should update active sheet', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const updated = await workbookStore.updateWorkbook(created._id, { activeSheet: 2 })

      expect(updated?.activeSheet).toBe(2)
    })

    it('should update named ranges', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const updated = await workbookStore.updateWorkbook(created._id, {
        namedRanges: {
          SalesData: 'Sheet1!A1:D100',
          TotalRow: 'Sheet1!A101:D101',
        },
      })

      expect(updated?.namedRanges?.SalesData).toBe('Sheet1!A1:D100')
      expect(updated?.namedRanges?.TotalRow).toBe('Sheet1!A101:D101')
    })

    it('should update styles', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const updated = await workbookStore.updateWorkbook(created._id, {
        styles: {
          header: { font: { bold: true, size: 14 } },
          currency: { numberFormat: '$#,##0.00' },
        },
      })

      expect(updated?.styles?.header?.font?.bold).toBe(true)
      expect(updated?.styles?.currency?.numberFormat).toBe('$#,##0.00')
    })

    it('should update updatedAt timestamp', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')
      const originalUpdatedAt = created.updatedAt

      await new Promise((resolve) => setTimeout(resolve, 10))

      const updated = await workbookStore.updateWorkbook(created._id, { name: 'NewName' })

      expect(updated?.updatedAt.getTime()).toBeGreaterThan(originalUpdatedAt.getTime())
    })
  })

  describe('deleteWorkbook', () => {
    it('should return false for non-existent workbook', async () => {
      const result = await workbookStore.deleteWorkbook('nonexistent')

      expect(result).toBe(false)
    })

    it('should delete existing workbook and return true', async () => {
      const created = await workbookStore.createWorkbook('My Workbook')

      const result = await workbookStore.deleteWorkbook(created._id)

      expect(result).toBe(true)

      const workbook = await workbookStore.getWorkbook(created._id)
      expect(workbook).toBeNull()
    })
  })

  describe('listWorkbooks', () => {
    it('should return empty array when no workbooks exist', async () => {
      const workbooks = await workbookStore.listWorkbooks()

      expect(workbooks).toHaveLength(0)
    })

    it('should return all workbooks', async () => {
      await workbookStore.createWorkbook('Workbook 1')
      await workbookStore.createWorkbook('Workbook 2')
      await workbookStore.createWorkbook('Workbook 3')

      const workbooks = await workbookStore.listWorkbooks()

      expect(workbooks).toHaveLength(3)
    })

    it('should return workbooks ordered by updatedAt descending', async () => {
      const wb1 = await workbookStore.createWorkbook('Workbook 1')
      await new Promise((resolve) => setTimeout(resolve, 10))
      const wb2 = await workbookStore.createWorkbook('Workbook 2')
      await new Promise((resolve) => setTimeout(resolve, 10))
      const wb3 = await workbookStore.createWorkbook('Workbook 3')

      const workbooks = await workbookStore.listWorkbooks()

      // Most recently updated should be first
      expect(workbooks[0].name).toBe('Workbook 3')
    })
  })
})

// ============================================================================
// StorageManager Tests
// ============================================================================

describe('StorageManager', () => {
  let mockClient: MongoClient
  let storageManager: StorageManager

  beforeEach(() => {
    mockClient = createMockClient()
    storageManager = new StorageManager(mockClient, 'testdb')
  })

  afterEach(async () => {
    await storageManager.close()
  })

  describe('constructor', () => {
    it('should initialize with cells, sheets, and workbooks stores', () => {
      expect(storageManager.cells).toBeInstanceOf(CellStore)
      expect(storageManager.sheets).toBeInstanceOf(SheetStore)
      expect(storageManager.workbooks).toBeInstanceOf(WorkbookStore)
    })

    it('should use default database name when not specified', () => {
      const manager = new StorageManager(mockClient)
      expect(mockClient.db).toHaveBeenCalledWith('excel')
    })

    it('should use specified database name', () => {
      const manager = new StorageManager(mockClient, 'customdb')
      expect(mockClient.db).toHaveBeenCalledWith('customdb')
    })
  })

  describe('withTransaction', () => {
    it('should execute function within a transaction', async () => {
      const fn = vi.fn(async (session: MongoSession) => {
        return 'result'
      })

      const result = await storageManager.withTransaction(fn)

      expect(result).toBe('result')
      expect(fn).toHaveBeenCalled()
    })

    it('should start transaction before function execution', async () => {
      const session = await mockClient.startSession()
      const fn = vi.fn(async () => 'result')

      await storageManager.withTransaction(fn)

      expect(session.startTransaction).toHaveBeenCalled()
    })

    it('should commit transaction on success', async () => {
      const session = await mockClient.startSession()
      const fn = vi.fn(async () => 'result')

      await storageManager.withTransaction(fn)

      expect(session.commitTransaction).toHaveBeenCalled()
    })

    it('should abort transaction on error', async () => {
      const session = await mockClient.startSession()
      const fn = vi.fn(async () => {
        throw new Error('Test error')
      })

      await expect(storageManager.withTransaction(fn)).rejects.toThrow('Test error')
      expect(session.abortTransaction).toHaveBeenCalled()
    })

    it('should end session after transaction', async () => {
      const session = await mockClient.startSession()
      const fn = vi.fn(async () => 'result')

      await storageManager.withTransaction(fn)

      expect(session.endSession).toHaveBeenCalled()
    })

    it('should support nested operations', async () => {
      const result = await storageManager.withTransaction(async (session) => {
        const workbook = await storageManager.workbooks.createWorkbook('Test')
        const sheet = await storageManager.sheets.createSheet(workbook._id, 'Sheet1')
        await storageManager.cells.setCell(sheet._id, 'A1', 100)

        return { workbook, sheet }
      })

      expect(result.workbook.name).toBe('Test')
      expect(result.sheet.name).toBe('Sheet1')
    })
  })

  describe('createIndexes', () => {
    it('should create all predefined indexes', async () => {
      await storageManager.createIndexes()

      // Verify cell indexes were created
      const cellsCollection = mockClient.db().collection('cells')
      expect(cellsCollection.createIndex).toHaveBeenCalled()

      // Verify sheet indexes were created
      const sheetsCollection = mockClient.db().collection('sheets')
      expect(sheetsCollection.createIndex).toHaveBeenCalled()

      // Verify workbook indexes were created
      const workbooksCollection = mockClient.db().collection('workbooks')
      expect(workbooksCollection.createIndex).toHaveBeenCalled()
    })
  })

  describe('close', () => {
    it('should close the client connection', async () => {
      await storageManager.close()

      expect(mockClient.close).toHaveBeenCalled()
    })
  })
})

// ============================================================================
// Transaction Tests
// ============================================================================

describe('Transaction Support', () => {
  let mockClient: MongoClient
  let storageManager: StorageManager

  beforeEach(() => {
    mockClient = createMockClient()
    storageManager = new StorageManager(mockClient, 'testdb')
  })

  describe('withTransaction', () => {
    it('should allow atomic multi-cell updates', async () => {
      await storageManager.withTransaction(async (session) => {
        await storageManager.cells.setCell('sheet1', 'A1', 100)
        await storageManager.cells.setCell('sheet1', 'A2', 200)
        await storageManager.cells.setCell('sheet1', 'A3', { ref: 'A3', formula: '=SUM(A1:A2)', value: 300 })
      })

      const a1 = await storageManager.cells.getCell('sheet1', 'A1')
      const a2 = await storageManager.cells.getCell('sheet1', 'A2')
      const a3 = await storageManager.cells.getCell('sheet1', 'A3')

      expect(a1?.value.v).toBe(100)
      expect(a2?.value.v).toBe(200)
      expect(a3?.value.f).toBe('=SUM(A1:A2)')
    })

    it('should rollback all changes on failure', async () => {
      // Set initial value
      await storageManager.cells.setCell('sheet1', 'A1', 'initial')

      try {
        await storageManager.withTransaction(async (session) => {
          await storageManager.cells.setCell('sheet1', 'A1', 'modified')
          throw new Error('Simulated failure')
        })
      } catch {
        // Expected
      }

      // Value should be rolled back
      const cell = await storageManager.cells.getCell('sheet1', 'A1')
      expect(cell?.value.v).toBe('initial')
    })

    it('should support cross-store transactions', async () => {
      await storageManager.withTransaction(async (session) => {
        const workbook = await storageManager.workbooks.createWorkbook('Transaction Test')
        const sheet = await storageManager.sheets.createSheet(workbook._id, 'Sheet1')
        await storageManager.cells.setCells(sheet._id, [
          { ref: 'A1', value: 1 },
          { ref: 'A2', value: 2 },
          { ref: 'A3', value: 3 },
        ])
      })

      const workbooks = await storageManager.workbooks.listWorkbooks()
      expect(workbooks.some((wb) => wb.name === 'Transaction Test')).toBe(true)
    })

    it('should return value from transaction function', async () => {
      const result = await storageManager.withTransaction(async (session) => {
        const workbook = await storageManager.workbooks.createWorkbook('Test')
        return workbook._id
      })

      expect(result).toBeDefined()
      expect(typeof result).toBe('string')
    })
  })
})

// ============================================================================
// Index Operations Tests
// ============================================================================

describe('Index Operations', () => {
  let mockClient: MongoClient
  let mockDb: MongoDatabase

  beforeEach(() => {
    mockClient = createMockClient()
    mockDb = mockClient.db('test')
  })

  describe('createIndex', () => {
    it('should create a simple index', async () => {
      const indexName = await createIndex(mockDb, 'cells', { sheet: 1 })

      expect(indexName).toBeDefined()
    })

    it('should create a compound index', async () => {
      const indexName = await createIndex(mockDb, 'cells', { sheet: 1, col: 1, row: 1 })

      expect(indexName).toBeDefined()
    })

    it('should create a unique index', async () => {
      const indexName = await createIndex(mockDb, 'cells', { _id: 1 }, { unique: true })

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { _id: 1 },
        expect.objectContaining({ unique: true })
      )
    })

    it('should create a sparse index', async () => {
      const indexName = await createIndex(
        mockDb,
        'cells',
        { 'value.f': 1 },
        { sparse: true }
      )

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { 'value.f': 1 },
        expect.objectContaining({ sparse: true })
      )
    })

    it('should create an index with custom name', async () => {
      const indexName = await createIndex(
        mockDb,
        'cells',
        { sheet: 1, row: 1 },
        { name: 'sheet_row_idx' }
      )

      expect(indexName).toBe('sheet_row_idx')
    })
  })

  describe('createCellIndexes', () => {
    it('should create all predefined cell indexes', async () => {
      await createCellIndexes(mockDb)

      const collection = mockDb.collection('cells')
      // Should create multiple indexes
      expect(collection.createIndex).toHaveBeenCalledTimes(6) // Based on CELL_INDEXES constant
    })
  })

  describe('Cell indexes for fast lookup', () => {
    let storageManager: StorageManager

    beforeEach(() => {
      storageManager = new StorageManager(mockClient, 'test')
    })

    it('should have index on sheet + col + row for unique cell lookup', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { sheet: 1, col: 1, row: 1 },
        expect.objectContaining({ unique: true })
      )
    })

    it('should have index on sheet + colIndex + row for range queries', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { sheet: 1, colIndex: 1, row: 1 },
        expect.anything()
      )
    })

    it('should have sparse index on formula field', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { sheet: 1, 'value.f': 1 },
        expect.objectContaining({ sparse: true })
      )
    })

    it('should have index on dependencies for dependency tracking', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { dependencies: 1 },
        expect.objectContaining({ sparse: true })
      )
    })

    it('should have index on dependents for reverse lookup', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { dependents: 1 },
        expect.objectContaining({ sparse: true })
      )
    })

    it('should have index on updatedAt for recent changes queries', async () => {
      await storageManager.createIndexes()

      const collection = mockDb.collection('cells')
      expect(collection.createIndex).toHaveBeenCalledWith(
        { updatedAt: -1 },
        expect.anything()
      )
    })
  })
})

// ============================================================================
// Utility Function Tests
// ============================================================================

describe('Utility Functions', () => {
  describe('parseCellRef', () => {
    it('should parse simple cell reference', () => {
      const result = parseCellRef('A1')

      expect(result.col).toBe('A')
      expect(result.row).toBe(1)
    })

    it('should parse cell reference with multiple column letters', () => {
      const result = parseCellRef('AA100')

      expect(result.col).toBe('AA')
      expect(result.row).toBe(100)
    })

    it('should parse cell reference with three column letters', () => {
      const result = parseCellRef('XFD1048576')

      expect(result.col).toBe('XFD')
      expect(result.row).toBe(1048576)
    })

    it('should handle lowercase column letters', () => {
      const result = parseCellRef('ab5')

      expect(result.col).toBe('AB')
      expect(result.row).toBe(5)
    })
  })

  describe('parseRange', () => {
    it('should parse simple range', () => {
      const result = parseRange('A1:B10')

      expect(result.startCol).toBe('A')
      expect(result.startRow).toBe(1)
      expect(result.endCol).toBe('B')
      expect(result.endRow).toBe(10)
    })

    it('should parse range with multiple column letters', () => {
      const result = parseRange('AA1:AZ100')

      expect(result.startCol).toBe('AA')
      expect(result.startRow).toBe(1)
      expect(result.endCol).toBe('AZ')
      expect(result.endRow).toBe(100)
    })

    it('should handle single cell range', () => {
      const result = parseRange('C5:C5')

      expect(result.startCol).toBe('C')
      expect(result.startRow).toBe(5)
      expect(result.endCol).toBe('C')
      expect(result.endRow).toBe(5)
    })
  })

  describe('generateCellId', () => {
    it('should generate correct cell ID', () => {
      const id = generateCellId('sheet1', 'A1')

      expect(id).toBe('sheet1!A1')
    })

    it('should handle sheet names with spaces', () => {
      const id = generateCellId('My Sheet', 'B2')

      expect(id).toBe("'My Sheet'!B2")
    })

    it('should escape quotes in sheet names', () => {
      const id = generateCellId("John's Data", 'C3')

      expect(id).toBe("'John''s Data'!C3")
    })
  })

  describe('generateId', () => {
    it('should generate unique IDs', () => {
      const id1 = generateId()
      const id2 = generateId()

      expect(id1).not.toBe(id2)
    })

    it('should generate string IDs', () => {
      const id = generateId()

      expect(typeof id).toBe('string')
      expect(id.length).toBeGreaterThan(0)
    })
  })
})

// ============================================================================
// Query Subscription Tests (q action)
// ============================================================================
// TDD RED: These tests define the API for query subscriptions.
// They should FAIL until the query subscription implementation is complete.

describe('Query Subscription (q action)', () => {
  let mockClient: MongoClient
  let mockDb: MongoDatabase
  let cellStore: CellStore

  beforeEach(async () => {
    mockClient = createMockClient()
    mockDb = mockClient.db('test')
    cellStore = new CellStore(mockDb)

    // Set up test data with various values for ordering/filtering
    await cellStore.setCell('sheet1', 'A1', { ref: 'A1', value: 10, metadata: { priority: 1, name: 'Alpha' } })
    await cellStore.setCell('sheet1', 'A2', { ref: 'A2', value: 30, metadata: { priority: 3, name: 'Charlie' } })
    await cellStore.setCell('sheet1', 'A3', { ref: 'A3', value: 20, metadata: { priority: 2, name: 'Bravo' } })
    await cellStore.setCell('sheet1', 'A4', { ref: 'A4', value: 50, metadata: { priority: 5, name: 'Echo' } })
    await cellStore.setCell('sheet1', 'A5', { ref: 'A5', value: 40, metadata: { priority: 4, name: 'Delta' } })
  })

  describe('orderByChild', () => {
    it('should order cells by child property value ascending', async () => {
      // Query cells ordered by metadata.priority
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.priority')

      const results = await query.get()

      expect(results).toHaveLength(5)
      expect(results[0].metadata?.priority).toBe(1)
      expect(results[1].metadata?.priority).toBe(2)
      expect(results[2].metadata?.priority).toBe(3)
      expect(results[3].metadata?.priority).toBe(4)
      expect(results[4].metadata?.priority).toBe(5)
    })

    it('should order cells by nested value property', async () => {
      // Query cells ordered by value.v (the actual cell value)
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')

      const results = await query.get()

      expect(results).toHaveLength(5)
      expect(results[0].value.v).toBe(10)
      expect(results[1].value.v).toBe(20)
      expect(results[2].value.v).toBe(30)
      expect(results[3].value.v).toBe(40)
      expect(results[4].value.v).toBe(50)
    })

    it('should order cells by string property alphabetically', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.name')

      const results = await query.get()

      expect(results).toHaveLength(5)
      expect(results[0].metadata?.name).toBe('Alpha')
      expect(results[1].metadata?.name).toBe('Bravo')
      expect(results[2].metadata?.name).toBe('Charlie')
      expect(results[3].metadata?.name).toBe('Delta')
      expect(results[4].metadata?.name).toBe('Echo')
    })

    it('should handle missing child properties', async () => {
      // Add a cell without the priority metadata
      await cellStore.setCell('sheet1', 'A6', { ref: 'A6', value: 60 })

      const query = cellStore.query('sheet1')
        .orderByChild('metadata.priority')

      const results = await query.get()

      // Cells with missing values should come first or last (consistent behavior)
      expect(results).toHaveLength(6)
    })

    it('should order by row number using special key', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('row')

      const results = await query.get()

      expect(results).toHaveLength(5)
      expect(results[0].row).toBe(1)
      expect(results[4].row).toBe(5)
    })
  })

  describe('limitToFirst', () => {
    it('should limit results to first N items', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(3)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(10)
      expect(results[1].value.v).toBe(20)
      expect(results[2].value.v).toBe(30)
    })

    it('should return all items when limit exceeds count', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(100)

      const results = await query.get()

      expect(results).toHaveLength(5)
    })

    it('should return empty array when limit is 0', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(0)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })

    it('should work with default ordering', async () => {
      const query = cellStore.query('sheet1')
        .limitToFirst(2)

      const results = await query.get()

      expect(results).toHaveLength(2)
    })
  })

  describe('limitToLast', () => {
    it('should limit results to last N items', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToLast(3)

      const results = await query.get()

      expect(results).toHaveLength(3)
      // Should be last 3 in ascending order: 30, 40, 50
      expect(results[0].value.v).toBe(30)
      expect(results[1].value.v).toBe(40)
      expect(results[2].value.v).toBe(50)
    })

    it('should return all items when limit exceeds count', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToLast(100)

      const results = await query.get()

      expect(results).toHaveLength(5)
    })

    it('should return empty array when limit is 0', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToLast(0)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })

    it('should work with default ordering', async () => {
      const query = cellStore.query('sheet1')
        .limitToLast(2)

      const results = await query.get()

      expect(results).toHaveLength(2)
    })
  })

  describe('startAt', () => {
    it('should filter to items starting at value (inclusive)', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(30)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(30)
      expect(results[1].value.v).toBe(40)
      expect(results[2].value.v).toBe(50)
    })

    it('should include exact matches', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(20)

      const results = await query.get()

      expect(results.some(r => r.value.v === 20)).toBe(true)
    })

    it('should work with string values', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.name')
        .startAt('Charlie')

      const results = await query.get()

      expect(results).toHaveLength(3) // Charlie, Delta, Echo
      expect(results[0].metadata?.name).toBe('Charlie')
    })

    it('should return empty when startAt exceeds all values', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(100)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })

    it('should work with child key for disambiguation', async () => {
      // startAt(value, key) for handling duplicate values
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(30, 'A3') // Starting at value 30, key A3

      const results = await query.get()

      expect(results.length).toBeGreaterThanOrEqual(1)
    })
  })

  describe('endAt', () => {
    it('should filter to items ending at value (inclusive)', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .endAt(30)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(10)
      expect(results[1].value.v).toBe(20)
      expect(results[2].value.v).toBe(30)
    })

    it('should include exact matches', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .endAt(40)

      const results = await query.get()

      expect(results.some(r => r.value.v === 40)).toBe(true)
    })

    it('should work with string values', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.name')
        .endAt('Charlie')

      const results = await query.get()

      expect(results).toHaveLength(3) // Alpha, Bravo, Charlie
      expect(results[2].metadata?.name).toBe('Charlie')
    })

    it('should return empty when endAt is below all values', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .endAt(5)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })

    it('should work with child key for disambiguation', async () => {
      // endAt(value, key) for handling duplicate values
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .endAt(30, 'A2')

      const results = await query.get()

      expect(results.length).toBeGreaterThanOrEqual(1)
    })
  })

  describe('startAt + endAt (range queries)', () => {
    it('should filter to items within range (inclusive)', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(20)
        .endAt(40)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(20)
      expect(results[1].value.v).toBe(30)
      expect(results[2].value.v).toBe(40)
    })

    it('should return single item when startAt equals endAt', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(30)
        .endAt(30)

      const results = await query.get()

      expect(results).toHaveLength(1)
      expect(results[0].value.v).toBe(30)
    })

    it('should return empty when range has no items', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(25)
        .endAt(28)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })

    it('should work with string ranges', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.name')
        .startAt('Bravo')
        .endAt('Delta')

      const results = await query.get()

      expect(results).toHaveLength(3) // Bravo, Charlie, Delta
    })
  })

  describe('equalTo', () => {
    it('should filter to items with exact value match', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .equalTo(30)

      const results = await query.get()

      expect(results).toHaveLength(1)
      expect(results[0].value.v).toBe(30)
    })

    it('should work with string values', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('metadata.name')
        .equalTo('Charlie')

      const results = await query.get()

      expect(results).toHaveLength(1)
      expect(results[0].metadata?.name).toBe('Charlie')
    })

    it('should return empty for non-existent value', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .equalTo(999)

      const results = await query.get()

      expect(results).toHaveLength(0)
    })
  })

  describe('combined query operations', () => {
    it('should combine orderByChild with limitToFirst', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(3)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(10)
      expect(results[2].value.v).toBe(30)
    })

    it('should combine orderByChild with limitToLast', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToLast(3)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(30)
      expect(results[2].value.v).toBe(50)
    })

    it('should combine startAt with limitToFirst', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(20)
        .limitToFirst(2)

      const results = await query.get()

      expect(results).toHaveLength(2)
      expect(results[0].value.v).toBe(20)
      expect(results[1].value.v).toBe(30)
    })

    it('should combine endAt with limitToLast', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .endAt(40)
        .limitToLast(2)

      const results = await query.get()

      expect(results).toHaveLength(2)
      expect(results[0].value.v).toBe(30)
      expect(results[1].value.v).toBe(40)
    })

    it('should combine startAt, endAt, and limitToFirst', async () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(10)
        .endAt(50)
        .limitToFirst(3)

      const results = await query.get()

      expect(results).toHaveLength(3)
      expect(results[0].value.v).toBe(10)
      expect(results[1].value.v).toBe(20)
      expect(results[2].value.v).toBe(30)
    })
  })

  describe('query subscription callbacks', () => {
    it('should call onValue callback with initial data', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(3)

      const unsubscribe = query.onValue(callback)

      // Wait for initial data
      await new Promise(resolve => setTimeout(resolve, 10))

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(expect.arrayContaining([
        expect.objectContaining({ value: expect.objectContaining({ v: 10 }) }),
      ]))

      unsubscribe()
    })

    it('should call onValue callback when data changes', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(3)

      const unsubscribe = query.onValue(callback)

      // Wait for initial callback
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).toHaveBeenCalledTimes(1)

      // Add a new cell that would affect the query
      await cellStore.setCell('sheet1', 'A6', { ref: 'A6', value: 5 })

      // Wait for update callback
      await new Promise(resolve => setTimeout(resolve, 10))

      // Should have been called again with updated data
      expect(callback).toHaveBeenCalledTimes(2)

      unsubscribe()
    })

    it('should stop receiving updates after unsubscribe', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .limitToFirst(3)

      const unsubscribe = query.onValue(callback)

      // Wait for initial callback
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).toHaveBeenCalledTimes(1)

      unsubscribe()

      // Add a new cell
      await cellStore.setCell('sheet1', 'A6', { ref: 'A6', value: 5 })

      // Wait for potential callback
      await new Promise(resolve => setTimeout(resolve, 10))

      // Should NOT have been called again
      expect(callback).toHaveBeenCalledTimes(1)
    })

    it('should call onChildAdded for new matching items', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(0)
        .endAt(100)

      const unsubscribe = query.onChildAdded(callback)

      // Wait for initial callbacks (one per existing child)
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).toHaveBeenCalledTimes(5)

      // Add a new cell within range
      await cellStore.setCell('sheet1', 'A6', { ref: 'A6', value: 25 })

      // Wait for callback
      await new Promise(resolve => setTimeout(resolve, 10))

      expect(callback).toHaveBeenCalledTimes(6)
      expect(callback).toHaveBeenLastCalledWith(
        expect.objectContaining({ value: expect.objectContaining({ v: 25 }) }),
        expect.any(String) // previous key
      )

      unsubscribe()
    })

    it('should call onChildChanged when item value changes', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')

      const unsubscribe = query.onChildChanged(callback)

      // Wait for setup
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).not.toHaveBeenCalled() // No changes yet

      // Update a cell
      await cellStore.setCell('sheet1', 'A1', { ref: 'A1', value: 15 })

      // Wait for callback
      await new Promise(resolve => setTimeout(resolve, 10))

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({ value: expect.objectContaining({ v: 15 }) }),
        expect.any(String) // previous key
      )

      unsubscribe()
    })

    it('should call onChildRemoved when item is deleted', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')

      const unsubscribe = query.onChildRemoved(callback)

      // Wait for setup
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).not.toHaveBeenCalled()

      // Delete a cell
      await cellStore.deleteCell('sheet1', 'A1')

      // Wait for callback
      await new Promise(resolve => setTimeout(resolve, 10))

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({ value: expect.objectContaining({ v: 10 }) })
      )

      unsubscribe()
    })

    it('should call onChildMoved when item order changes', async () => {
      const callback = vi.fn()

      const query = cellStore.query('sheet1')
        .orderByChild('value.v')

      const unsubscribe = query.onChildMoved(callback)

      // Wait for setup
      await new Promise(resolve => setTimeout(resolve, 10))
      expect(callback).not.toHaveBeenCalled()

      // Update A1 value to change its position in order
      await cellStore.setCell('sheet1', 'A1', { ref: 'A1', value: 100 })

      // Wait for callback
      await new Promise(resolve => setTimeout(resolve, 10))

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({ value: expect.objectContaining({ v: 100 }) }),
        expect.any(String) // previous key
      )

      unsubscribe()
    })
  })

  describe('query interface', () => {
    it('should have query method on CellStore', () => {
      expect(typeof cellStore.query).toBe('function')
    })

    it('should return a Query object with builder methods', () => {
      const query = cellStore.query('sheet1')

      expect(typeof query.orderByChild).toBe('function')
      expect(typeof query.limitToFirst).toBe('function')
      expect(typeof query.limitToLast).toBe('function')
      expect(typeof query.startAt).toBe('function')
      expect(typeof query.endAt).toBe('function')
      expect(typeof query.equalTo).toBe('function')
      expect(typeof query.get).toBe('function')
      expect(typeof query.onValue).toBe('function')
      expect(typeof query.onChildAdded).toBe('function')
      expect(typeof query.onChildChanged).toBe('function')
      expect(typeof query.onChildRemoved).toBe('function')
      expect(typeof query.onChildMoved).toBe('function')
    })

    it('should be chainable', () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')
        .startAt(10)
        .endAt(50)
        .limitToFirst(5)

      expect(query).toBeDefined()
      expect(typeof query.get).toBe('function')
    })

    it('should throw error when using multiple orderBy methods', () => {
      const query = cellStore.query('sheet1')
        .orderByChild('value.v')

      expect(() => query.orderByChild('metadata.priority')).toThrow()
    })

    it('should throw error when using limitToFirst and limitToLast together', () => {
      const query = cellStore.query('sheet1')
        .limitToFirst(5)

      expect(() => query.limitToLast(3)).toThrow()
    })
  })

  describe('orderByKey', () => {
    it('should order by cell key (ref)', async () => {
      // Add cells in random order
      await cellStore.setCell('sheet1', 'Z1', { ref: 'Z1', value: 1 })
      await cellStore.setCell('sheet1', 'B1', { ref: 'B1', value: 2 })
      await cellStore.setCell('sheet1', 'M1', { ref: 'M1', value: 3 })

      const query = cellStore.query('sheet1')
        .orderByKey()

      const results = await query.get()

      // Keys should be ordered: A1, A2, A3, A4, A5, B1, M1, Z1
      // (existing test data + new cells)
      expect(results.length).toBeGreaterThanOrEqual(3)
    })
  })

  describe('orderByValue', () => {
    it('should order by cell primitive value', async () => {
      const query = cellStore.query('sheet1')
        .orderByValue()

      const results = await query.get()

      // Should be ordered by value.v
      expect(results[0].value.v).toBe(10)
      expect(results[4].value.v).toBe(50)
    })
  })
})
