/**
 * RTDB Transaction Protocol Tests for database.do
 *
 * TDD RED: Tests for read-modify-write transaction flow.
 * These tests should FAIL until RTDB transaction protocol is implemented.
 *
 * RTDB (Real-Time Database) transaction protocol provides:
 * - Optimistic locking with version fields
 * - Atomic read-modify-write operations
 * - Conflict detection and automatic retry
 * - Snapshot isolation for consistent reads
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import {
  CellStore,
  SheetStore,
  WorkbookStore,
  StorageManager,
  TransactionError,
  ConflictError,
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
} from './types'
import type { Cell, CellPrimitive } from '../types'

// ============================================================================
// RTDB Transaction Types (to be implemented)
// ============================================================================

/**
 * Transaction options for RTDB protocol
 */
interface RTDBTransactionOptions {
  /** Maximum number of retry attempts on conflict */
  maxRetries?: number
  /** Timeout in milliseconds */
  timeout?: number
  /** Isolation level */
  isolation?: 'read-committed' | 'repeatable-read' | 'snapshot'
}

/**
 * Transaction context passed to transaction functions
 */
interface RTDBTransactionContext {
  /** Read a value with version tracking */
  read<T>(collection: string, id: string): Promise<{ value: T | null; version: number }>
  /** Write a value, checking version for conflicts */
  write<T>(collection: string, id: string, value: T, expectedVersion: number): Promise<void>
  /** Delete a value, checking version for conflicts */
  delete(collection: string, id: string, expectedVersion: number): Promise<void>
  /** Get current transaction timestamp */
  timestamp(): number
}

// ConflictError is now imported from ./mongo

/**
 * Extended storage manager with RTDB transaction support
 */
interface IRTDBStorageManager {
  readonly cells: CellStore
  readonly sheets: SheetStore
  readonly workbooks: WorkbookStore

  /**
   * Execute a read-modify-write transaction with automatic retry on conflict
   */
  runTransaction<T>(
    fn: (ctx: RTDBTransactionContext) => Promise<T>,
    options?: RTDBTransactionOptions
  ): Promise<T>

  /**
   * Read a cell with version information for transactions
   */
  readCellWithVersion(sheetId: string, cellRef: string): Promise<{ cell: Cell | null; version: number }>

  /**
   * Write a cell with version check (compare-and-swap)
   */
  writeCellWithVersion(
    sheetId: string,
    cellRef: string,
    value: CellPrimitive | CellData,
    expectedVersion: number
  ): Promise<Cell>

  /**
   * Atomic increment operation
   */
  incrementCell(sheetId: string, cellRef: string, delta: number): Promise<Cell>

  /**
   * Batch atomic operations
   */
  atomicBatch(operations: AtomicOperation[]): Promise<void>
}

/**
 * Atomic operation for batch processing
 */
interface AtomicOperation {
  type: 'set' | 'increment' | 'delete'
  sheetId: string
  cellRef: string
  value?: CellPrimitive | CellData
  delta?: number
  expectedVersion?: number
}

// ============================================================================
// Mock RTDB Client with Version Support
// ============================================================================

interface VersionedDocument<T> {
  data: T
  version: number
  updatedAt: Date
}

function createMockRTDBCollection<T>(): MongoCollection<T> & {
  _versions: Map<string, number>
  _getVersion(id: string): number
  _setVersion(id: string, version: number): void
} {
  const data = new Map<string, T>()
  const versions = new Map<string, number>()

  const collection = {
    _versions: versions,
    _getVersion(id: string): number {
      return versions.get(id) || 0
    },
    _setVersion(id: string, version: number): void {
      versions.set(id, version)
    },

    findOne: vi.fn(async (filter: MongoFilter): Promise<T | null> => {
      const id = (filter as { _id?: string })._id
      if (id) {
        return data.get(id) || null
      }
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
      versions.set(id, 1)
      return { insertedId: id, acknowledged: true }
    }),

    insertMany: vi.fn(async (docs: T[]): Promise<InsertManyResult> => {
      const ids: string[] = []
      for (const doc of docs) {
        const id = (doc as { _id?: string })._id || `mock-${Date.now()}-${Math.random()}`
        ;(doc as { _id: string })._id = id
        data.set(id, doc)
        versions.set(id, 1)
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
      // Increment version on update
      const currentVersion = versions.get(id!) || 0
      versions.set(id!, currentVersion + 1)
      return { matchedCount: 1, modifiedCount: 1, acknowledged: true }
    }),

    updateMany: vi.fn(async (_filter: MongoFilter, _update: MongoUpdate<T>): Promise<UpdateResult> => {
      return { matchedCount: 0, modifiedCount: 0, acknowledged: true }
    }),

    deleteOne: vi.fn(async (filter: MongoFilter): Promise<DeleteResult> => {
      const id = (filter as { _id?: string })._id
      if (id && data.has(id)) {
        data.delete(id)
        versions.delete(id)
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
          versions.delete(id)
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

  return collection as MongoCollection<T> & {
    _versions: Map<string, number>
    _getVersion(id: string): number
    _setVersion(id: string, version: number): void
  }
}

function createMockRTDBDatabase(): MongoDatabase {
  const collections = new Map<string, MongoCollection<unknown>>()

  return {
    collection: <T>(name: string): MongoCollection<T> => {
      if (!collections.has(name)) {
        collections.set(name, createMockRTDBCollection<T>())
      }
      return collections.get(name) as MongoCollection<T>
    },
  }
}

function createMockRTDBSession(): MongoSession {
  return {
    startTransaction: vi.fn(),
    commitTransaction: vi.fn(async () => {}),
    abortTransaction: vi.fn(async () => {}),
    endSession: vi.fn(async () => {}),
  }
}

function createMockRTDBClient(): MongoClient {
  const db = createMockRTDBDatabase()
  const session = createMockRTDBSession()

  return {
    db: vi.fn((_name?: string) => db),
    startSession: vi.fn(async () => session),
    close: vi.fn(async () => {}),
  }
}

// ============================================================================
// RTDB Transaction Protocol Tests
// ============================================================================

describe('RTDB Transaction Protocol', () => {
  let mockClient: MongoClient
  let storageManager: StorageManager

  beforeEach(() => {
    mockClient = createMockRTDBClient()
    storageManager = new StorageManager(mockClient, 'testdb')
  })

  afterEach(async () => {
    await storageManager.close()
  })

  describe('Read-Modify-Write Pattern', () => {
    it('should read cell with version information', async () => {
      // Set initial value
      await storageManager.cells.setCell('sheet1', 'A1', 100)

      // Read with version - this requires IRTDBStorageManager interface
      const rtdbManager = storageManager as unknown as IRTDBStorageManager
      const { cell, version } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      expect(cell).not.toBeNull()
      expect(cell?.value.v).toBe(100)
      expect(version).toBeGreaterThan(0)
      expect(typeof version).toBe('number')
    })

    it('should write cell with version check (compare-and-swap)', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      // Set initial value
      await storageManager.cells.setCell('sheet1', 'A1', 100)

      // Read with version
      const { version } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      // Write with correct version should succeed
      const updated = await rtdbManager.writeCellWithVersion('sheet1', 'A1', 200, version)

      expect(updated.value.v).toBe(200)
    })

    it('should reject write with stale version', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      // Set initial value
      await storageManager.cells.setCell('sheet1', 'A1', 100)

      // Read with version
      const { version } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      // Simulate concurrent modification by another client
      await storageManager.cells.setCell('sheet1', 'A1', 150)

      // Write with stale version should fail
      await expect(
        rtdbManager.writeCellWithVersion('sheet1', 'A1', 200, version)
      ).rejects.toThrow(ConflictError)
    })

    it('should provide expected and actual version in conflict error', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      const { version: oldVersion } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      // Concurrent modification
      await storageManager.cells.setCell('sheet1', 'A1', 150)
      const { version: newVersion } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      try {
        await rtdbManager.writeCellWithVersion('sheet1', 'A1', 200, oldVersion)
        expect.fail('Should have thrown ConflictError')
      } catch (error) {
        expect(error).toBeInstanceOf(ConflictError)
        const conflictError = error as ConflictError
        expect(conflictError.expectedVersion).toBe(oldVersion)
        expect(conflictError.actualVersion).toBe(newVersion)
      }
    })
  })

  describe('Automatic Transaction Retry', () => {
    it('should retry transaction on conflict', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager
      let attemptCount = 0

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      const result = await rtdbManager.runTransaction(
        async (ctx) => {
          attemptCount++
          const { value, version } = await ctx.read<Cell>('cells', 'sheet1!A1')

          if (attemptCount === 1) {
            // Simulate concurrent modification on first attempt
            await storageManager.cells.setCell('sheet1', 'A1', 150)
          }

          const newValue = (value?.value.v as number) + 10
          await ctx.write('cells', 'sheet1!A1', { ...value, value: { v: newValue, t: 'number' } }, version)

          return newValue
        },
        { maxRetries: 3 }
      )

      expect(attemptCount).toBe(2) // First attempt fails, second succeeds
      expect(result).toBe(160) // 150 + 10 after retry
    })

    it('should fail after max retries exceeded', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager
      let attemptCount = 0

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      await expect(
        rtdbManager.runTransaction(
          async (ctx) => {
            attemptCount++
            const { value, version } = await ctx.read<Cell>('cells', 'sheet1!A1')

            // Always simulate conflict
            await storageManager.cells.setCell('sheet1', 'A1', 100 + attemptCount)

            await ctx.write('cells', 'sheet1!A1', value, version)
          },
          { maxRetries: 3 }
        )
      ).rejects.toThrow(TransactionError)

      expect(attemptCount).toBe(4) // Initial + 3 retries
    })

    it('should not retry on non-conflict errors', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager
      let attemptCount = 0

      await expect(
        rtdbManager.runTransaction(
          async () => {
            attemptCount++
            throw new Error('Non-conflict error')
          },
          { maxRetries: 3 }
        )
      ).rejects.toThrow('Non-conflict error')

      expect(attemptCount).toBe(1) // No retries for non-conflict errors
    })
  })

  describe('Atomic Operations', () => {
    it('should atomically increment a cell value', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      const updated = await rtdbManager.incrementCell('sheet1', 'A1', 5)

      expect(updated.value.v).toBe(105)
      expect(updated.value.t).toBe('number')
    })

    it('should handle concurrent increments correctly', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      // Simulate concurrent increments
      const [result1, result2, result3] = await Promise.all([
        rtdbManager.incrementCell('sheet1', 'A1', 10),
        rtdbManager.incrementCell('sheet1', 'A1', 20),
        rtdbManager.incrementCell('sheet1', 'A1', 30),
      ])

      // All increments should be applied
      const final = await storageManager.cells.getCell('sheet1', 'A1')
      expect(final?.value.v).toBe(160) // 100 + 10 + 20 + 30
    })

    it('should throw error when incrementing non-numeric cell', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 'text')

      await expect(rtdbManager.incrementCell('sheet1', 'A1', 5)).rejects.toThrow(
        'Cannot increment non-numeric cell'
      )
    })

    it('should execute atomic batch operations', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      await storageManager.cells.setCell('sheet1', 'A2', 200)
      await storageManager.cells.setCell('sheet1', 'A3', 300)

      await rtdbManager.atomicBatch([
        { type: 'set', sheetId: 'sheet1', cellRef: 'A1', value: 110 },
        { type: 'increment', sheetId: 'sheet1', cellRef: 'A2', delta: 50 },
        { type: 'delete', sheetId: 'sheet1', cellRef: 'A3' },
      ])

      const a1 = await storageManager.cells.getCell('sheet1', 'A1')
      const a2 = await storageManager.cells.getCell('sheet1', 'A2')
      const a3 = await storageManager.cells.getCell('sheet1', 'A3')

      expect(a1?.value.v).toBe(110)
      expect(a2?.value.v).toBe(250)
      expect(a3).toBeNull()
    })

    it('should rollback batch on any failure', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      await storageManager.cells.setCell('sheet1', 'A2', 'not-a-number')

      await expect(
        rtdbManager.atomicBatch([
          { type: 'set', sheetId: 'sheet1', cellRef: 'A1', value: 110 },
          { type: 'increment', sheetId: 'sheet1', cellRef: 'A2', delta: 50 }, // Should fail
        ])
      ).rejects.toThrow()

      // A1 should be unchanged due to rollback
      const a1 = await storageManager.cells.getCell('sheet1', 'A1')
      expect(a1?.value.v).toBe(100)
    })
  })

  describe('Transaction Isolation', () => {
    it('should provide snapshot isolation within transaction', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      const result = await rtdbManager.runTransaction(
        async (ctx) => {
          const { value: read1 } = await ctx.read<Cell>('cells', 'sheet1!A1')

          // External modification during transaction
          await storageManager.cells.setCell('sheet1', 'A1', 999)

          // Should still see the snapshot value
          const { value: read2 } = await ctx.read<Cell>('cells', 'sheet1!A1')

          return {
            firstRead: read1?.value.v,
            secondRead: read2?.value.v,
          }
        },
        { isolation: 'snapshot' }
      )

      // Both reads should see the same value (snapshot isolation)
      expect(result.firstRead).toBe(100)
      expect(result.secondRead).toBe(100)
    })

    it('should see own writes within transaction', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      const result = await rtdbManager.runTransaction(async (ctx) => {
        const { value: initial, version } = await ctx.read<Cell>('cells', 'sheet1!A1')

        // Write within transaction
        await ctx.write(
          'cells',
          'sheet1!A1',
          { ...initial, value: { v: 200, t: 'number' } },
          version
        )

        // Should see own write
        const { value: afterWrite } = await ctx.read<Cell>('cells', 'sheet1!A1')

        return afterWrite?.value.v
      })

      expect(result).toBe(200)
    })

    it('should maintain read-your-writes consistency', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      const result = await rtdbManager.runTransaction(async (ctx) => {
        // Write new cell
        await ctx.write('cells', 'sheet1!B1', { value: { v: 'new', t: 'string' } }, 0)

        // Should be able to read the new cell
        const { value } = await ctx.read<Cell>('cells', 'sheet1!B1')

        return value?.value.v
      })

      expect(result).toBe('new')
    })
  })

  describe('Transaction Timeout', () => {
    it('should abort transaction on timeout', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await expect(
        rtdbManager.runTransaction(
          async () => {
            // Simulate long-running operation
            await new Promise((resolve) => setTimeout(resolve, 1000))
            return 'completed'
          },
          { timeout: 100 }
        )
      ).rejects.toThrow('Transaction timeout')
    })

    it('should cleanup resources on timeout', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      try {
        await rtdbManager.runTransaction(
          async (ctx) => {
            const { value, version } = await ctx.read<Cell>('cells', 'sheet1!A1')

            // Write that might be pending
            await ctx.write('cells', 'sheet1!A1', { ...value, value: { v: 200, t: 'number' } }, version)

            // Timeout before commit
            await new Promise((resolve) => setTimeout(resolve, 1000))
          },
          { timeout: 50 }
        )
      } catch {
        // Expected timeout
      }

      // Value should be unchanged (transaction was not committed)
      const cell = await storageManager.cells.getCell('sheet1', 'A1')
      expect(cell?.value.v).toBe(100)
    })
  })

  describe('Version Tracking', () => {
    it('should increment version on each update', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      const { version: v1 } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      await storageManager.cells.setCell('sheet1', 'A1', 200)
      const { version: v2 } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      await storageManager.cells.setCell('sheet1', 'A1', 300)
      const { version: v3 } = await rtdbManager.readCellWithVersion('sheet1', 'A1')

      expect(v2).toBeGreaterThan(v1)
      expect(v3).toBeGreaterThan(v2)
    })

    it('should return version 0 for non-existent cell', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      const { cell, version } = await rtdbManager.readCellWithVersion('sheet1', 'Z99')

      expect(cell).toBeNull()
      expect(version).toBe(0)
    })

    it('should allow create with version 0', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      const created = await rtdbManager.writeCellWithVersion('sheet1', 'A1', 100, 0)

      expect(created.value.v).toBe(100)

      const { version } = await rtdbManager.readCellWithVersion('sheet1', 'A1')
      expect(version).toBe(1)
    })

    it('should reject create if cell already exists', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      // Try to create (version 0) when cell exists
      await expect(rtdbManager.writeCellWithVersion('sheet1', 'A1', 200, 0)).rejects.toThrow(
        ConflictError
      )
    })
  })

  describe('Transaction Context', () => {
    it('should provide timestamp in transaction context', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      const timestamps: number[] = []

      await rtdbManager.runTransaction(async (ctx) => {
        timestamps.push(ctx.timestamp())
        await new Promise((resolve) => setTimeout(resolve, 10))
        timestamps.push(ctx.timestamp())
      })

      // Timestamps within same transaction should be the same (logical time)
      expect(timestamps[0]).toBe(timestamps[1])
    })

    it('should support delete in transaction context', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      await rtdbManager.runTransaction(async (ctx) => {
        const { version } = await ctx.read<Cell>('cells', 'sheet1!A1')
        await ctx.delete('cells', 'sheet1!A1', version)
      })

      const cell = await storageManager.cells.getCell('sheet1', 'A1')
      expect(cell).toBeNull()
    })

    it('should reject delete with stale version', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)

      await expect(
        rtdbManager.runTransaction(async (ctx) => {
          const { version } = await ctx.read<Cell>('cells', 'sheet1!A1')

          // Concurrent modification
          await storageManager.cells.setCell('sheet1', 'A1', 200)

          await ctx.delete('cells', 'sheet1!A1', version)
        })
      ).rejects.toThrow(ConflictError)
    })
  })

  describe('Multi-Cell Transactions', () => {
    it('should atomically update multiple cells', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      await storageManager.cells.setCell('sheet1', 'A2', 200)

      await rtdbManager.runTransaction(async (ctx) => {
        const { value: cell1, version: v1 } = await ctx.read<Cell>('cells', 'sheet1!A1')
        const { value: cell2, version: v2 } = await ctx.read<Cell>('cells', 'sheet1!A2')

        // Transfer value from A1 to A2
        const transfer = 50
        await ctx.write(
          'cells',
          'sheet1!A1',
          { ...cell1, value: { v: (cell1?.value.v as number) - transfer, t: 'number' } },
          v1
        )
        await ctx.write(
          'cells',
          'sheet1!A2',
          { ...cell2, value: { v: (cell2?.value.v as number) + transfer, t: 'number' } },
          v2
        )
      })

      const a1 = await storageManager.cells.getCell('sheet1', 'A1')
      const a2 = await storageManager.cells.getCell('sheet1', 'A2')

      expect(a1?.value.v).toBe(50)
      expect(a2?.value.v).toBe(250)
      // Total should remain constant (atomicity)
      expect((a1?.value.v as number) + (a2?.value.v as number)).toBe(300)
    })

    it('should rollback all cells on partial failure', async () => {
      const rtdbManager = storageManager as unknown as IRTDBStorageManager

      await storageManager.cells.setCell('sheet1', 'A1', 100)
      await storageManager.cells.setCell('sheet1', 'A2', 200)

      try {
        await rtdbManager.runTransaction(async (ctx) => {
          const { value: cell1, version: v1 } = await ctx.read<Cell>('cells', 'sheet1!A1')

          // Write first cell
          await ctx.write(
            'cells',
            'sheet1!A1',
            { ...cell1, value: { v: 999, t: 'number' } },
            v1
          )

          // Simulate concurrent modification of second cell
          await storageManager.cells.setCell('sheet1', 'A2', 300)

          const { version: staleV2 } = await ctx.read<Cell>('cells', 'sheet1!A2')

          // This should fail due to stale version
          await ctx.write(
            'cells',
            'sheet1!A2',
            { value: { v: 888, t: 'number' } },
            staleV2 - 1 // Force stale version
          )
        })
      } catch {
        // Expected
      }

      // Both cells should be at their original or concurrent-modified values
      const a1 = await storageManager.cells.getCell('sheet1', 'A1')
      const a2 = await storageManager.cells.getCell('sheet1', 'A2')

      // A1 should be rolled back
      expect(a1?.value.v).toBe(100)
      // A2 should have the concurrent modification
      expect(a2?.value.v).toBe(300)
    })
  })
})

describe('RTDB Formula Dependency Transactions', () => {
  let mockClient: MongoClient
  let storageManager: StorageManager

  beforeEach(() => {
    mockClient = createMockRTDBClient()
    storageManager = new StorageManager(mockClient, 'testdb')
  })

  afterEach(async () => {
    await storageManager.close()
  })

  it('should atomically update cell and its dependents', async () => {
    const rtdbManager = storageManager as unknown as IRTDBStorageManager

    // Set up cells with formula dependency
    await storageManager.cells.setCell('sheet1', 'A1', 10)
    await storageManager.cells.setCell('sheet1', 'A2', 20)
    await storageManager.cells.setCell('sheet1', 'A3', {
      ref: 'A3',
      formula: '=A1+A2',
      value: 30,
    })

    // Atomically update A1 and recalculate A3
    await rtdbManager.runTransaction(async (ctx) => {
      const { value: a1, version: v1 } = await ctx.read<Cell>('cells', 'sheet1!A1')
      const { value: a2 } = await ctx.read<Cell>('cells', 'sheet1!A2')
      const { value: a3, version: v3 } = await ctx.read<Cell>('cells', 'sheet1!A3')

      const newA1Value = 15
      const newA3Value = newA1Value + (a2?.value.v as number)

      await ctx.write('cells', 'sheet1!A1', { ...a1, value: { v: newA1Value, t: 'number' } }, v1)
      await ctx.write(
        'cells',
        'sheet1!A3',
        { ...a3, value: { v: newA3Value, t: 'number', f: '=A1+A2' } },
        v3
      )
    })

    const a1 = await storageManager.cells.getCell('sheet1', 'A1')
    const a3 = await storageManager.cells.getCell('sheet1', 'A3')

    expect(a1?.value.v).toBe(15)
    expect(a3?.value.v).toBe(35) // 15 + 20
  })

  it('should handle circular dependency detection in transaction', async () => {
    const rtdbManager = storageManager as unknown as IRTDBStorageManager

    await storageManager.cells.setCell('sheet1', 'A1', {
      ref: 'A1',
      formula: '=B1+1',
      value: 11,
    })
    await storageManager.cells.setCell('sheet1', 'B1', 10)

    // Try to create circular dependency: B1 depends on A1
    await expect(
      rtdbManager.runTransaction(async (ctx) => {
        const { value: b1, version: v1 } = await ctx.read<Cell>('cells', 'sheet1!B1')

        // This would create: A1 -> B1 -> A1 (circular)
        await ctx.write(
          'cells',
          'sheet1!B1',
          { ...b1, value: { v: 0, t: 'number', f: '=A1+1' } },
          v1
        )
      })
    ).rejects.toThrow('Circular dependency detected')
  })
})
