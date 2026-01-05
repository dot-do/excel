/**
 * Database Durable Object Child Event Tests
 *
 * TDD RED: Tests for child event differentiation in database.do.
 * These tests should FAIL until implementation is complete.
 *
 * Child events are Firebase-style real-time events:
 * - onChildAdded: Emitted when a new document is added to a collection
 * - onChildChanged: Emitted when an existing document is modified
 * - onChildRemoved: Emitted when a document is deleted from a collection
 * - onChildMoved: Emitted when a document's order/position changes
 */

import { describe, it, expect, beforeEach, vi } from 'vitest'
import type { Cell } from '../types'
import type { Sheet, Workbook } from './types'

// ============================================================================
// Types for Child Events
// ============================================================================

/** Event types for child event differentiation */
export type ChildEventType = 'child_added' | 'child_changed' | 'child_removed' | 'child_moved'

/** Base child event payload */
export interface ChildEvent<T = unknown> {
  /** Type of child event */
  type: ChildEventType
  /** The collection/path this event occurred in */
  collection: string
  /** The document key/ID */
  key: string
  /** Current value of the document (null for removed) */
  data: T | null
  /** Previous value before change (null for added) */
  previousData?: T | null
  /** Previous sibling key for ordering (for onChildMoved) */
  previousKey?: string | null
  /** New sibling key for ordering (for onChildMoved) */
  newPreviousKey?: string | null
  /** Timestamp of the event */
  timestamp: number
}

/** Callback function for child event listeners */
export type ChildEventCallback<T = unknown> = (event: ChildEvent<T>) => void

/** Interface for objects that emit child events */
export interface ChildEventEmitter {
  onChildAdded<T>(collection: string, callback: ChildEventCallback<T>): () => void
  onChildChanged<T>(collection: string, callback: ChildEventCallback<T>): () => void
  onChildRemoved<T>(collection: string, callback: ChildEventCallback<T>): () => void
  onChildMoved<T>(collection: string, callback: ChildEventCallback<T>): () => void
  off(collection: string, eventType?: ChildEventType): void
}

/** Interface for the DatabaseDO that will implement child events */
export interface DatabaseDO extends ChildEventEmitter {
  // Collection operations that should emit events
  insert<T>(collection: string, key: string, data: T): Promise<T>
  update<T>(collection: string, key: string, data: Partial<T>): Promise<T>
  delete(collection: string, key: string): Promise<boolean>
  move<T>(collection: string, key: string, newPosition: number): Promise<T>
  get<T>(collection: string, key: string): Promise<T | null>
  list<T>(collection: string): Promise<T[]>
}

// ============================================================================
// Mock DatabaseDO (Stub Implementation)
// ============================================================================

/**
 * Stub implementation that should fail all tests.
 * The real implementation will be in the GREEN phase.
 */
class StubDatabaseDO implements DatabaseDO {
  onChildAdded<T>(_collection: string, _callback: ChildEventCallback<T>): () => void {
    throw new Error('Not implemented: onChildAdded')
  }

  onChildChanged<T>(_collection: string, _callback: ChildEventCallback<T>): () => void {
    throw new Error('Not implemented: onChildChanged')
  }

  onChildRemoved<T>(_collection: string, _callback: ChildEventCallback<T>): () => void {
    throw new Error('Not implemented: onChildRemoved')
  }

  onChildMoved<T>(_collection: string, _callback: ChildEventCallback<T>): () => void {
    throw new Error('Not implemented: onChildMoved')
  }

  off(_collection: string, _eventType?: ChildEventType): void {
    throw new Error('Not implemented: off')
  }

  async insert<T>(_collection: string, _key: string, _data: T): Promise<T> {
    throw new Error('Not implemented: insert')
  }

  async update<T>(_collection: string, _key: string, _data: Partial<T>): Promise<T> {
    throw new Error('Not implemented: update')
  }

  async delete(_collection: string, _key: string): Promise<boolean> {
    throw new Error('Not implemented: delete')
  }

  async move<T>(_collection: string, _key: string, _newPosition: number): Promise<T> {
    throw new Error('Not implemented: move')
  }

  async get<T>(_collection: string, _key: string): Promise<T | null> {
    throw new Error('Not implemented: get')
  }

  async list<T>(_collection: string): Promise<T[]> {
    throw new Error('Not implemented: list')
  }
}

// ============================================================================
// onChildAdded Event Tests
// ============================================================================

describe('onChildAdded', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  describe('event registration', () => {
    it('should register a listener for child_added events', () => {
      const callback = vi.fn()

      expect(() => {
        db.onChildAdded('cells', callback)
      }).not.toThrow()
    })

    it('should return an unsubscribe function', () => {
      const callback = vi.fn()

      const unsubscribe = db.onChildAdded('cells', callback)

      expect(typeof unsubscribe).toBe('function')
    })

    it('should allow multiple listeners on the same collection', () => {
      const callback1 = vi.fn()
      const callback2 = vi.fn()

      expect(() => {
        db.onChildAdded('cells', callback1)
        db.onChildAdded('cells', callback2)
      }).not.toThrow()
    })

    it('should allow listeners on different collections', () => {
      const callback = vi.fn()

      expect(() => {
        db.onChildAdded('cells', callback)
        db.onChildAdded('sheets', callback)
        db.onChildAdded('workbooks', callback)
      }).not.toThrow()
    })
  })

  describe('event emission', () => {
    it('should emit child_added event when a new document is inserted', async () => {
      const callback = vi.fn()
      db.onChildAdded<Cell>('cells', callback)

      const cellData = {
        _id: 'sheet1!A1',
        sheet: 'sheet1',
        row: 1,
        col: 'A',
        colIndex: 0,
        value: { v: 42, t: 'number' as const },
        createdAt: new Date(),
        updatedAt: new Date(),
      }

      await db.insert('cells', 'sheet1!A1', cellData)

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          type: 'child_added',
          collection: 'cells',
          key: 'sheet1!A1',
          data: expect.objectContaining({ _id: 'sheet1!A1' }),
        })
      )
    })

    it('should include timestamp in child_added event', async () => {
      const callback = vi.fn()
      db.onChildAdded('cells', callback)

      const beforeTime = Date.now()
      await db.insert('cells', 'key1', { value: 'test' })
      const afterTime = Date.now()

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          timestamp: expect.any(Number),
        })
      )

      const event = callback.mock.calls[0][0] as ChildEvent
      expect(event.timestamp).toBeGreaterThanOrEqual(beforeTime)
      expect(event.timestamp).toBeLessThanOrEqual(afterTime)
    })

    it('should emit child_added for each document in bulk insert', async () => {
      const callback = vi.fn()
      db.onChildAdded('cells', callback)

      await db.insert('cells', 'cell1', { value: 1 })
      await db.insert('cells', 'cell2', { value: 2 })
      await db.insert('cells', 'cell3', { value: 3 })

      expect(callback).toHaveBeenCalledTimes(3)
    })

    it('should not emit child_added for updates to existing documents', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildAdded('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should have null previousData in child_added events', async () => {
      const callback = vi.fn()
      db.onChildAdded('cells', callback)

      await db.insert('cells', 'key1', { value: 'test' })

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          previousData: null,
        })
      )
    })
  })

  describe('unsubscribe behavior', () => {
    it('should stop receiving events after unsubscribe', async () => {
      const callback = vi.fn()
      const unsubscribe = db.onChildAdded('cells', callback)

      await db.insert('cells', 'cell1', { value: 1 })
      expect(callback).toHaveBeenCalledTimes(1)

      unsubscribe()

      await db.insert('cells', 'cell2', { value: 2 })
      expect(callback).toHaveBeenCalledTimes(1) // Still 1, not 2
    })

    it('should not affect other listeners when unsubscribing', async () => {
      const callback1 = vi.fn()
      const callback2 = vi.fn()

      const unsubscribe1 = db.onChildAdded('cells', callback1)
      db.onChildAdded('cells', callback2)

      unsubscribe1()

      await db.insert('cells', 'cell1', { value: 1 })

      expect(callback1).not.toHaveBeenCalled()
      expect(callback2).toHaveBeenCalledTimes(1)
    })
  })
})

// ============================================================================
// onChildChanged Event Tests
// ============================================================================

describe('onChildChanged', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  describe('event registration', () => {
    it('should register a listener for child_changed events', () => {
      const callback = vi.fn()

      expect(() => {
        db.onChildChanged('cells', callback)
      }).not.toThrow()
    })

    it('should return an unsubscribe function', () => {
      const callback = vi.fn()

      const unsubscribe = db.onChildChanged('cells', callback)

      expect(typeof unsubscribe).toBe('function')
    })
  })

  describe('event emission', () => {
    it('should emit child_changed event when a document is updated', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildChanged<{ value: number }>('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          type: 'child_changed',
          collection: 'cells',
          key: 'cell1',
          data: expect.objectContaining({ value: 2 }),
        })
      )
    })

    it('should include previousData in child_changed events', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1, name: 'test' })

      db.onChildChanged('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          previousData: expect.objectContaining({ value: 1, name: 'test' }),
          data: expect.objectContaining({ value: 2, name: 'test' }),
        })
      )
    })

    it('should emit child_changed for partial updates', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { a: 1, b: 2, c: 3 })

      db.onChildChanged('cells', callback)

      await db.update('cells', 'cell1', { b: 20 })

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          data: expect.objectContaining({ a: 1, b: 20, c: 3 }),
        })
      )
    })

    it('should not emit child_changed for new document inserts', async () => {
      const callback = vi.fn()
      db.onChildChanged('cells', callback)

      await db.insert('cells', 'cell1', { value: 1 })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should not emit child_changed for document deletions', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildChanged('cells', callback)

      await db.delete('cells', 'cell1')

      expect(callback).not.toHaveBeenCalled()
    })

    it('should emit separate child_changed events for multiple updates', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildChanged('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })
      await db.update('cells', 'cell1', { value: 3 })
      await db.update('cells', 'cell1', { value: 4 })

      expect(callback).toHaveBeenCalledTimes(3)
    })

    it('should not emit child_changed if data is unchanged', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildChanged('cells', callback)

      // Update with same value
      await db.update('cells', 'cell1', { value: 1 })

      expect(callback).not.toHaveBeenCalled()
    })
  })

  describe('unsubscribe behavior', () => {
    it('should stop receiving events after unsubscribe', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      const unsubscribe = db.onChildChanged('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })
      expect(callback).toHaveBeenCalledTimes(1)

      unsubscribe()

      await db.update('cells', 'cell1', { value: 3 })
      expect(callback).toHaveBeenCalledTimes(1)
    })
  })
})

// ============================================================================
// onChildRemoved Event Tests
// ============================================================================

describe('onChildRemoved', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  describe('event registration', () => {
    it('should register a listener for child_removed events', () => {
      const callback = vi.fn()

      expect(() => {
        db.onChildRemoved('cells', callback)
      }).not.toThrow()
    })

    it('should return an unsubscribe function', () => {
      const callback = vi.fn()

      const unsubscribe = db.onChildRemoved('cells', callback)

      expect(typeof unsubscribe).toBe('function')
    })
  })

  describe('event emission', () => {
    it('should emit child_removed event when a document is deleted', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildRemoved<{ value: number }>('cells', callback)

      await db.delete('cells', 'cell1')

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          type: 'child_removed',
          collection: 'cells',
          key: 'cell1',
        })
      )
    })

    it('should include the removed data in child_removed events', async () => {
      const callback = vi.fn()
      const originalData = { value: 42, name: 'test cell' }

      await db.insert('cells', 'cell1', originalData)

      db.onChildRemoved('cells', callback)

      await db.delete('cells', 'cell1')

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          data: null,
          previousData: expect.objectContaining(originalData),
        })
      )
    })

    it('should not emit child_removed for document inserts', async () => {
      const callback = vi.fn()
      db.onChildRemoved('cells', callback)

      await db.insert('cells', 'cell1', { value: 1 })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should not emit child_removed for document updates', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })

      db.onChildRemoved('cells', callback)

      await db.update('cells', 'cell1', { value: 2 })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should not emit child_removed for non-existent document deletion', async () => {
      const callback = vi.fn()
      db.onChildRemoved('cells', callback)

      await db.delete('cells', 'nonexistent')

      expect(callback).not.toHaveBeenCalled()
    })

    it('should emit child_removed for each deleted document', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })
      await db.insert('cells', 'cell2', { value: 2 })
      await db.insert('cells', 'cell3', { value: 3 })

      db.onChildRemoved('cells', callback)

      await db.delete('cells', 'cell1')
      await db.delete('cells', 'cell2')
      await db.delete('cells', 'cell3')

      expect(callback).toHaveBeenCalledTimes(3)
    })
  })

  describe('unsubscribe behavior', () => {
    it('should stop receiving events after unsubscribe', async () => {
      const callback = vi.fn()

      await db.insert('cells', 'cell1', { value: 1 })
      await db.insert('cells', 'cell2', { value: 2 })

      const unsubscribe = db.onChildRemoved('cells', callback)

      await db.delete('cells', 'cell1')
      expect(callback).toHaveBeenCalledTimes(1)

      unsubscribe()

      await db.delete('cells', 'cell2')
      expect(callback).toHaveBeenCalledTimes(1)
    })
  })
})

// ============================================================================
// onChildMoved Event Tests
// ============================================================================

describe('onChildMoved', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  describe('event registration', () => {
    it('should register a listener for child_moved events', () => {
      const callback = vi.fn()

      expect(() => {
        db.onChildMoved('sheets', callback)
      }).not.toThrow()
    })

    it('should return an unsubscribe function', () => {
      const callback = vi.fn()

      const unsubscribe = db.onChildMoved('sheets', callback)

      expect(typeof unsubscribe).toBe('function')
    })
  })

  describe('event emission', () => {
    it('should emit child_moved event when a document position changes', async () => {
      const callback = vi.fn()

      // Create ordered sheets
      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
      await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })
      await db.insert('sheets', 'sheet3', { name: 'Sheet3', index: 2 })

      db.onChildMoved<{ name: string; index: number }>('sheets', callback)

      // Move sheet1 to end (index 2)
      await db.move('sheets', 'sheet1', 2)

      expect(callback).toHaveBeenCalledTimes(1)
      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          type: 'child_moved',
          collection: 'sheets',
          key: 'sheet1',
        })
      )
    })

    it('should include previous position in child_moved events', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
      await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })
      await db.insert('sheets', 'sheet3', { name: 'Sheet3', index: 2 })

      db.onChildMoved('sheets', callback)

      await db.move('sheets', 'sheet3', 0)

      expect(callback).toHaveBeenCalledWith(
        expect.objectContaining({
          previousKey: 'sheet2', // Was after sheet2
          newPreviousKey: null, // Now at the beginning
        })
      )
    })

    it('should not emit child_moved for document inserts', async () => {
      const callback = vi.fn()
      db.onChildMoved('sheets', callback)

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should not emit child_moved for document updates that do not change position', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })

      db.onChildMoved('sheets', callback)

      await db.update('sheets', 'sheet1', { name: 'Renamed Sheet' })

      expect(callback).not.toHaveBeenCalled()
    })

    it('should not emit child_moved for document deletions', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })

      db.onChildMoved('sheets', callback)

      await db.delete('sheets', 'sheet1')

      expect(callback).not.toHaveBeenCalled()
    })

    it('should emit child_moved for reordering multiple documents', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
      await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })
      await db.insert('sheets', 'sheet3', { name: 'Sheet3', index: 2 })

      db.onChildMoved('sheets', callback)

      // Move sheet3 to beginning, sheet2 to end
      await db.move('sheets', 'sheet3', 0)
      await db.move('sheets', 'sheet2', 2)

      expect(callback).toHaveBeenCalledTimes(2)
    })

    it('should not emit child_moved when moved to same position', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
      await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })

      db.onChildMoved('sheets', callback)

      // Move to same position
      await db.move('sheets', 'sheet1', 0)

      expect(callback).not.toHaveBeenCalled()
    })
  })

  describe('unsubscribe behavior', () => {
    it('should stop receiving events after unsubscribe', async () => {
      const callback = vi.fn()

      await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
      await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })
      await db.insert('sheets', 'sheet3', { name: 'Sheet3', index: 2 })

      const unsubscribe = db.onChildMoved('sheets', callback)

      await db.move('sheets', 'sheet1', 2)
      expect(callback).toHaveBeenCalledTimes(1)

      unsubscribe()

      await db.move('sheets', 'sheet2', 0)
      expect(callback).toHaveBeenCalledTimes(1)
    })
  })
})

// ============================================================================
// Combined Event Listener Tests
// ============================================================================

describe('Combined Child Event Listeners', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should emit correct event types for different operations', async () => {
    const addedCallback = vi.fn()
    const changedCallback = vi.fn()
    const removedCallback = vi.fn()

    db.onChildAdded('cells', addedCallback)
    db.onChildChanged('cells', changedCallback)
    db.onChildRemoved('cells', removedCallback)

    // Insert
    await db.insert('cells', 'cell1', { value: 1 })
    expect(addedCallback).toHaveBeenCalledTimes(1)
    expect(changedCallback).not.toHaveBeenCalled()
    expect(removedCallback).not.toHaveBeenCalled()

    addedCallback.mockClear()

    // Update
    await db.update('cells', 'cell1', { value: 2 })
    expect(addedCallback).not.toHaveBeenCalled()
    expect(changedCallback).toHaveBeenCalledTimes(1)
    expect(removedCallback).not.toHaveBeenCalled()

    changedCallback.mockClear()

    // Delete
    await db.delete('cells', 'cell1')
    expect(addedCallback).not.toHaveBeenCalled()
    expect(changedCallback).not.toHaveBeenCalled()
    expect(removedCallback).toHaveBeenCalledTimes(1)
  })

  it('should handle listeners on different collections independently', async () => {
    const cellsCallback = vi.fn()
    const sheetsCallback = vi.fn()

    db.onChildAdded('cells', cellsCallback)
    db.onChildAdded('sheets', sheetsCallback)

    await db.insert('cells', 'cell1', { value: 1 })

    expect(cellsCallback).toHaveBeenCalledTimes(1)
    expect(sheetsCallback).not.toHaveBeenCalled()
  })

  it('should allow multiple event types to be registered for same collection', async () => {
    const addedCallback = vi.fn()
    const changedCallback = vi.fn()
    const removedCallback = vi.fn()
    const movedCallback = vi.fn()

    db.onChildAdded('sheets', addedCallback)
    db.onChildChanged('sheets', changedCallback)
    db.onChildRemoved('sheets', removedCallback)
    db.onChildMoved('sheets', movedCallback)

    await db.insert('sheets', 'sheet1', { name: 'Sheet1', index: 0 })
    await db.insert('sheets', 'sheet2', { name: 'Sheet2', index: 1 })

    expect(addedCallback).toHaveBeenCalledTimes(2)

    await db.update('sheets', 'sheet1', { name: 'Renamed' })
    expect(changedCallback).toHaveBeenCalledTimes(1)

    await db.move('sheets', 'sheet2', 0)
    expect(movedCallback).toHaveBeenCalledTimes(1)

    await db.delete('sheets', 'sheet1')
    expect(removedCallback).toHaveBeenCalledTimes(1)
  })
})

// ============================================================================
// off() Method Tests
// ============================================================================

describe('off() method', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should remove all listeners for a collection when called with collection only', async () => {
    const addedCallback = vi.fn()
    const changedCallback = vi.fn()
    const removedCallback = vi.fn()

    db.onChildAdded('cells', addedCallback)
    db.onChildChanged('cells', changedCallback)
    db.onChildRemoved('cells', removedCallback)

    db.off('cells')

    await db.insert('cells', 'cell1', { value: 1 })
    await db.update('cells', 'cell1', { value: 2 })
    await db.delete('cells', 'cell1')

    expect(addedCallback).not.toHaveBeenCalled()
    expect(changedCallback).not.toHaveBeenCalled()
    expect(removedCallback).not.toHaveBeenCalled()
  })

  it('should remove only specific event type listeners when eventType is provided', async () => {
    const addedCallback = vi.fn()
    const changedCallback = vi.fn()

    db.onChildAdded('cells', addedCallback)
    db.onChildChanged('cells', changedCallback)

    db.off('cells', 'child_added')

    await db.insert('cells', 'cell1', { value: 1 })
    expect(addedCallback).not.toHaveBeenCalled()

    await db.update('cells', 'cell1', { value: 2 })
    expect(changedCallback).toHaveBeenCalledTimes(1)
  })

  it('should not affect other collections when removing listeners', async () => {
    const cellsCallback = vi.fn()
    const sheetsCallback = vi.fn()

    db.onChildAdded('cells', cellsCallback)
    db.onChildAdded('sheets', sheetsCallback)

    db.off('cells')

    await db.insert('cells', 'cell1', { value: 1 })
    await db.insert('sheets', 'sheet1', { name: 'Sheet1' })

    expect(cellsCallback).not.toHaveBeenCalled()
    expect(sheetsCallback).toHaveBeenCalledTimes(1)
  })
})

// ============================================================================
// Event Ordering Tests
// ============================================================================

describe('Event Ordering', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should emit events in order of operations', async () => {
    const events: string[] = []

    db.onChildAdded('cells', () => events.push('added'))
    db.onChildChanged('cells', () => events.push('changed'))
    db.onChildRemoved('cells', () => events.push('removed'))

    await db.insert('cells', 'cell1', { value: 1 })
    await db.update('cells', 'cell1', { value: 2 })
    await db.delete('cells', 'cell1')

    expect(events).toEqual(['added', 'changed', 'removed'])
  })

  it('should emit events synchronously before operation resolves', async () => {
    const events: Array<{ event: string; time: number }> = []
    const start = Date.now()

    db.onChildAdded('cells', () => {
      events.push({ event: 'added', time: Date.now() - start })
    })

    await db.insert('cells', 'cell1', { value: 1 })
    const afterInsert = Date.now() - start

    // Event should have been recorded before the await resolved
    expect(events).toHaveLength(1)
    expect(events[0].time).toBeLessThanOrEqual(afterInsert)
  })
})

// ============================================================================
// Error Handling Tests
// ============================================================================

describe('Error Handling in Child Events', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should continue emitting events to other listeners if one throws', async () => {
    const callback1 = vi.fn(() => {
      throw new Error('Callback error')
    })
    const callback2 = vi.fn()

    db.onChildAdded('cells', callback1)
    db.onChildAdded('cells', callback2)

    await db.insert('cells', 'cell1', { value: 1 })

    expect(callback1).toHaveBeenCalled()
    expect(callback2).toHaveBeenCalled()
  })

  it('should not reject the operation promise if callback throws', async () => {
    const callback = vi.fn(() => {
      throw new Error('Callback error')
    })

    db.onChildAdded('cells', callback)

    await expect(db.insert('cells', 'cell1', { value: 1 })).resolves.toBeDefined()
  })
})

// ============================================================================
// Memory and Performance Tests
// ============================================================================

describe('Memory and Performance', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should not leak memory when listeners are unsubscribed', () => {
    const callbacks: Array<ChildEventCallback<unknown>> = []

    // Create many listeners
    for (let i = 0; i < 100; i++) {
      const callback = vi.fn()
      callbacks.push(callback)
      const unsubscribe = db.onChildAdded('cells', callback)
      unsubscribe()
    }

    // After unsubscribing, the callback list should be empty
    // This is implementation-dependent, but the concept should hold
    expect(() => {
      db.off('cells')
    }).not.toThrow()
  })

  it('should handle rapid subscribe/unsubscribe cycles', () => {
    expect(() => {
      for (let i = 0; i < 1000; i++) {
        const callback = vi.fn()
        const unsubscribe = db.onChildAdded('cells', callback)
        unsubscribe()
      }
    }).not.toThrow()
  })

  it('should handle many simultaneous listeners efficiently', async () => {
    const NUM_LISTENERS = 100
    const callbacks = Array.from({ length: NUM_LISTENERS }, () => vi.fn())

    callbacks.forEach((cb) => db.onChildAdded('cells', cb))

    await db.insert('cells', 'cell1', { value: 1 })

    callbacks.forEach((cb) => {
      expect(cb).toHaveBeenCalledTimes(1)
    })
  })
})

// ============================================================================
// Edge Cases
// ============================================================================

describe('Edge Cases', () => {
  let db: DatabaseDO

  beforeEach(() => {
    db = new StubDatabaseDO()
  })

  it('should handle empty collection names', async () => {
    const callback = vi.fn()

    expect(() => {
      db.onChildAdded('', callback)
    }).not.toThrow()
  })

  it('should handle special characters in collection names', async () => {
    const callback = vi.fn()

    expect(() => {
      db.onChildAdded('my-collection.subcollection', callback)
    }).not.toThrow()
  })

  it('should handle null/undefined data gracefully', async () => {
    const callback = vi.fn()
    db.onChildAdded('cells', callback)

    await db.insert('cells', 'cell1', null as unknown)

    expect(callback).toHaveBeenCalled()
  })

  it('should handle very large documents', async () => {
    const callback = vi.fn()
    db.onChildAdded('cells', callback)

    const largeData = {
      array: Array.from({ length: 10000 }, (_, i) => ({ index: i, value: `item-${i}` })),
    }

    await db.insert('cells', 'cell1', largeData)

    expect(callback).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({
          array: expect.arrayContaining([{ index: 0, value: 'item-0' }]),
        }),
      })
    )
  })

  it('should handle deeply nested updates', async () => {
    const callback = vi.fn()

    await db.insert('cells', 'cell1', {
      level1: {
        level2: {
          level3: {
            value: 'original',
          },
        },
      },
    })

    db.onChildChanged('cells', callback)

    await db.update('cells', 'cell1', {
      level1: {
        level2: {
          level3: {
            value: 'updated',
          },
        },
      },
    })

    expect(callback).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({
          level1: expect.objectContaining({
            level2: expect.objectContaining({
              level3: expect.objectContaining({
                value: 'updated',
              }),
            }),
          }),
        }),
      })
    )
  })
})
