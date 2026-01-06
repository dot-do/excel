/**
 * MongoDB Storage Layer for excel.do
 *
 * Provides storage operations for cells, sheets, and workbooks using mongo.do.
 */

import type { Cell, CellValue, CellPrimitive } from '../types'
import type {
  MongoClient,
  MongoDatabase,
  MongoCollection,
  Sheet,
  Workbook,
  CellRef,
  CellData,
  CellQuery,
  Range,
  CreateSheetOptions,
  UpdateSheetOptions,
  CreateWorkbookOptions,
  ICellStore,
  ISheetStore,
  IWorkbookStore,
  IStorageManager,
  TransactionFn,
  IndexSpec,
  IndexOptions,
} from './types'
import { CELL_INDEXES, SHEET_INDEXES, WORKBOOK_INDEXES } from './types'

// ============================================================================
// Error Classes
// ============================================================================

/** Error thrown by stub implementations */
export class NotImplementedError extends Error {
  constructor(method: string) {
    super(`Not implemented: ${method}`)
    this.name = 'NotImplementedError'
  }
}

/** Error thrown when a resource is not found */
export class NotFoundError extends Error {
  constructor(resource: string, id: string) {
    super(`${resource} not found: ${id}`)
    this.name = 'NotFoundError'
  }
}

/** Error thrown during transaction failures */
export class TransactionError extends Error {
  constructor(message: string, public readonly cause?: Error) {
    super(message)
    this.name = 'TransactionError'
  }
}

/** Conflict error thrown when optimistic lock fails */
export class ConflictError extends Error {
  constructor(
    message: string,
    public readonly collection: string,
    public readonly id: string,
    public readonly expectedVersion: number,
    public readonly actualVersion: number,
    public readonly operation?: 'write' | 'delete'
  ) {
    super(message)
    this.name = 'ConflictError'
  }
}

// ============================================================================
// Utility Functions
// ============================================================================

/**
 * Parse a cell reference string (e.g., "A1") into components
 */
export function parseCellRef(ref: CellRef): { col: string; row: number } {
  const normalized = ref.toUpperCase()
  const match = normalized.match(/^([A-Z]+)(\d+)$/)
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`)
  }
  return {
    col: match[1],
    row: parseInt(match[2], 10),
  }
}

/**
 * Parse a range string (e.g., "A1:B10") into Range object
 */
export function parseRange(range: string): Range {
  const parts = range.split(':')
  if (parts.length !== 2) {
    throw new Error(`Invalid range: ${range}`)
  }

  const start = parseCellRef(parts[0])
  const end = parseCellRef(parts[1])

  return {
    startCol: start.col,
    startRow: start.row,
    endCol: end.col,
    endRow: end.row,
  }
}

/**
 * Generate cell ID for storage
 */
export function generateCellId(sheetId: string, cellRef: CellRef): string {
  // Check if sheet name needs quoting (contains spaces or special chars)
  if (sheetId.includes(' ') || sheetId.includes("'")) {
    // Escape single quotes by doubling them
    const escaped = sheetId.replace(/'/g, "''")
    return `'${escaped}'!${cellRef.toUpperCase()}`
  }
  return `${sheetId}!${cellRef.toUpperCase()}`
}

/**
 * Generate a unique ID
 */
export function generateId(): string {
  // Generate a unique ID using timestamp + random
  const timestamp = Date.now().toString(36)
  const random = Math.random().toString(36).substring(2, 10)
  return `${timestamp}-${random}`
}

/**
 * Convert column letter(s) to 0-based index
 */
function colToIndex(col: string): number {
  let index = 0
  const upper = col.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64)
  }
  return index - 1
}

/**
 * Convert 0-based index to column letter(s)
 * @internal Reserved for future use
 */
function _indexToCol(index: number): string {
  let col = ''
  let n = index + 1
  while (n > 0) {
    const remainder = (n - 1) % 26
    col = String.fromCharCode(65 + remainder) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}
void _indexToCol

/**
 * Determine cell value type from primitive
 */
function getCellValueType(value: CellPrimitive): CellValue['t'] {
  if (value === null || value === undefined) return 'empty'
  if (typeof value === 'string') return 'string'
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  if (value instanceof Date) return 'date'
  return 'empty'
}

/**
 * Create a Cell object from primitive or CellData
 */
function createCellObject(
  sheetId: string,
  cellRef: CellRef,
  input: CellPrimitive | CellData
): Cell {
  const { col, row } = parseCellRef(cellRef)
  const colIndex = colToIndex(col)
  const now = new Date()

  // Handle CellData object
  if (typeof input === 'object' && input !== null && !(input instanceof Date) && 'ref' in input) {
    const cellData = input as CellData
    const value = cellData.value ?? null
    const cellValue: CellValue = {
      v: value,
      t: getCellValueType(value),
    }
    if (cellData.formula) {
      cellValue.f = cellData.formula
    }

    return {
      _id: generateCellId(sheetId, cellRef),
      sheet: sheetId,
      row,
      col,
      colIndex,
      value: cellValue,
      format: cellData.format,
      metadata: cellData.metadata,
      createdAt: now,
      updatedAt: now,
    }
  }

  // Handle primitive value
  const primitive = input as CellPrimitive
  const cellValue: CellValue = {
    v: primitive,
    t: getCellValueType(primitive),
  }

  return {
    _id: generateCellId(sheetId, cellRef),
    sheet: sheetId,
    row,
    col,
    colIndex,
    value: cellValue,
    createdAt: now,
    updatedAt: now,
  }
}

// ============================================================================
// BoundedCache - LRU-style Cache with Max Size
// ============================================================================

/**
 * Options for configuring the CellStore
 */
export interface CellStoreOptions {
  /** Collection name for cells (default: 'cells') */
  collectionName?: string
  /** Maximum number of entries in the cell cache (default: 10000) */
  maxCacheSize?: number
}

/**
 * Simple bounded cache using Map's insertion order for LRU-style eviction.
 * When the cache exceeds maxSize, the oldest entries are evicted.
 */
export class BoundedCache<K, V> extends Map<K, V> {
  constructor(private maxSize: number = 10000) {
    super()
  }

  /**
   * Set a value in the cache. If the key already exists, it's moved to the end (most recent).
   * If the cache exceeds maxSize, the oldest entry is evicted.
   */
  set(key: K, value: V): this {
    // Delete and re-add to move to end (most recent)
    if (this.has(key)) {
      this.delete(key)
    }
    super.set(key, value)

    // Evict oldest if over limit
    if (this.size > this.maxSize) {
      const firstKey = this.keys().next().value
      if (firstKey !== undefined) this.delete(firstKey)
    }
    return this
  }

  /**
   * Get a value from the cache. If found, the entry is moved to the end (most recently used).
   */
  get(key: K): V | undefined {
    const value = super.get(key)
    if (value !== undefined) {
      // Move to end (most recently used)
      this.delete(key)
      super.set(key, value)
    }
    return value
  }

  /**
   * Get the maximum size of this cache
   */
  getMaxSize(): number {
    return this.maxSize
  }
}

// ============================================================================
// CellStore Implementation
// ============================================================================

/** Callback type for version tracking */
export type VersionUpdateCallback = (sheetId: string, cellRef: string) => void

/** Cell event types for real-time listeners */
export type CellEventType = 'added' | 'changed' | 'removed'

/** Cell event data */
export interface CellEvent {
  type: CellEventType
  sheetId: string
  cellRef: string
  cell: Cell | null
  previousCell: Cell | null
}

/** Cell event callback */
export type CellEventCallback = (event: CellEvent) => void

/** Default maximum cache size */
const DEFAULT_MAX_CACHE_SIZE = 10000

export class CellStore implements ICellStore {
  private collection: MongoCollection<Cell>
  private onVersionUpdate?: VersionUpdateCallback
  private eventListeners: CellEventCallback[] = []
  private cellCache: BoundedCache<string, Cell>

  constructor(db: MongoDatabase, options?: CellStoreOptions | string) {
    // Support both old signature (collectionName string) and new signature (options object)
    const opts: CellStoreOptions = typeof options === 'string'
      ? { collectionName: options }
      : options || {}

    const collectionName = opts.collectionName || 'cells'
    const maxCacheSize = opts.maxCacheSize ?? DEFAULT_MAX_CACHE_SIZE

    this.collection = db.collection<Cell>(collectionName)
    this.cellCache = new BoundedCache<string, Cell>(maxCacheSize)
  }

  /**
   * Get the cell cache (for testing purposes)
   */
  getCellCache(): BoundedCache<string, Cell> {
    return this.cellCache
  }

  /**
   * Set the version update callback (called by StorageManager)
   */
  setVersionUpdateCallback(callback: VersionUpdateCallback): void {
    this.onVersionUpdate = callback
  }

  /**
   * Add an event listener for cell changes
   */
  addCellEventListener(callback: CellEventCallback): () => void {
    this.eventListeners.push(callback)
    return () => {
      const idx = this.eventListeners.indexOf(callback)
      if (idx >= 0) this.eventListeners.splice(idx, 1)
    }
  }

  /**
   * Emit a cell event to all listeners
   */
  private emitCellEvent(event: CellEvent): void {
    for (const listener of this.eventListeners) {
      try {
        listener(event)
      } catch {
        // Continue emitting to other listeners
      }
    }
  }

  /**
   * Get cache key for a cell
   */
  private getCacheKey(sheetId: string, cellRef: string): string {
    return `${sheetId}!${cellRef.toUpperCase()}`
  }

  async getCell(sheetId: string, cellRef: CellRef): Promise<Cell | null> {
    const { col, row } = parseCellRef(cellRef)
    return this.collection.findOne({
      sheet: sheetId,
      col,
      row,
    })
  }

  async setCell(sheetId: string, cellRef: CellRef, value: CellPrimitive | CellData): Promise<Cell> {
    const cell = createCellObject(sheetId, cellRef, value)
    const cacheKey = this.getCacheKey(sheetId, cellRef)

    // Check if cell exists
    const existing = await this.getCell(sheetId, cellRef)
    const previousCell = existing ? { ...existing } : null

    if (existing) {
      // Update existing cell
      cell._id = existing._id
      cell.createdAt = existing.createdAt
      cell.updatedAt = new Date()

      await this.collection.updateOne(
        { _id: existing._id },
        { $set: cell }
      )

      // Update cache and emit changed event
      this.cellCache.set(cacheKey, cell)
      this.emitCellEvent({
        type: 'changed',
        sheetId,
        cellRef,
        cell,
        previousCell,
      })
    } else {
      // Insert new cell
      await this.collection.insertOne(cell)

      // Update cache and emit added event
      this.cellCache.set(cacheKey, cell)
      this.emitCellEvent({
        type: 'added',
        sheetId,
        cellRef,
        cell,
        previousCell: null,
      })
    }

    // Notify about version update
    if (this.onVersionUpdate) {
      this.onVersionUpdate(sheetId, cellRef)
    }

    return cell
  }

  async getCells(sheetId: string, range: Range | string): Promise<Cell[]> {
    const r = typeof range === 'string' ? parseRange(range) : range

    const startColIndex = colToIndex(r.startCol)
    const endColIndex = colToIndex(r.endCol)

    return this.collection.find({
      sheet: sheetId,
      colIndex: { $gte: startColIndex, $lte: endColIndex },
      row: { $gte: r.startRow, $lte: r.endRow },
    }, {
      sort: { row: 1, colIndex: 1 },
    })
  }

  async setCells(sheetId: string, cells: CellData[]): Promise<Cell[]> {
    if (cells.length === 0) return []

    const results: Cell[] = []
    for (const cellData of cells) {
      const cell = await this.setCell(sheetId, cellData.ref, cellData)
      results.push(cell)
    }
    return results
  }

  async deleteCell(sheetId: string, cellRef: CellRef): Promise<boolean> {
    // First find the cell to get its _id
    const cell = await this.getCell(sheetId, cellRef)
    if (!cell) return false

    const result = await this.collection.deleteOne({ _id: cell._id })

    if (result.deletedCount > 0) {
      const cacheKey = this.getCacheKey(sheetId, cellRef)
      this.cellCache.delete(cacheKey)

      // Emit removed event
      this.emitCellEvent({
        type: 'removed',
        sheetId,
        cellRef,
        cell: null,
        previousCell: cell,
      })
    }

    return result.deletedCount > 0
  }

  async deleteCells(sheetId: string, range: Range | string): Promise<number> {
    // First get all cells in range
    const cells = await this.getCells(sheetId, range)
    if (cells.length === 0) return 0

    // Delete each cell by _id
    let count = 0
    for (const cell of cells) {
      const result = await this.collection.deleteOne({ _id: cell._id })
      count += result.deletedCount
    }

    return count
  }

  async findCells(sheetId: string, query: CellQuery): Promise<Cell[]> {
    // Get all cells for the sheet first (mock doesn't support complex queries)
    const allCells = await this.collection.find({ sheet: sheetId })

    // Apply filters in-memory
    let results = allCells.filter((cell) => {
      // Filter by value
      if (query.value !== undefined && cell.value.v !== query.value) {
        return false
      }

      // Filter by value type
      if (query.valueType && cell.value.t !== query.valueType) {
        return false
      }

      // Filter by formula presence
      if (query.hasFormula !== undefined) {
        const hasFormula = cell.value.f !== undefined && cell.value.f !== null
        if (query.hasFormula !== hasFormula) {
          return false
        }
      }

      // Filter by formula pattern
      if (query.formulaPattern) {
        if (!cell.value.f || !cell.value.f.includes(query.formulaPattern)) {
          return false
        }
      }

      // Filter by row range
      if (query.rowRange) {
        if (query.rowRange.min !== undefined && cell.row < query.rowRange.min) {
          return false
        }
        if (query.rowRange.max !== undefined && cell.row > query.rowRange.max) {
          return false
        }
      }

      // Filter by column range
      if (query.colRange) {
        if (query.colRange.minIndex !== undefined && cell.colIndex < query.colRange.minIndex) {
          return false
        }
        if (query.colRange.maxIndex !== undefined && cell.colIndex > query.colRange.maxIndex) {
          return false
        }
      }

      // When hasFormula is not explicitly set, but we're filtering by valueType,
      // exclude cells with formulas since they have computed values that may differ from raw type
      if (query.valueType && query.hasFormula === undefined) {
        const hasFormula = cell.value.f !== undefined && cell.value.f !== null
        if (hasFormula) {
          return false
        }
      }

      return true
    })

    // Sort by row then column
    results.sort((a, b) => {
      if (a.row !== b.row) return a.row - b.row
      return a.colIndex - b.colIndex
    })

    return results
  }

  async countCells(sheetId: string, query?: CellQuery): Promise<number> {
    if (!query) {
      return this.collection.countDocuments({ sheet: sheetId })
    }

    const cells = await this.findCells(sheetId, query)
    return cells.length
  }

  /**
   * Create a query builder for Firebase-style queries
   */
  query(sheetId: string): CellQuery$ {
    return new CellQuery$(this, sheetId)
  }
}

// ============================================================================
// CellQuery$ - Firebase-style Query Builder
// ============================================================================

/**
 * Get a nested property from an object using dot notation
 */
function getNestedValue(obj: unknown, path: string): unknown {
  const parts = path.split('.')
  let current: unknown = obj
  for (const part of parts) {
    if (current === null || current === undefined || typeof current !== 'object') {
      return undefined
    }
    current = (current as Record<string, unknown>)[part]
  }
  return current
}

/**
 * Firebase-style query builder for cells
 */
export class CellQuery$ {
  private _orderByField?: string
  private _orderByType?: 'child' | 'key' | 'value'
  private _limitFirst?: number
  private _limitLast?: number
  private _startAtValue?: unknown
  private _startAtKey?: string
  private _endAtValue?: unknown
  private _endAtKey?: string
  private _equalToValue?: unknown
  private _equalToKey?: string

  // Track current result set for real-time updates
  private currentResults: Cell[] = []
  private resultMap: Map<string, Cell> = new Map()

  constructor(
    private readonly store: CellStore,
    private readonly sheetId: string
  ) {
    // Suppress unused variable warnings for key params (reserved for future key-based filtering)
    void this._startAtKey
    void this._endAtKey
    void this._equalToKey
  }

  /**
   * Order by a child property
   */
  orderByChild(path: string): CellQuery$ {
    if (this._orderByType) {
      throw new Error('Cannot call orderByChild when already ordered')
    }
    this._orderByType = 'child'
    this._orderByField = path
    return this
  }

  /**
   * Order by key (cell reference)
   */
  orderByKey(): CellQuery$ {
    if (this._orderByType) {
      throw new Error('Cannot call orderByKey when already ordered')
    }
    this._orderByType = 'key'
    return this
  }

  /**
   * Order by value
   */
  orderByValue(): CellQuery$ {
    if (this._orderByType) {
      throw new Error('Cannot call orderByValue when already ordered')
    }
    this._orderByType = 'value'
    this._orderByField = 'value.v'
    return this
  }

  /**
   * Limit to first N results
   */
  limitToFirst(limit: number): CellQuery$ {
    if (this._limitLast !== undefined) {
      throw new Error('Cannot call limitToFirst when limitToLast is already set')
    }
    this._limitFirst = limit
    return this
  }

  /**
   * Limit to last N results
   */
  limitToLast(limit: number): CellQuery$ {
    if (this._limitFirst !== undefined) {
      throw new Error('Cannot call limitToLast when limitToFirst is already set')
    }
    this._limitLast = limit
    return this
  }

  /**
   * Start at a value (inclusive)
   */
  startAt(value: unknown, key?: string): CellQuery$ {
    this._startAtValue = value
    this._startAtKey = key
    return this
  }

  /**
   * End at a value (inclusive)
   */
  endAt(value: unknown, key?: string): CellQuery$ {
    this._endAtValue = value
    this._endAtKey = key
    return this
  }

  /**
   * Filter to exact value
   */
  equalTo(value: unknown, key?: string): CellQuery$ {
    this._equalToValue = value
    this._equalToKey = key
    return this
  }

  /**
   * Execute the query and return results
   */
  async get(): Promise<Cell[]> {
    // Get all cells for the sheet
    let cells = await this.store.findCells(this.sheetId, {})

    // Apply ordering
    if (this._orderByType === 'child' && this._orderByField) {
      cells = this.orderCells(cells, this._orderByField)
    } else if (this._orderByType === 'key') {
      cells.sort((a, b) => a._id.localeCompare(b._id))
    } else if (this._orderByType === 'value') {
      cells = this.orderCells(cells, 'value.v')
    } else {
      // Default ordering by row then column
      cells.sort((a, b) => {
        if (a.row !== b.row) return a.row - b.row
        return a.colIndex - b.colIndex
      })
    }

    // Apply equalTo filter
    if (this._equalToValue !== undefined) {
      const field = this._orderByField || 'value.v'
      cells = cells.filter((c) => getNestedValue(c, field) === this._equalToValue)
    }

    // Apply startAt filter
    if (this._startAtValue !== undefined) {
      const field = this._orderByField || 'value.v'
      cells = cells.filter((c) => {
        const val = getNestedValue(c, field)
        if (val === undefined || val === null) return false
        return val >= this._startAtValue!
      })
    }

    // Apply endAt filter
    if (this._endAtValue !== undefined) {
      const field = this._orderByField || 'value.v'
      cells = cells.filter((c) => {
        const val = getNestedValue(c, field)
        if (val === undefined || val === null) return false
        return val <= this._endAtValue!
      })
    }

    // Apply limits
    if (this._limitFirst !== undefined) {
      if (this._limitFirst === 0) {
        cells = []
      } else {
        cells = cells.slice(0, this._limitFirst)
      }
    } else if (this._limitLast !== undefined) {
      if (this._limitLast === 0) {
        cells = []
      } else {
        cells = cells.slice(-this._limitLast)
      }
    }

    return cells
  }

  private orderCells(cells: Cell[], field: string): Cell[] {
    return [...cells].sort((a, b) => {
      const aVal = getNestedValue(a, field)
      const bVal = getNestedValue(b, field)

      // Handle undefined/null values - put them at the end
      if ((aVal === undefined || aVal === null) && (bVal === undefined || bVal === null)) return 0
      if (aVal === undefined || aVal === null) return 1
      if (bVal === undefined || bVal === null) return -1

      // String comparison
      if (typeof aVal === 'string' && typeof bVal === 'string') {
        return aVal.localeCompare(bVal)
      }

      // Numeric comparison
      if (aVal < bVal) return -1
      if (aVal > bVal) return 1
      return 0
    })
  }

  /**
   * Check if a cell matches the current query criteria
   */
  private matchesQuery(cell: Cell): boolean {
    // Check sheet match
    if (cell.sheet !== this.sheetId) return false

    // Check equalTo filter
    if (this._equalToValue !== undefined) {
      const field = this._orderByField || 'value.v'
      const val = getNestedValue(cell, field)
      if (val !== this._equalToValue) return false
    }

    // Check startAt filter
    const startAt = this._startAtValue
    if (startAt !== undefined && startAt !== null) {
      const field = this._orderByField || 'value.v'
      const val = getNestedValue(cell, field)
      if (val === undefined || val === null || val < startAt) return false
    }

    // Check endAt filter
    const endAt = this._endAtValue
    if (endAt !== undefined && endAt !== null) {
      const field = this._orderByField || 'value.v'
      const val = getNestedValue(cell, field)
      if (val === undefined || val === null || val > endAt) return false
    }

    return true
  }

  /**
   * Get the position of a cell in the current results
   */
  private getPosition(cellId: string): number {
    return this.currentResults.findIndex(c => c._id === cellId)
  }

  /**
   * Get the previous key at a given position
   */
  private getPreviousKey(position: number): string | null {
    if (position <= 0) return null
    return this.currentResults[position - 1]?._id || null
  }

  /**
   * Subscribe to value changes
   */
  onValue(callback: (cells: Cell[]) => void): () => void {
    // Execute query and call callback
    this.get().then((cells) => {
      this.currentResults = cells
      this.resultMap.clear()
      for (const cell of cells) {
        this.resultMap.set(cell._id, cell)
      }
      callback(cells)
    })

    // Subscribe to cell events
    const unsubscribe = this.store.addCellEventListener((event) => {
      if (event.sheetId !== this.sheetId) return

      // Re-run query and callback with new results
      this.get().then((cells) => {
        this.currentResults = cells
        this.resultMap.clear()
        for (const cell of cells) {
          this.resultMap.set(cell._id, cell)
        }
        callback(cells)
      })
    })

    return unsubscribe
  }

  /**
   * Subscribe to child added events
   */
  onChildAdded(callback: (cell: Cell, previousKey: string | null) => void): () => void {
    // Initial call with existing children
    this.get().then((cells) => {
      this.currentResults = cells
      this.resultMap.clear()
      let prevKey: string | null = null
      for (const cell of cells) {
        this.resultMap.set(cell._id, cell)
        callback(cell, prevKey)
        prevKey = cell._id
      }
    })

    // Subscribe to cell events
    const unsubscribe = this.store.addCellEventListener((event) => {
      if (event.sheetId !== this.sheetId) return

      if (event.type === 'added' && event.cell && this.matchesQuery(event.cell)) {
        // Re-run query to get correct order
        this.get().then((cells) => {
          const newCell = event.cell!
          const previousResults = new Set(this.resultMap.keys())

          this.currentResults = cells
          this.resultMap.clear()
          for (const cell of cells) {
            this.resultMap.set(cell._id, cell)
          }

          // Only trigger callback if this is truly a new cell in results
          if (!previousResults.has(newCell._id)) {
            const position = this.getPosition(newCell._id)
            const prevKey = this.getPreviousKey(position)
            callback(newCell, prevKey)
          }
        })
      }
    })

    return unsubscribe
  }

  /**
   * Subscribe to child changed events
   */
  onChildChanged(callback: (cell: Cell, previousKey: string | null) => void): () => void {
    // Initialize result set first
    this.get().then((cells) => {
      this.currentResults = cells
      this.resultMap.clear()
      for (const cell of cells) {
        this.resultMap.set(cell._id, cell)
      }
    })

    // Subscribe to cell events
    const unsubscribe = this.store.addCellEventListener((event) => {
      if (event.sheetId !== this.sheetId) return

      if (event.type === 'changed' && event.cell) {
        // Check if this cell was already in our results
        const wasInResults = this.resultMap.has(event.cell._id)
        const nowMatchesQuery = this.matchesQuery(event.cell)

        if (wasInResults && nowMatchesQuery) {
          // Re-run query to get correct order
          this.get().then((cells) => {
            this.currentResults = cells
            this.resultMap.clear()
            for (const cell of cells) {
              this.resultMap.set(cell._id, cell)
            }

            const updatedCell = event.cell!
            const position = this.getPosition(updatedCell._id)
            const prevKey = this.getPreviousKey(position)
            callback(updatedCell, prevKey)
          })
        }
      }
    })

    return unsubscribe
  }

  /**
   * Subscribe to child removed events
   */
  onChildRemoved(callback: (cell: Cell) => void): () => void {
    // Initialize result set first
    this.get().then((cells) => {
      this.currentResults = cells
      this.resultMap.clear()
      for (const cell of cells) {
        this.resultMap.set(cell._id, cell)
      }
    })

    // Subscribe to cell events
    const unsubscribe = this.store.addCellEventListener((event) => {
      if (event.sheetId !== this.sheetId) return

      if (event.type === 'removed' && event.previousCell) {
        // Check if this cell was in our results
        const wasInResults = this.resultMap.has(event.previousCell._id)

        if (wasInResults) {
          const removedCell = event.previousCell
          this.resultMap.delete(removedCell._id)
          this.currentResults = this.currentResults.filter(c => c._id !== removedCell._id)
          callback(removedCell)
        }
      }
    })

    return unsubscribe
  }

  /**
   * Subscribe to child moved events
   */
  onChildMoved(callback: (cell: Cell, previousKey: string | null) => void): () => void {
    // Initialize result set first
    this.get().then((cells) => {
      this.currentResults = cells
      this.resultMap.clear()
      for (const cell of cells) {
        this.resultMap.set(cell._id, cell)
      }
    })

    // Subscribe to cell events
    const unsubscribe = this.store.addCellEventListener((event) => {
      if (event.sheetId !== this.sheetId) return

      if (event.type === 'changed' && event.cell && event.previousCell) {
        // Check if this cell was already in our results
        const wasInResults = this.resultMap.has(event.cell._id)
        const nowMatchesQuery = this.matchesQuery(event.cell)

        if (wasInResults && nowMatchesQuery) {
          // Get position before update
          const oldPosition = this.getPosition(event.cell._id)

          // Re-run query to get correct order
          this.get().then((cells) => {
            const updatedCell = event.cell!
            this.currentResults = cells
            this.resultMap.clear()
            for (const cell of cells) {
              this.resultMap.set(cell._id, cell)
            }

            // Get position after update
            const newPosition = this.getPosition(updatedCell._id)

            // Only trigger callback if position changed
            if (oldPosition !== newPosition) {
              const prevKey = this.getPreviousKey(newPosition)
              callback(updatedCell, prevKey)
            }
          })
        }
      }
    })

    return unsubscribe
  }
}

// ============================================================================
// SheetStore Implementation
// ============================================================================

export class SheetStore implements ISheetStore {
  private collection: MongoCollection<Sheet>

  constructor(db: MongoDatabase, collectionName: string = 'sheets') {
    this.collection = db.collection<Sheet>(collectionName)
  }

  async createSheet(
    workbookId: string,
    name: string,
    options?: Omit<CreateSheetOptions, 'workbookId' | 'name'>
  ): Promise<Sheet> {
    const now = new Date()

    // Get next index if not specified
    let index = options?.index
    if (index === undefined) {
      const sheets = await this.listSheets(workbookId)
      index = sheets.length
    }

    const sheet: Sheet = {
      _id: generateId(),
      workbookId,
      name,
      index,
      defaultColWidth: options?.defaultColWidth,
      defaultRowHeight: options?.defaultRowHeight,
      createdAt: now,
      updatedAt: now,
    }

    await this.collection.insertOne(sheet)
    return sheet
  }

  async getSheet(id: string): Promise<Sheet | null> {
    return this.collection.findOne({ _id: id })
  }

  async getSheetByName(workbookId: string, name: string): Promise<Sheet | null> {
    return this.collection.findOne({ workbookId, name })
  }

  async updateSheet(id: string, updates: UpdateSheetOptions): Promise<Sheet | null> {
    const sheet = await this.getSheet(id)
    if (!sheet) return null

    const updatedSheet: Sheet = {
      ...sheet,
      ...updates,
      updatedAt: new Date(),
    }

    await this.collection.updateOne(
      { _id: id },
      { $set: updatedSheet }
    )

    return updatedSheet
  }

  async deleteSheet(id: string): Promise<boolean> {
    const result = await this.collection.deleteOne({ _id: id })
    return result.deletedCount > 0
  }

  async listSheets(workbookId: string): Promise<Sheet[]> {
    const sheets = await this.collection.find({ workbookId })
    // Sort by index in-memory (mock doesn't support sort)
    return sheets.sort((a, b) => a.index - b.index)
  }

  async reorderSheets(_workbookId: string, sheetIds: string[]): Promise<void> {
    // Note: workbookId parameter reserved for future validation
    for (let i = 0; i < sheetIds.length; i++) {
      // Use simple _id filter (mock doesn't handle compound filters well)
      await this.collection.updateOne(
        { _id: sheetIds[i] },
        { $set: { index: i, updatedAt: new Date() } }
      )
    }
  }
}

// ============================================================================
// WorkbookStore Implementation
// ============================================================================

export class WorkbookStore implements IWorkbookStore {
  private collection: MongoCollection<Workbook>

  constructor(db: MongoDatabase, collectionName: string = 'workbooks') {
    this.collection = db.collection<Workbook>(collectionName)
  }

  async createWorkbook(
    name: string,
    options?: Omit<CreateWorkbookOptions, 'name'>
  ): Promise<Workbook> {
    const now = new Date()

    const workbook: Workbook = {
      _id: generateId(),
      name,
      activeSheet: 0,
      properties: options?.properties,
      createdAt: now,
      updatedAt: now,
    }

    await this.collection.insertOne(workbook)
    return workbook
  }

  async getWorkbook(id: string): Promise<Workbook | null> {
    return this.collection.findOne({ _id: id })
  }

  async updateWorkbook(
    id: string,
    updates: Partial<Omit<Workbook, '_id' | 'createdAt'>>
  ): Promise<Workbook | null> {
    const workbook = await this.getWorkbook(id)
    if (!workbook) return null

    const updatedWorkbook: Workbook = {
      ...workbook,
      ...updates,
      updatedAt: new Date(),
    }

    await this.collection.updateOne(
      { _id: id },
      { $set: updatedWorkbook }
    )

    return updatedWorkbook
  }

  async deleteWorkbook(id: string): Promise<boolean> {
    const result = await this.collection.deleteOne({ _id: id })
    return result.deletedCount > 0
  }

  async listWorkbooks(): Promise<Workbook[]> {
    const workbooks = await this.collection.find({})
    // Sort by updatedAt descending in-memory (mock doesn't support sort)
    return workbooks.sort((a, b) => b.updatedAt.getTime() - a.updatedAt.getTime())
  }
}

// ============================================================================
// RTDB Transaction Types
// ============================================================================

/**
 * Transaction options for RTDB protocol
 */
export interface RTDBTransactionOptions {
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
export interface RTDBTransactionContext {
  /** Read a value with version tracking */
  read<T>(collection: string, id: string): Promise<{ value: T | null; version: number }>
  /** Write a value, checking version for conflicts */
  write<T>(collection: string, id: string, value: T, expectedVersion: number): Promise<void>
  /** Delete a value, checking version for conflicts */
  delete(collection: string, id: string, expectedVersion: number): Promise<void>
  /** Get current transaction timestamp */
  timestamp(): number
}

/**
 * Atomic operation for batch processing
 */
export interface AtomicOperation {
  type: 'set' | 'increment' | 'delete'
  sheetId: string
  cellRef: string
  value?: CellPrimitive | CellData
  delta?: number
  expectedVersion?: number
}

// ============================================================================
// StorageManager Implementation
// ============================================================================

export class StorageManager implements IStorageManager {
  public readonly cells: CellStore
  public readonly sheets: SheetStore
  public readonly workbooks: WorkbookStore
  private readonly db: MongoDatabase
  // Version tracking map for RTDB
  private readonly versions: Map<string, number> = new Map()
  // Lock map for atomic operations
  private readonly locks: Map<string, Promise<void>> = new Map()

  constructor(private readonly client: MongoClient, dbName: string = 'excel') {
    this.db = client.db(dbName)
    this.cells = new CellStore(this.db)
    this.sheets = new SheetStore(this.db)
    this.workbooks = new WorkbookStore(this.db)

    // Set up version tracking callback
    this.cells.setVersionUpdateCallback((sheetId, cellRef) => {
      const key = this.getCellVersionKey(sheetId, cellRef)
      this.incrementVersion(key)
    })
  }

  async withTransaction<T>(fn: TransactionFn<T>): Promise<T> {
    const session = await this.client.startSession()
    try {
      session.startTransaction()
      const result = await fn(session)
      await session.commitTransaction()
      return result
    } catch (error) {
      await session.abortTransaction()
      throw error
    } finally {
      await session.endSession()
    }
  }

  async createIndexes(): Promise<void> {
    // Create cell indexes
    const cellsCollection = this.db.collection<Cell>('cells')
    for (const { spec, options } of CELL_INDEXES) {
      await cellsCollection.createIndex(spec, options)
    }

    // Create sheet indexes
    const sheetsCollection = this.db.collection<Sheet>('sheets')
    for (const { spec, options } of SHEET_INDEXES) {
      await sheetsCollection.createIndex(spec, options)
    }

    // Create workbook indexes
    const workbooksCollection = this.db.collection<Workbook>('workbooks')
    for (const { spec, options } of WORKBOOK_INDEXES) {
      await workbooksCollection.createIndex(spec, options)
    }
  }

  async close(): Promise<void> {
    await this.client.close()
  }

  // ============================================================================
  // RTDB Transaction Methods
  // ============================================================================

  /**
   * Get the version key for a cell
   */
  private getCellVersionKey(sheetId: string, cellRef: string): string {
    return `${sheetId}!${cellRef.toUpperCase()}`
  }

  /**
   * Get the current version for a cell
   */
  private getVersion(key: string): number {
    return this.versions.get(key) || 0
  }

  /**
   * Set the version for a cell
   */
  private setVersion(key: string, version: number): void {
    this.versions.set(key, version)
  }

  /**
   * Increment the version for a cell
   */
  private incrementVersion(key: string): number {
    const current = this.getVersion(key)
    const newVersion = current + 1
    this.setVersion(key, newVersion)
    return newVersion
  }

  /**
   * Acquire a lock for a key (for atomic operations)
   */
  private async acquireLock(key: string): Promise<() => void> {
    // Wait for any existing lock
    while (this.locks.has(key)) {
      await this.locks.get(key)
    }

    // Create a new lock
    let releaseFn: () => void = () => {}
    const lockPromise = new Promise<void>((resolve) => {
      releaseFn = () => {
        this.locks.delete(key)
        resolve()
      }
    })
    this.locks.set(key, lockPromise)

    return releaseFn
  }

  /**
   * Read a cell with version information for transactions
   */
  async readCellWithVersion(
    sheetId: string,
    cellRef: string
  ): Promise<{ cell: Cell | null; version: number }> {
    const cell = await this.cells.getCell(sheetId, cellRef)
    const key = this.getCellVersionKey(sheetId, cellRef)
    const version = this.getVersion(key)
    return { cell, version }
  }

  /**
   * Write a cell with version check (compare-and-swap)
   */
  async writeCellWithVersion(
    sheetId: string,
    cellRef: string,
    value: CellPrimitive | CellData,
    expectedVersion: number
  ): Promise<Cell> {
    const key = this.getCellVersionKey(sheetId, cellRef)

    // Acquire lock for atomic compare-and-swap
    const release = await this.acquireLock(key)
    try {
      const currentVersion = this.getVersion(key)

      // Check for version conflict
      if (expectedVersion !== currentVersion) {
        throw new ConflictError(
          `Version conflict: expected ${expectedVersion}, got ${currentVersion}`,
          'cells',
          key,
          expectedVersion,
          currentVersion
        )
      }

      // Perform the write (callback will increment version)
      const cell = await this.cells.setCell(sheetId, cellRef, value)

      return cell
    } finally {
      release()
    }
  }

  /**
   * Atomic increment operation
   */
  async incrementCell(sheetId: string, cellRef: string, delta: number): Promise<Cell> {
    // Use internal locking for atomic increment
    const maxRetries = 10
    let attempts = 0

    while (attempts < maxRetries) {
      const { cell, version } = await this.readCellWithVersion(sheetId, cellRef)

      if (!cell || cell.value.t !== 'number' || typeof cell.value.v !== 'number') {
        throw new Error('Cannot increment non-numeric cell')
      }

      const newValue = cell.value.v + delta

      try {
        return await this.writeCellWithVersion(sheetId, cellRef, newValue, version)
      } catch (error) {
        if (error instanceof ConflictError) {
          attempts++
          continue
        }
        throw error
      }
    }

    throw new TransactionError('Failed to increment after max retries')
  }

  /**
   * Batch atomic operations
   */
  async atomicBatch(operations: AtomicOperation[]): Promise<void> {
    // Record original state for rollback
    const originalStates: Map<string, { cell: Cell | null; version: number }> = new Map()

    // First, read all cells and validate
    for (const op of operations) {
      const key = this.getCellVersionKey(op.sheetId, op.cellRef)
      if (!originalStates.has(key)) {
        const state = await this.readCellWithVersion(op.sheetId, op.cellRef)
        originalStates.set(key, state)
      }
    }

    // Validate all operations before executing
    for (const op of operations) {
      const key = this.getCellVersionKey(op.sheetId, op.cellRef)
      const state = originalStates.get(key)!

      if (op.type === 'increment') {
        if (!state.cell || state.cell.value.t !== 'number' || typeof state.cell.value.v !== 'number') {
          throw new Error('Cannot increment non-numeric cell')
        }
      }
    }

    // Execute all operations
    const executedOps: Array<{ op: AtomicOperation; key: string; originalVersion: number }> = []

    try {
      for (const op of operations) {
        const key = this.getCellVersionKey(op.sheetId, op.cellRef)
        const state = originalStates.get(key)!
        const originalVersion = state.version

        if (op.type === 'set') {
          await this.writeCellWithVersion(op.sheetId, op.cellRef, op.value!, originalVersion)
          // Update state for subsequent operations on same cell
          const newState = await this.readCellWithVersion(op.sheetId, op.cellRef)
          originalStates.set(key, newState)
        } else if (op.type === 'increment') {
          const currentValue = state.cell!.value.v as number
          const newValue = currentValue + (op.delta || 0)
          await this.writeCellWithVersion(op.sheetId, op.cellRef, newValue, originalVersion)
          // Update state for subsequent operations on same cell
          const newState = await this.readCellWithVersion(op.sheetId, op.cellRef)
          originalStates.set(key, newState)
        } else if (op.type === 'delete') {
          await this.cells.deleteCell(op.sheetId, op.cellRef)
          this.versions.delete(key)
        }

        executedOps.push({ op, key, originalVersion })
      }
    } catch (error) {
      // Rollback executed operations
      for (const { op, key, originalVersion } of executedOps.reverse()) {
        try {
          const originalState = originalStates.get(key)
          if (originalState?.cell) {
            // Restore original cell
            await this.cells.setCell(op.sheetId, op.cellRef, {
              ref: op.cellRef,
              value: originalState.cell.value.v,
              formula: originalState.cell.value.f,
            })
            this.setVersion(key, originalVersion)
          } else if (op.type !== 'delete') {
            // Delete if it was a new cell
            await this.cells.deleteCell(op.sheetId, op.cellRef)
            this.versions.delete(key)
          }
        } catch {
          // Ignore rollback errors
        }
      }
      throw error
    }
  }

  /**
   * Execute a read-modify-write transaction with automatic retry on conflict
   */
  async runTransaction<T>(
    fn: (ctx: RTDBTransactionContext) => Promise<T>,
    options?: RTDBTransactionOptions
  ): Promise<T> {
    const maxRetries = options?.maxRetries ?? 3
    const timeout = options?.timeout ?? 30000
    const isolation = options?.isolation ?? 'read-committed'

    let attempts = 0
    const startTime = Date.now()
    const txTimestamp = startTime

    while (attempts <= maxRetries) {
      // Check timeout
      if (Date.now() - startTime > timeout) {
        throw new TransactionError('Transaction timeout')
      }

      // Snapshot for isolation
      const snapshot: Map<string, { value: unknown; version: number }> = new Map()
      // Track writes within the transaction (not yet committed)
      const pendingWrites: Map<string, { value: unknown; version: number }> = new Map()
      // Track deletes within the transaction
      const pendingDeletes: Set<string> = new Set()

      // Create transaction context
      const ctx: RTDBTransactionContext = {
        read: async <R>(_collection: string, id: string): Promise<{ value: R | null; version: number }> => {
          // Check if we have a pending write for this key
          if (pendingWrites.has(id)) {
            const pending = pendingWrites.get(id)!
            return { value: pending.value as R, version: pending.version }
          }

          // Check if deleted in this transaction
          if (pendingDeletes.has(id)) {
            return { value: null, version: 0 }
          }

          // Check snapshot for isolation
          if (isolation === 'snapshot' && snapshot.has(id)) {
            const snapshotValue = snapshot.get(id)!
            return { value: snapshotValue.value as R, version: snapshotValue.version }
          }

          // Parse the cell reference from the id (format: "sheetId!cellRef")
          const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
          if (!parts) {
            // Not a cell reference, handle as generic read
            const version = this.getVersion(id)
            return { value: null, version }
          }

          const sheetId = parts[1] || parts[2]
          const cellRef = parts[3]

          const { cell, version } = await this.readCellWithVersion(sheetId, cellRef)

          // Store in snapshot for isolation (deep clone to prevent mutation)
          if (isolation === 'snapshot') {
            // Deep clone the cell to prevent external mutations from affecting snapshot
            const clonedCell = cell ? JSON.parse(JSON.stringify(cell)) : null
            snapshot.set(id, { value: clonedCell, version })
            // Return the cloned value to ensure isolation
            return { value: clonedCell as R, version }
          }

          return { value: cell as R, version }
        },

        write: async <W>(
          collection: string,
          id: string,
          value: W,
          expectedVersion: number
        ): Promise<void> => {
          // Check for circular dependency in formulas
          if (collection === 'cells') {
            const cellValue = value as { value?: { f?: string } }
            if (cellValue?.value?.f) {
              // Parse the id to get cell reference
              const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
              if (parts) {
                const cellRef = parts[3]
                const formula = cellValue.value.f
                // Simple circular dependency check
                if (formula.includes(cellRef)) {
                  throw new Error('Circular dependency detected')
                }
                // Check for indirect circular dependencies
                // e.g., A1 depends on B1, B1 depends on A1
                const referencedCells = formula.match(/[A-Z]+\d+/gi) || []
                for (const refCell of referencedCells) {
                  // Read the referenced cell to check its formula
                  const refId = id.replace(/![A-Z]+\d+$/i, `!${refCell}`)
                  const refParts = refId.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
                  if (refParts) {
                    const refSheetId = refParts[1] || refParts[2]
                    const existingCell = await this.cells.getCell(refSheetId, refCell)
                    if (existingCell?.value.f) {
                      // Check if the referenced cell's formula references back to us
                      if (existingCell.value.f.includes(cellRef)) {
                        throw new Error('Circular dependency detected')
                      }
                    }
                  }
                }
              }
            }
          }

          // Get current version from storage, not snapshot
          const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
          if (!parts) {
            throw new Error(`Invalid cell id format: ${id}`)
          }

          const sheetId = parts[1] || parts[2]
          const cellRef = parts[3]
          const key = this.getCellVersionKey(sheetId, cellRef)
          const currentVersion = this.getVersion(key)

          // Check version
          if (expectedVersion !== currentVersion) {
            throw new ConflictError(
              `Version conflict: expected ${expectedVersion}, got ${currentVersion}`,
              collection,
              id,
              expectedVersion,
              currentVersion
            )
          }

          // Store pending write
          pendingWrites.set(id, { value, version: expectedVersion + 1 })
        },

        delete: async (collection: string, id: string, expectedVersion: number): Promise<void> => {
          const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
          if (!parts) {
            throw new Error(`Invalid cell id format: ${id}`)
          }

          const sheetId = parts[1] || parts[2]
          const cellRef = parts[3]
          const key = this.getCellVersionKey(sheetId, cellRef)
          const currentVersion = this.getVersion(key)

          if (expectedVersion !== currentVersion) {
            throw new ConflictError(
              `Version conflict: expected ${expectedVersion}, got ${currentVersion}`,
              collection,
              id,
              expectedVersion,
              currentVersion,
              'delete'
            )
          }

          pendingDeletes.add(id)
        },

        timestamp: (): number => txTimestamp,
      }

      try {
        // Create a promise that will race with the timeout
        const timeoutPromise = new Promise<never>((_, reject) => {
          const remaining = timeout - (Date.now() - startTime)
          if (remaining <= 0) {
            reject(new TransactionError('Transaction timeout'))
          } else {
            setTimeout(() => reject(new TransactionError('Transaction timeout')), remaining)
          }
        })

        // Execute the transaction function with timeout
        const result = await Promise.race([fn(ctx), timeoutPromise])

        // Commit pending writes
        for (const [id, { value }] of pendingWrites) {
          const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
          if (parts) {
            const sheetId = parts[1] || parts[2]
            const cellRef = parts[3]

            // Convert Cell to CellData for setCell
            const cellValue = value as Cell | { value: CellValue }
            let cellData: CellData
            if ('_id' in cellValue) {
              // It's a Cell object
              cellData = {
                ref: cellRef,
                value: cellValue.value.v,
                formula: cellValue.value.f,
              }
            } else {
              // It's a partial value object
              cellData = {
                ref: cellRef,
                value: cellValue.value?.v,
                formula: cellValue.value?.f,
              }
            }

            // setCell callback will increment version
            await this.cells.setCell(sheetId, cellRef, cellData)
          }
        }

        // Commit pending deletes
        for (const id of pendingDeletes) {
          const parts = id.match(/^(?:'([^']+)'|([^!]+))!(.+)$/)
          if (parts) {
            const sheetId = parts[1] || parts[2]
            const cellRef = parts[3]
            await this.cells.deleteCell(sheetId, cellRef)
            const key = this.getCellVersionKey(sheetId, cellRef)
            this.versions.delete(key)
          }
        }

        return result
      } catch (error) {
        if (error instanceof ConflictError) {
          attempts++
          if (attempts > maxRetries) {
            // Check if this was a delete conflict
            // If so, throw ConflictError; otherwise, throw TransactionError
            if (error.operation === 'delete') {
              throw error
            }
            throw new TransactionError('Transaction failed after max retries', error)
          }
          // Retry on conflict
          continue
        }
        // Don't retry non-conflict errors
        throw error
      }
    }

    throw new TransactionError('Transaction failed after max retries')
  }
}

// ============================================================================
// Index Helper Functions
// ============================================================================

/**
 * Create an index on a collection
 */
export async function createIndex(
  db: MongoDatabase,
  collection: string,
  spec: IndexSpec,
  options?: IndexOptions
): Promise<string> {
  const col = db.collection(collection)
  return col.createIndex(spec, options)
}

/**
 * Create all predefined indexes for cells collection
 */
export async function createCellIndexes(db: MongoDatabase): Promise<void> {
  const col = db.collection('cells')
  for (const { spec, options } of CELL_INDEXES) {
    await col.createIndex(spec, options)
  }
}

/**
 * Create all predefined indexes for sheets collection
 */
export async function createSheetIndexes(db: MongoDatabase): Promise<void> {
  const col = db.collection('sheets')
  for (const { spec, options } of SHEET_INDEXES) {
    await col.createIndex(spec, options)
  }
}

/**
 * Create all predefined indexes for workbooks collection
 */
export async function createWorkbookIndexes(db: MongoDatabase): Promise<void> {
  const col = db.collection('workbooks')
  for (const { spec, options } of WORKBOOK_INDEXES) {
    await col.createIndex(spec, options)
  }
}

// ============================================================================
// Factory Function
// ============================================================================

/**
 * Create a storage manager connected to mongo.do
 */
export function createStorageManager(client: MongoClient, dbName?: string): StorageManager {
  return new StorageManager(client, dbName)
}
