/**
 * MongoDB Storage Layer for excel.do
 *
 * Provides storage operations for cells, sheets, and workbooks using mongo.do.
 */

import type { Cell, CellValue, CellPrimitive, CellFormat, CellMetadata } from '../types'
import type {
  MongoClient,
  MongoDatabase,
  MongoCollection,
  MongoSession,
  MongoFilter,
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
 */
function indexToCol(index: number): string {
  let col = ''
  let n = index + 1
  while (n > 0) {
    const remainder = (n - 1) % 26
    col = String.fromCharCode(65 + remainder) + col
    n = Math.floor((n - 1) / 26)
  }
  return col
}

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
// CellStore Implementation
// ============================================================================

export class CellStore implements ICellStore {
  private collection: MongoCollection<Cell>

  constructor(
    private readonly db: MongoDatabase,
    private readonly collectionName: string = 'cells'
  ) {
    this.collection = db.collection<Cell>(collectionName)
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

    // Check if cell exists
    const existing = await this.getCell(sheetId, cellRef)

    if (existing) {
      // Update existing cell
      cell._id = existing._id
      cell.createdAt = existing.createdAt
      cell.updatedAt = new Date()

      await this.collection.updateOne(
        { _id: existing._id },
        { $set: cell }
      )
    } else {
      // Insert new cell
      await this.collection.insertOne(cell)
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
}

// ============================================================================
// SheetStore Implementation
// ============================================================================

export class SheetStore implements ISheetStore {
  private collection: MongoCollection<Sheet>

  constructor(
    private readonly db: MongoDatabase,
    private readonly collectionName: string = 'sheets'
  ) {
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

  async reorderSheets(workbookId: string, sheetIds: string[]): Promise<void> {
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

  constructor(
    private readonly db: MongoDatabase,
    private readonly collectionName: string = 'workbooks'
  ) {
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
// StorageManager Implementation
// ============================================================================

export class StorageManager implements IStorageManager {
  public readonly cells: CellStore
  public readonly sheets: SheetStore
  public readonly workbooks: WorkbookStore
  private readonly db: MongoDatabase

  constructor(private readonly client: MongoClient, dbName: string = 'excel') {
    this.db = client.db(dbName)
    this.cells = new CellStore(this.db)
    this.sheets = new SheetStore(this.db)
    this.workbooks = new WorkbookStore(this.db)
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
