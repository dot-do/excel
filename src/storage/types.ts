/**
 * Storage Layer Types for excel.do
 *
 * Type definitions for mongo.do storage operations.
 */

import type { Cell, CellValue, CellPrimitive, CellFormat, CellMetadata } from '../types'

// ============================================================================
// Mongo.do Client Interface
// ============================================================================

/** MongoDB-style filter query */
export interface MongoFilter<T = unknown> {
  [key: string]: T | MongoFilter<T> | unknown
}

/** MongoDB-style update operations */
export interface MongoUpdate<T> {
  $set?: Partial<T>
  $unset?: Partial<Record<keyof T, 1>>
  $push?: Partial<Record<keyof T, unknown>>
  $pull?: Partial<Record<keyof T, unknown>>
  $inc?: Partial<Record<keyof T, number>>
}

/** MongoDB-style sort specification */
export interface MongoSort {
  [key: string]: 1 | -1
}

/** Find options */
export interface FindOptions {
  limit?: number
  skip?: number
  sort?: MongoSort
  projection?: Record<string, 0 | 1>
}

/** Insert result */
export interface InsertResult {
  insertedId: string
  acknowledged: boolean
}

/** Insert many result */
export interface InsertManyResult {
  insertedIds: string[]
  insertedCount: number
  acknowledged: boolean
}

/** Update result */
export interface UpdateResult {
  matchedCount: number
  modifiedCount: number
  upsertedId?: string
  acknowledged: boolean
}

/** Delete result */
export interface DeleteResult {
  deletedCount: number
  acknowledged: boolean
}

/** Index specification */
export interface IndexSpec {
  [key: string]: 1 | -1 | 'text' | '2dsphere' | '2d' | 'hashed'
}

/** Index options */
export interface IndexOptions {
  unique?: boolean
  sparse?: boolean
  background?: boolean
  name?: string
  expireAfterSeconds?: number
  partialFilterExpression?: MongoFilter
}

/** Index information */
export interface IndexInfo {
  name: string
  key: IndexSpec
  unique?: boolean
  sparse?: boolean
}

/** Collection interface matching mongo.do client */
export interface MongoCollection<T> {
  findOne(filter: MongoFilter): Promise<T | null>
  find(filter: MongoFilter, options?: FindOptions): Promise<T[]>
  insertOne(doc: T): Promise<InsertResult>
  insertMany(docs: T[]): Promise<InsertManyResult>
  updateOne(filter: MongoFilter, update: MongoUpdate<T>): Promise<UpdateResult>
  updateMany(filter: MongoFilter, update: MongoUpdate<T>): Promise<UpdateResult>
  deleteOne(filter: MongoFilter): Promise<DeleteResult>
  deleteMany(filter: MongoFilter): Promise<DeleteResult>
  countDocuments(filter?: MongoFilter): Promise<number>
  createIndex(spec: IndexSpec, options?: IndexOptions): Promise<string>
  dropIndex(name: string): Promise<void>
  listIndexes(): Promise<IndexInfo[]>
}

/** Database interface matching mongo.do client */
export interface MongoDatabase {
  collection<T>(name: string): MongoCollection<T>
}

/** Session interface for transactions */
export interface MongoSession {
  startTransaction(): void
  commitTransaction(): Promise<void>
  abortTransaction(): Promise<void>
  endSession(): Promise<void>
}

/** Client interface matching mongo.do client */
export interface MongoClient {
  db(name?: string): MongoDatabase
  startSession(): Promise<MongoSession>
  close(): Promise<void>
}

// ============================================================================
// Sheet and Workbook Types
// ============================================================================

/** Sheet document stored in mongo.do */
export interface Sheet {
  _id: string
  workbookId: string
  name: string
  index: number
  /** Default column width */
  defaultColWidth?: number
  /** Default row height */
  defaultRowHeight?: number
  /** Column widths by index */
  colWidths?: Record<number, number>
  /** Row heights by row number */
  rowHeights?: Record<number, number>
  /** Hidden columns */
  hiddenCols?: number[]
  /** Hidden rows */
  hiddenRows?: number[]
  /** Frozen panes */
  freeze?: {
    rows?: number
    cols?: number
  }
  /** Sheet-level protection */
  protection?: {
    password?: string
    sheet?: boolean
    objects?: boolean
    scenarios?: boolean
    formatCells?: boolean
    formatColumns?: boolean
    formatRows?: boolean
    insertColumns?: boolean
    insertRows?: boolean
    insertHyperlinks?: boolean
    deleteColumns?: boolean
    deleteRows?: boolean
    sort?: boolean
    autoFilter?: boolean
    pivotTables?: boolean
  }
  createdAt: Date
  updatedAt: Date
}

/** Options for creating a sheet */
export interface CreateSheetOptions {
  workbookId: string
  name: string
  index?: number
  defaultColWidth?: number
  defaultRowHeight?: number
}

/** Options for updating a sheet */
export interface UpdateSheetOptions {
  name?: string
  index?: number
  defaultColWidth?: number
  defaultRowHeight?: number
  colWidths?: Record<number, number>
  rowHeights?: Record<number, number>
  hiddenCols?: number[]
  hiddenRows?: number[]
  freeze?: { rows?: number; cols?: number }
  protection?: Sheet['protection']
}

/** Workbook document stored in mongo.do */
export interface Workbook {
  _id: string
  name: string
  /** Active sheet index */
  activeSheet?: number
  /** Named ranges at workbook level */
  namedRanges?: Record<string, string>
  /** Workbook-level styles */
  styles?: Record<string, CellFormat>
  /** Custom number formats */
  numberFormats?: Record<number, string>
  /** Workbook properties */
  properties?: {
    title?: string
    subject?: string
    author?: string
    keywords?: string
    comments?: string
    lastAuthor?: string
    created?: Date
    modified?: Date
  }
  createdAt: Date
  updatedAt: Date
}

/** Options for creating a workbook */
export interface CreateWorkbookOptions {
  name: string
  properties?: Workbook['properties']
}

// ============================================================================
// Cell Storage Types
// ============================================================================

/** Cell reference for storage operations */
export type CellRef = string // e.g., "A1", "B10"

/** Range specification */
export interface Range {
  startCol: string
  startRow: number
  endCol: string
  endRow: number
}

/** Cell data for bulk operations */
export interface CellData {
  ref: CellRef
  value?: CellPrimitive
  formula?: string
  format?: CellFormat
  metadata?: CellMetadata
}

/** Query options for finding cells */
export interface CellQuery {
  /** Filter by value */
  value?: CellPrimitive
  /** Filter by value type */
  valueType?: CellValue['t']
  /** Filter by formula presence */
  hasFormula?: boolean
  /** Filter by formula content (regex) */
  formulaPattern?: string
  /** Filter by row range */
  rowRange?: { min?: number; max?: number }
  /** Filter by column range */
  colRange?: { minIndex?: number; maxIndex?: number }
}

// ============================================================================
// Store Interfaces
// ============================================================================

/** Cell storage operations */
export interface ICellStore {
  /** Get a single cell */
  getCell(sheetId: string, cellRef: CellRef): Promise<Cell | null>

  /** Set a single cell value */
  setCell(sheetId: string, cellRef: CellRef, value: CellPrimitive | CellData): Promise<Cell>

  /** Get cells in a range */
  getCells(sheetId: string, range: Range | string): Promise<Cell[]>

  /** Set multiple cells */
  setCells(sheetId: string, cells: CellData[]): Promise<Cell[]>

  /** Delete a cell */
  deleteCell(sheetId: string, cellRef: CellRef): Promise<boolean>

  /** Delete cells in a range */
  deleteCells(sheetId: string, range: Range | string): Promise<number>

  /** Find cells matching query */
  findCells(sheetId: string, query: CellQuery): Promise<Cell[]>

  /** Count cells matching query */
  countCells(sheetId: string, query?: CellQuery): Promise<number>
}

/** Sheet storage operations */
export interface ISheetStore {
  /** Create a new sheet */
  createSheet(workbookId: string, name: string, options?: Omit<CreateSheetOptions, 'workbookId' | 'name'>): Promise<Sheet>

  /** Get a sheet by ID */
  getSheet(id: string): Promise<Sheet | null>

  /** Get a sheet by name within a workbook */
  getSheetByName(workbookId: string, name: string): Promise<Sheet | null>

  /** Update a sheet */
  updateSheet(id: string, updates: UpdateSheetOptions): Promise<Sheet | null>

  /** Delete a sheet */
  deleteSheet(id: string): Promise<boolean>

  /** List sheets in a workbook */
  listSheets(workbookId: string): Promise<Sheet[]>

  /** Reorder sheets */
  reorderSheets(workbookId: string, sheetIds: string[]): Promise<void>
}

/** Workbook storage operations */
export interface IWorkbookStore {
  /** Create a new workbook */
  createWorkbook(name: string, options?: Omit<CreateWorkbookOptions, 'name'>): Promise<Workbook>

  /** Get a workbook by ID */
  getWorkbook(id: string): Promise<Workbook | null>

  /** Update a workbook */
  updateWorkbook(id: string, updates: Partial<Omit<Workbook, '_id' | 'createdAt'>>): Promise<Workbook | null>

  /** Delete a workbook */
  deleteWorkbook(id: string): Promise<boolean>

  /** List all workbooks */
  listWorkbooks(): Promise<Workbook[]>
}

/** Transaction function type */
export type TransactionFn<T> = (session: MongoSession) => Promise<T>

/** Storage manager with transaction support */
export interface IStorageManager {
  readonly cells: ICellStore
  readonly sheets: ISheetStore
  readonly workbooks: IWorkbookStore

  /** Execute operations in a transaction */
  withTransaction<T>(fn: TransactionFn<T>): Promise<T>

  /** Create indexes for optimal performance */
  createIndexes(): Promise<void>

  /** Close the connection */
  close(): Promise<void>
}

// ============================================================================
// Index Definitions
// ============================================================================

/** Predefined indexes for cells collection */
export const CELL_INDEXES: Array<{ spec: IndexSpec; options?: IndexOptions }> = [
  // Compound index for sheet + cell lookup
  { spec: { sheet: 1, col: 1, row: 1 }, options: { unique: true } },
  // Index for range queries
  { spec: { sheet: 1, colIndex: 1, row: 1 }, options: {} },
  // Index for formula cells
  { spec: { sheet: 1, 'value.f': 1 }, options: { sparse: true } },
  // Index for dependencies tracking
  { spec: { dependencies: 1 }, options: { sparse: true } },
  { spec: { dependents: 1 }, options: { sparse: true } },
  // Index for recent updates
  { spec: { updatedAt: -1 }, options: {} },
]

/** Predefined indexes for sheets collection */
export const SHEET_INDEXES: Array<{ spec: IndexSpec; options?: IndexOptions }> = [
  // Unique name per workbook
  { spec: { workbookId: 1, name: 1 }, options: { unique: true } },
  // Index order lookup
  { spec: { workbookId: 1, index: 1 } },
]

/** Predefined indexes for workbooks collection */
export const WORKBOOK_INDEXES: Array<{ spec: IndexSpec; options?: IndexOptions }> = [
  // Workbook name index
  { spec: { name: 1 } },
  // Recent updates
  { spec: { updatedAt: -1 } },
]
