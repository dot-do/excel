/**
 * Database.do Hierarchical Path Storage
 *
 * TDD RED STUB: Minimal implementation to make tests compile and fail.
 * This file provides only the type definitions and exports needed for tests.
 * The actual implementation will be done in the GREEN phase.
 */

// ============================================================================
// Error Classes
// ============================================================================

export class PathNotFoundError extends Error {
  constructor(public readonly path: string) {
    super(`Path not found: ${path}`)
    this.name = 'PathNotFoundError'
  }
}

export class InvalidPathError extends Error {
  constructor(public readonly path: string, reason?: string) {
    super(`Invalid path: ${path}${reason ? ` - ${reason}` : ''}`)
    this.name = 'InvalidPathError'
  }
}

// ============================================================================
// Types
// ============================================================================

/** Document stored at a path */
export interface PathDocument {
  path: string
  data: unknown
  metadata: Record<string, unknown>
  createdAt: Date
  updatedAt: Date
}

/** Query options for path-based queries */
export interface PathQuery {
  prefix: string
  where?: Record<string, unknown>
  orderBy?: { field: string; direction: 'asc' | 'desc' }
  limit?: number
  offset?: number
  select?: string[]
}

/** Event emitted when data changes */
export interface PathEvent {
  type: 'create' | 'update' | 'delete'
  path: string
  data?: unknown
  timestamp: Date
}

/** Watch callback function */
export type WatchCallback = (event: PathEvent) => void

/** Unwatch function to stop watching */
export type UnwatchFn = () => void

/** Transaction context for atomic operations */
export interface TransactionContext {
  get(path: string): Promise<PathDocument | null>
  set(path: string, data: unknown, options?: SetOptions): Promise<PathDocument>
  delete(path: string): Promise<boolean>
}

/** Set options */
export interface SetOptions {
  metadata?: Record<string, unknown>
}

/** List options */
export interface ListOptions {
  recursive?: boolean
  limit?: number
  offset?: number
  filter?: { metadata?: Record<string, unknown> }
}

/** Delete options */
export interface DeleteOptions {
  recursive?: boolean
}

/** Move/Copy options */
export interface MoveOptions {
  recursive?: boolean
}

/** Count options */
export interface CountOptions {
  recursive?: boolean
}

// ============================================================================
// Database Client Interface
// ============================================================================

export interface DatabaseClient {
  get(path: string): Promise<PathDocument | null>
  set(path: string, data: unknown, options?: SetOptions): Promise<PathDocument>
  delete(path: string): Promise<boolean>
  list(prefix: string, options?: { recursive?: boolean; limit?: number }): Promise<PathDocument[]>
  exists(path: string): Promise<boolean>
  close(): Promise<void>
}

export interface DatabaseClientOptions {
  url: string
  namespace?: string
  token?: string
}

// ============================================================================
// Path Utility Functions (STUBS - Will fail tests)
// ============================================================================

export interface ParsedPath {
  segments: string[]
  isAbsolute: boolean
}

export function parsePath(_path: string): ParsedPath {
  throw new Error('Not implemented: parsePath')
}

export function normalizePath(_path: string): string {
  throw new Error('Not implemented: normalizePath')
}

export function getParentPath(_path: string): string | null {
  throw new Error('Not implemented: getParentPath')
}

export function getPathSegments(_path: string): string[] {
  throw new Error('Not implemented: getPathSegments')
}

export function joinPath(..._segments: string[]): string {
  throw new Error('Not implemented: joinPath')
}

export function isValidPath(_path: string, _options?: { strict?: boolean }): boolean {
  throw new Error('Not implemented: isValidPath')
}

// ============================================================================
// PathStore (STUB - Will fail tests)
// ============================================================================

export class PathStore {
  constructor(_client: DatabaseClient) {
    // Stub constructor
  }

  async get(_path: string): Promise<PathDocument | null> {
    throw new Error('Not implemented: PathStore.get')
  }

  async set(_path: string, _data: unknown, _options?: SetOptions): Promise<PathDocument> {
    throw new Error('Not implemented: PathStore.set')
  }

  async delete(_path: string, _options?: DeleteOptions): Promise<boolean | number> {
    throw new Error('Not implemented: PathStore.delete')
  }

  async list(_prefix: string, _options?: ListOptions): Promise<PathDocument[]> {
    throw new Error('Not implemented: PathStore.list')
  }

  async exists(_path: string): Promise<boolean> {
    throw new Error('Not implemented: PathStore.exists')
  }

  async count(_prefix: string, _options?: CountOptions): Promise<number> {
    throw new Error('Not implemented: PathStore.count')
  }

  async move(_source: string, _target: string, _options?: MoveOptions): Promise<void> {
    throw new Error('Not implemented: PathStore.move')
  }

  async copy(_source: string, _target: string, _options?: MoveOptions): Promise<void> {
    throw new Error('Not implemented: PathStore.copy')
  }

  async query(_query: PathQuery): Promise<PathDocument[]> {
    throw new Error('Not implemented: PathStore.query')
  }

  watch(_prefix: string, _callback: WatchCallback): UnwatchFn {
    throw new Error('Not implemented: PathStore.watch')
  }

  async transaction<T>(_fn: (tx: TransactionContext) => Promise<T>): Promise<T> {
    throw new Error('Not implemented: PathStore.transaction')
  }

  async close(): Promise<void> {
    throw new Error('Not implemented: PathStore.close')
  }
}

// ============================================================================
// Factory Function (STUB - Will fail tests)
// ============================================================================

export function createDatabaseClient(_options: DatabaseClientOptions): DatabaseClient {
  throw new Error('Not implemented: createDatabaseClient')
}
