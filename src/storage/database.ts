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
// Path Utility Functions
// ============================================================================

export interface ParsedPath {
  segments: string[]
  isAbsolute: boolean
}

const INVALID_PATH_CHARS = /[?#<>\x00]/
const MAX_PATH_LENGTH = 1024
const MAX_PATH_DEPTH = 64

export function parsePath(path: string): ParsedPath {
  const isAbsolute = path.startsWith('/')
  const segments = path
    .split('/')
    .filter((s) => s !== '' && s !== '.')
    .filter(Boolean)

  return { segments, isAbsolute }
}

export function normalizePath(path: string): string {
  const isAbsolute = path.startsWith('/')
  const segments = path.split('/').filter((s) => s !== '' && s !== '.')

  const normalizedSegments: string[] = []
  for (const segment of segments) {
    if (segment === '..') {
      if (normalizedSegments.length > 0) {
        normalizedSegments.pop()
      }
    } else {
      normalizedSegments.push(segment)
    }
  }

  const result = normalizedSegments.join('/')
  return isAbsolute ? `/${result}` || '/' : result
}

export function getParentPath(path: string): string | null {
  const normalized = normalizePath(path)
  if (normalized === '/') {
    return null
  }

  const segments = normalized.split('/').filter(Boolean)
  if (segments.length <= 1) {
    return '/'
  }

  segments.pop()
  return '/' + segments.join('/')
}

export function getPathSegments(path: string): string[] {
  return parsePath(path).segments
}

export function joinPath(...segments: string[]): string {
  const parts: string[] = []
  let isAbsolute = false

  for (const segment of segments) {
    if (!segment) continue
    const cleaned = segment.replace(/^\/+|\/+$/g, '')
    if (segment.startsWith('/') && parts.length === 0) {
      isAbsolute = true
    }
    if (cleaned) {
      parts.push(cleaned)
    }
  }

  const result = parts.join('/')
  return isAbsolute ? `/${result}` : result
}

export function isValidPath(path: string, options?: { strict?: boolean }): boolean {
  // Check for invalid characters
  if (INVALID_PATH_CHARS.test(path)) {
    return false
  }

  // Check max length
  if (path.length > MAX_PATH_LENGTH) {
    return false
  }

  // Check depth
  const segments = getPathSegments(path)
  if (segments.length > MAX_PATH_DEPTH) {
    return false
  }

  // Strict mode requires absolute path
  if (options?.strict && !path.startsWith('/')) {
    return false
  }

  return true
}

// ============================================================================
// PathStore Implementation
// ============================================================================

interface WatcherEntry {
  id: string
  prefix: string
  callback: WatchCallback
}

export class PathStore {
  private client: DatabaseClient
  private watchers: WatcherEntry[] = []
  private watcherIdCounter = 0

  constructor(client: DatabaseClient) {
    this.client = client
  }

  private validatePath(path: string): void {
    if (!isValidPath(path)) {
      throw new InvalidPathError(path)
    }
  }

  private emitEvent(event: PathEvent): void {
    for (const watcher of this.watchers) {
      if (event.path.startsWith(watcher.prefix)) {
        try {
          watcher.callback(event)
        } catch {
          // Continue emitting to other watchers
        }
      }
    }
  }

  async get(path: string): Promise<PathDocument | null> {
    const normalized = normalizePath(path)
    this.validatePath(normalized)
    return this.client.get(normalized)
  }

  async set(path: string, data: unknown, options?: SetOptions): Promise<PathDocument> {
    const normalized = normalizePath(path)
    this.validatePath(normalized)

    // Create parent paths implicitly
    const parent = getParentPath(normalized)
    if (parent && parent !== '/') {
      const parentExists = await this.client.exists(parent)
      if (!parentExists) {
        await this.client.set(parent, {})
      }
    }

    const existing = await this.client.get(normalized)
    const result = await this.client.set(normalized, data, options)

    this.emitEvent({
      type: existing ? 'update' : 'create',
      path: normalized,
      data,
      timestamp: new Date(),
    })

    return result
  }

  async delete(path: string, options?: DeleteOptions): Promise<boolean | number> {
    const normalized = normalizePath(path)
    this.validatePath(normalized)

    if (options?.recursive) {
      const children = await this.client.list(normalized, { recursive: true })
      let count = 0

      // Delete children first
      for (const child of children) {
        if (child.path !== normalized) {
          const deleted = await this.client.delete(child.path)
          if (deleted) {
            count++
            this.emitEvent({
              type: 'delete',
              path: child.path,
              timestamp: new Date(),
            })
          }
        }
      }

      // Delete the main path
      const mainExists = await this.client.exists(normalized)
      if (mainExists) {
        await this.client.delete(normalized)
        count++
        this.emitEvent({
          type: 'delete',
          path: normalized,
          timestamp: new Date(),
        })
      }

      return count
    }

    const result = await this.client.delete(normalized)
    if (result) {
      this.emitEvent({
        type: 'delete',
        path: normalized,
        timestamp: new Date(),
      })
    }
    return result
  }

  async list(prefix: string, options?: ListOptions): Promise<PathDocument[]> {
    const normalized = normalizePath(prefix)
    let results = await this.client.list(normalized, {
      recursive: options?.recursive,
      limit: undefined, // Get all first, then apply pagination
    })

    // Filter by metadata if specified
    if (options?.filter?.metadata) {
      results = results.filter((doc) => {
        for (const [key, value] of Object.entries(options.filter!.metadata!)) {
          if (doc.metadata[key] !== value) {
            return false
          }
        }
        return true
      })
    }

    // Sort by path
    results.sort((a, b) => a.path.localeCompare(b.path))

    // Apply pagination
    if (options?.offset) {
      results = results.slice(options.offset)
    }
    if (options?.limit) {
      results = results.slice(0, options.limit)
    }

    return results
  }

  async exists(path: string): Promise<boolean> {
    const normalized = normalizePath(path)
    return this.client.exists(normalized)
  }

  async count(prefix: string, options?: CountOptions): Promise<number> {
    const results = await this.list(prefix, { recursive: options?.recursive })
    return results.length
  }

  async move(source: string, target: string, options?: MoveOptions): Promise<void> {
    const normalizedSource = normalizePath(source)
    const normalizedTarget = normalizePath(target)

    const sourceDoc = await this.client.get(normalizedSource)
    if (!sourceDoc) {
      throw new PathNotFoundError(normalizedSource)
    }

    if (options?.recursive) {
      const children = await this.client.list(normalizedSource, { recursive: true })

      // Move main document
      await this.client.set(normalizedTarget, sourceDoc.data, { metadata: sourceDoc.metadata })
      await this.client.delete(normalizedSource)

      // Move children
      for (const child of children) {
        if (child.path !== normalizedSource) {
          const relativePath = child.path.slice(normalizedSource.length)
          const newPath = normalizedTarget + relativePath
          await this.client.set(newPath, child.data, { metadata: child.metadata })
          await this.client.delete(child.path)
        }
      }
    } else {
      await this.client.set(normalizedTarget, sourceDoc.data, { metadata: sourceDoc.metadata })
      await this.client.delete(normalizedSource)
    }
  }

  async copy(source: string, target: string, options?: MoveOptions): Promise<void> {
    const normalizedSource = normalizePath(source)
    const normalizedTarget = normalizePath(target)

    const sourceDoc = await this.client.get(normalizedSource)
    if (!sourceDoc) {
      throw new PathNotFoundError(normalizedSource)
    }

    if (options?.recursive) {
      const children = await this.client.list(normalizedSource, { recursive: true })

      // Copy main document
      await this.client.set(normalizedTarget, sourceDoc.data, { metadata: sourceDoc.metadata })

      // Copy children
      for (const child of children) {
        if (child.path !== normalizedSource) {
          const relativePath = child.path.slice(normalizedSource.length)
          const newPath = normalizedTarget + relativePath
          await this.client.set(newPath, child.data, { metadata: child.metadata })
        }
      }
    } else {
      await this.client.set(normalizedTarget, sourceDoc.data, { metadata: sourceDoc.metadata })
    }
  }

  async query(query: PathQuery): Promise<PathDocument[]> {
    let results = await this.list(query.prefix, { recursive: true })

    // Apply where clause
    if (query.where) {
      results = results.filter((doc) => this.matchesWhere(doc.data, query.where!))
    }

    // Apply ordering
    if (query.orderBy) {
      const { field, direction } = query.orderBy
      results.sort((a, b) => {
        const aVal = (a.data as Record<string, unknown>)?.[field]
        const bVal = (b.data as Record<string, unknown>)?.[field]
        if (aVal === bVal) return 0
        if (aVal === undefined || aVal === null) return 1
        if (bVal === undefined || bVal === null) return -1
        const cmp = (aVal as number | string) < (bVal as number | string) ? -1 : 1
        return direction === 'desc' ? -cmp : cmp
      })
    }

    // Apply projection
    if (query.select) {
      results = results.map((doc) => ({
        ...doc,
        data: this.projectData(doc.data, query.select!),
      }))
    }

    // Apply pagination
    if (query.offset) {
      results = results.slice(query.offset)
    }
    if (query.limit) {
      results = results.slice(0, query.limit)
    }

    return results
  }

  private matchesWhere(data: unknown, where: Record<string, unknown>): boolean {
    if (!data || typeof data !== 'object') return false
    const obj = data as Record<string, unknown>

    for (const [key, condition] of Object.entries(where)) {
      if (key === '$or') {
        const orConditions = condition as Record<string, unknown>[]
        if (!orConditions.some((c) => this.matchesWhere(data, c))) {
          return false
        }
        continue
      }

      const value = obj[key]

      if (condition && typeof condition === 'object' && !Array.isArray(condition)) {
        const ops = condition as Record<string, unknown>
        for (const [op, opVal] of Object.entries(ops)) {
          const numValue = value as number
          const numOpVal = opVal as number
          switch (op) {
            case '$gte':
              if (!(value !== undefined && value !== null && numValue >= numOpVal)) return false
              break
            case '$gt':
              if (!(value !== undefined && value !== null && numValue > numOpVal)) return false
              break
            case '$lte':
              if (!(value !== undefined && value !== null && numValue <= numOpVal)) return false
              break
            case '$lt':
              if (!(value !== undefined && value !== null && numValue < numOpVal)) return false
              break
            case '$ne':
              if (value === opVal) return false
              break
          }
        }
      } else if (value !== condition) {
        return false
      }
    }

    return true
  }

  private projectData(data: unknown, fields: string[]): unknown {
    if (!data || typeof data !== 'object') return data
    const obj = data as Record<string, unknown>
    const result: Record<string, unknown> = {}
    for (const field of fields) {
      if (field in obj) {
        result[field] = obj[field]
      }
    }
    return result
  }

  watch(prefix: string, callback: WatchCallback): UnwatchFn {
    const id = `watcher_${++this.watcherIdCounter}`
    const normalized = normalizePath(prefix)
    const entry: WatcherEntry = { id, prefix: normalized, callback }
    this.watchers.push(entry)

    return () => {
      const index = this.watchers.findIndex((w) => w.id === id)
      if (index >= 0) {
        this.watchers.splice(index, 1)
      }
    }
  }

  async transaction<T>(fn: (tx: TransactionContext) => Promise<T>): Promise<T> {
    // Simple implementation: track changes and rollback on error
    const changes: Array<{ path: string; previousDoc: PathDocument | null }> = []

    const tx: TransactionContext = {
      get: async (path: string): Promise<PathDocument | null> => {
        const normalized = normalizePath(path)
        return this.client.get(normalized)
      },
      set: async (path: string, data: unknown, options?: SetOptions): Promise<PathDocument> => {
        const normalized = normalizePath(path)
        const previousDoc = await this.client.get(normalized)
        changes.push({ path: normalized, previousDoc })
        return this.client.set(normalized, data, options)
      },
      delete: async (path: string): Promise<boolean> => {
        const normalized = normalizePath(path)
        const previousDoc = await this.client.get(normalized)
        if (previousDoc) {
          changes.push({ path: normalized, previousDoc })
        }
        return this.client.delete(normalized)
      },
    }

    try {
      const result = await fn(tx)
      return result
    } catch (error) {
      // Rollback changes
      for (const change of changes.reverse()) {
        if (change.previousDoc) {
          await this.client.set(change.path, change.previousDoc.data, {
            metadata: change.previousDoc.metadata,
          })
        } else {
          await this.client.delete(change.path)
        }
      }
      throw error
    }
  }

  async close(): Promise<void> {
    this.watchers = []
    await this.client.close()
  }
}

// ============================================================================
// Factory Function
// ============================================================================

class InMemoryDatabaseClient implements DatabaseClient {
  private storage = new Map<string, PathDocument>()
  private options: DatabaseClientOptions

  constructor(options: DatabaseClientOptions) {
    this.options = options
  }

  async get(path: string): Promise<PathDocument | null> {
    const key = this.getKey(path)
    return this.storage.get(key) || null
  }

  async set(path: string, data: unknown, options?: { metadata?: Record<string, unknown> }): Promise<PathDocument> {
    const key = this.getKey(path)
    const now = new Date()
    const existing = this.storage.get(key)

    const doc: PathDocument = {
      path,
      data,
      metadata: options?.metadata || {},
      createdAt: existing?.createdAt || now,
      updatedAt: now,
    }

    this.storage.set(key, doc)
    return doc
  }

  async delete(path: string): Promise<boolean> {
    const key = this.getKey(path)
    return this.storage.delete(key)
  }

  async list(prefix: string, options?: { recursive?: boolean; limit?: number }): Promise<PathDocument[]> {
    const results: PathDocument[] = []
    const normalizedPrefix = prefix.endsWith('/') ? prefix : prefix + '/'

    for (const [_key, doc] of this.storage.entries()) {
      // Check if path matches prefix
      if (doc.path === prefix || doc.path.startsWith(normalizedPrefix)) {
        if (!options?.recursive && doc.path !== prefix) {
          // Only include direct children
          const relative = doc.path.slice(normalizedPrefix.length)
          if (relative.includes('/')) {
            continue
          }
        }
        results.push(doc)
      }
    }

    return options?.limit ? results.slice(0, options.limit) : results
  }

  async exists(path: string): Promise<boolean> {
    const key = this.getKey(path)
    return this.storage.has(key)
  }

  async close(): Promise<void> {
    this.storage.clear()
  }

  private getKey(path: string): string {
    const namespace = this.options.namespace || 'default'
    return `${namespace}:${path}`
  }
}

export function createDatabaseClient(options: DatabaseClientOptions): DatabaseClient {
  return new InMemoryDatabaseClient(options)
}
