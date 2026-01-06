/**
 * Database.do Hierarchical Path Storage Tests
 *
 * TDD RED: Tests for path-based document storage and retrieval.
 * These tests should FAIL until implementation is complete.
 *
 * The hierarchical path storage allows data to be stored and retrieved
 * using paths like /path/to/data, similar to a filesystem or REST API.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import {
  PathStore,
  DatabaseClient,
  createDatabaseClient,
  parsePath,
  normalizePath,
  getParentPath,
  getPathSegments,
  joinPath,
  isValidPath,
  PathDocument,
  PathQuery,
  PathNotFoundError,
  InvalidPathError,
} from './database'

// ============================================================================
// Mock database.do Client
// ============================================================================

function createMockDatabaseClient(): DatabaseClient {
  const storage = new Map<string, PathDocument>()

  return {
    get: vi.fn(async (path: string): Promise<PathDocument | null> => {
      return storage.get(path) || null
    }),

    set: vi.fn(async (path: string, data: unknown, options?: { metadata?: Record<string, unknown> }): Promise<PathDocument> => {
      const now = new Date()
      const doc: PathDocument = {
        path,
        data,
        metadata: options?.metadata || {},
        createdAt: storage.get(path)?.createdAt || now,
        updatedAt: now,
      }
      storage.set(path, doc)
      return doc
    }),

    delete: vi.fn(async (path: string): Promise<boolean> => {
      return storage.delete(path)
    }),

    list: vi.fn(async (prefix: string, options?: { recursive?: boolean; limit?: number }): Promise<PathDocument[]> => {
      const results: PathDocument[] = []
      const prefixWithSlash = prefix.endsWith('/') ? prefix : prefix + '/'

      for (const [key, doc] of storage.entries()) {
        // Only include children under the prefix, not the prefix itself
        if (key.startsWith(prefixWithSlash)) {
          if (!options?.recursive) {
            // Only include direct children
            const relative = key.slice(prefixWithSlash.length)
            // Skip if there's a slash in the remaining path (nested child)
            if (relative.includes('/')) {
              continue
            }
          }
          results.push(doc)
        }
      }
      return options?.limit ? results.slice(0, options.limit) : results
    }),

    exists: vi.fn(async (path: string): Promise<boolean> => {
      return storage.has(path)
    }),

    close: vi.fn(async (): Promise<void> => {}),
  }
}

// ============================================================================
// Path Utility Function Tests
// ============================================================================

describe('Path Utility Functions', () => {
  describe('parsePath', () => {
    it('should parse a simple path', () => {
      const result = parsePath('/users')

      expect(result.segments).toEqual(['users'])
      expect(result.isAbsolute).toBe(true)
    })

    it('should parse a multi-level path', () => {
      const result = parsePath('/users/123/profile')

      expect(result.segments).toEqual(['users', '123', 'profile'])
      expect(result.isAbsolute).toBe(true)
    })

    it('should parse root path', () => {
      const result = parsePath('/')

      expect(result.segments).toEqual([])
      expect(result.isAbsolute).toBe(true)
    })

    it('should handle trailing slashes', () => {
      const result = parsePath('/users/123/')

      expect(result.segments).toEqual(['users', '123'])
    })

    it('should handle double slashes', () => {
      const result = parsePath('/users//profile')

      expect(result.segments).toEqual(['users', 'profile'])
    })

    it('should identify relative paths', () => {
      const result = parsePath('users/123')

      expect(result.isAbsolute).toBe(false)
      expect(result.segments).toEqual(['users', '123'])
    })
  })

  describe('normalizePath', () => {
    it('should normalize a simple path', () => {
      expect(normalizePath('/users')).toBe('/users')
    })

    it('should remove trailing slashes', () => {
      expect(normalizePath('/users/')).toBe('/users')
    })

    it('should collapse double slashes', () => {
      expect(normalizePath('/users//profile')).toBe('/users/profile')
    })

    it('should handle root path', () => {
      expect(normalizePath('/')).toBe('/')
    })

    it('should resolve relative paths with dots', () => {
      expect(normalizePath('/users/./profile')).toBe('/users/profile')
    })

    it('should resolve parent directory references', () => {
      expect(normalizePath('/users/123/../profile')).toBe('/users/profile')
    })

    it('should handle complex mixed paths', () => {
      expect(normalizePath('/users/./123/../456/profile/')).toBe('/users/456/profile')
    })

    it('should not go above root', () => {
      expect(normalizePath('/users/../..')).toBe('/')
    })
  })

  describe('getParentPath', () => {
    it('should return parent of a simple path', () => {
      expect(getParentPath('/users')).toBe('/')
    })

    it('should return parent of a nested path', () => {
      expect(getParentPath('/users/123/profile')).toBe('/users/123')
    })

    it('should return null for root path', () => {
      expect(getParentPath('/')).toBeNull()
    })

    it('should handle trailing slashes', () => {
      expect(getParentPath('/users/123/')).toBe('/users')
    })
  })

  describe('getPathSegments', () => {
    it('should return segments of a path', () => {
      expect(getPathSegments('/users/123/profile')).toEqual(['users', '123', 'profile'])
    })

    it('should return empty array for root', () => {
      expect(getPathSegments('/')).toEqual([])
    })

    it('should handle single segment', () => {
      expect(getPathSegments('/users')).toEqual(['users'])
    })
  })

  describe('joinPath', () => {
    it('should join two paths', () => {
      expect(joinPath('/users', 'profile')).toBe('/users/profile')
    })

    it('should join multiple segments', () => {
      expect(joinPath('/users', '123', 'profile')).toBe('/users/123/profile')
    })

    it('should handle absolute second path', () => {
      expect(joinPath('/users', '/profile')).toBe('/users/profile')
    })

    it('should join with root', () => {
      expect(joinPath('/', 'users')).toBe('/users')
    })

    it('should handle empty segments', () => {
      expect(joinPath('/users', '', 'profile')).toBe('/users/profile')
    })
  })

  describe('isValidPath', () => {
    it('should accept valid absolute paths', () => {
      expect(isValidPath('/users')).toBe(true)
      expect(isValidPath('/users/123/profile')).toBe(true)
      expect(isValidPath('/')).toBe(true)
    })

    it('should reject paths with invalid characters', () => {
      expect(isValidPath('/users?query=1')).toBe(false)
      expect(isValidPath('/users#hash')).toBe(false)
      expect(isValidPath('/users<script>')).toBe(false)
    })

    it('should reject paths with null bytes', () => {
      expect(isValidPath('/users\x00')).toBe(false)
    })

    it('should reject paths exceeding max length', () => {
      const longPath = '/' + 'a'.repeat(1025)
      expect(isValidPath(longPath)).toBe(false)
    })

    it('should reject relative paths when strict mode is enabled', () => {
      expect(isValidPath('users/123', { strict: true })).toBe(false)
    })

    it('should reject paths with too many segments', () => {
      const deepPath = '/' + Array(65).fill('segment').join('/')
      expect(isValidPath(deepPath)).toBe(false)
    })
  })
})

// ============================================================================
// PathStore Tests
// ============================================================================

describe('PathStore', () => {
  let mockClient: DatabaseClient
  let pathStore: PathStore

  beforeEach(() => {
    mockClient = createMockDatabaseClient()
    pathStore = new PathStore(mockClient)
  })

  afterEach(async () => {
    await pathStore.close()
  })

  describe('get', () => {
    it('should return null for non-existent path', async () => {
      const result = await pathStore.get('/nonexistent')

      expect(result).toBeNull()
    })

    it('should return data for existing path', async () => {
      await pathStore.set('/users/123', { name: 'John', email: 'john@example.com' })

      const result = await pathStore.get('/users/123')

      expect(result).not.toBeNull()
      expect(result?.data).toEqual({ name: 'John', email: 'john@example.com' })
    })

    it('should return path document with metadata', async () => {
      await pathStore.set('/users/123', { name: 'John' }, { metadata: { version: 1 } })

      const result = await pathStore.get('/users/123')

      expect(result?.metadata).toEqual({ version: 1 })
    })

    it('should normalize path before lookup', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      const result = await pathStore.get('/users/123/')

      expect(result).not.toBeNull()
      expect(result?.data).toEqual({ name: 'John' })
    })

    it('should handle root path', async () => {
      await pathStore.set('/', { root: true })

      const result = await pathStore.get('/')

      expect(result?.data).toEqual({ root: true })
    })

    it('should throw InvalidPathError for invalid paths', async () => {
      await expect(pathStore.get('/invalid?query')).rejects.toThrow(InvalidPathError)
    })
  })

  describe('set', () => {
    it('should store data at a path', async () => {
      const result = await pathStore.set('/users/123', { name: 'John' })

      expect(result.path).toBe('/users/123')
      expect(result.data).toEqual({ name: 'John' })
      expect(result.createdAt).toBeInstanceOf(Date)
      expect(result.updatedAt).toBeInstanceOf(Date)
    })

    it('should store primitive values', async () => {
      const result = await pathStore.set('/config/timeout', 5000)

      expect(result.data).toBe(5000)
    })

    it('should store arrays', async () => {
      const result = await pathStore.set('/users/123/roles', ['admin', 'user'])

      expect(result.data).toEqual(['admin', 'user'])
    })

    it('should store nested objects', async () => {
      const data = {
        user: {
          profile: {
            name: 'John',
            settings: {
              theme: 'dark',
            },
          },
        },
      }

      const result = await pathStore.set('/data/complex', data)

      expect(result.data).toEqual(data)
    })

    it('should store null values', async () => {
      const result = await pathStore.set('/users/123/deletedAt', null)

      expect(result.data).toBeNull()
    })

    it('should update existing data', async () => {
      await pathStore.set('/users/123', { name: 'John' })
      const result = await pathStore.set('/users/123', { name: 'Jane' })

      expect(result.data).toEqual({ name: 'Jane' })

      const fetched = await pathStore.get('/users/123')
      expect(fetched?.data).toEqual({ name: 'Jane' })
    })

    it('should preserve createdAt on update', async () => {
      const first = await pathStore.set('/users/123', { name: 'John' })
      await new Promise((resolve) => setTimeout(resolve, 10))
      const second = await pathStore.set('/users/123', { name: 'Jane' })

      expect(second.createdAt.getTime()).toBe(first.createdAt.getTime())
      expect(second.updatedAt.getTime()).toBeGreaterThan(first.createdAt.getTime())
    })

    it('should store with custom metadata', async () => {
      const result = await pathStore.set('/users/123', { name: 'John' }, {
        metadata: {
          version: 2,
          author: 'system',
          tags: ['important'],
        },
      })

      expect(result.metadata).toEqual({
        version: 2,
        author: 'system',
        tags: ['important'],
      })
    })

    it('should throw InvalidPathError for invalid paths', async () => {
      await expect(pathStore.set('/invalid<path>', {})).rejects.toThrow(InvalidPathError)
    })

    it('should create parent paths implicitly', async () => {
      await pathStore.set('/users/123/profile/settings', { theme: 'dark' })

      // Parent paths should exist as containers
      const parent = await pathStore.get('/users/123/profile')
      expect(parent).not.toBeNull()
    })
  })

  describe('delete', () => {
    it('should delete existing path and return true', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      const result = await pathStore.delete('/users/123')

      expect(result).toBe(true)

      const fetched = await pathStore.get('/users/123')
      expect(fetched).toBeNull()
    })

    it('should return false for non-existent path', async () => {
      const result = await pathStore.delete('/nonexistent')

      expect(result).toBe(false)
    })

    it('should delete recursively when option is set', async () => {
      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/users/123/profile', { bio: 'Developer' })
      await pathStore.set('/users/123/settings', { theme: 'dark' })

      const count = await pathStore.delete('/users/123', { recursive: true })

      expect(count).toBe(3)

      expect(await pathStore.get('/users/123')).toBeNull()
      expect(await pathStore.get('/users/123/profile')).toBeNull()
      expect(await pathStore.get('/users/123/settings')).toBeNull()
    })

    it('should not delete children by default', async () => {
      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/users/123/profile', { bio: 'Developer' })

      await pathStore.delete('/users/123')

      // Child should still exist
      const child = await pathStore.get('/users/123/profile')
      expect(child).not.toBeNull()
    })

    it('should throw InvalidPathError for invalid paths', async () => {
      await expect(pathStore.delete('/invalid#path')).rejects.toThrow(InvalidPathError)
    })
  })

  describe('list', () => {
    beforeEach(async () => {
      await pathStore.set('/users/1', { name: 'Alice' })
      await pathStore.set('/users/2', { name: 'Bob' })
      await pathStore.set('/users/3', { name: 'Charlie' })
      await pathStore.set('/users/1/profile', { bio: 'Dev' })
      await pathStore.set('/users/1/settings', { theme: 'light' })
      await pathStore.set('/posts/1', { title: 'Hello' })
    })

    it('should list direct children of a path', async () => {
      const results = await pathStore.list('/users')

      expect(results).toHaveLength(3)
      expect(results.map((r) => r.path)).toEqual(
        expect.arrayContaining(['/users/1', '/users/2', '/users/3'])
      )
    })

    it('should list recursively when option is set', async () => {
      const results = await pathStore.list('/users', { recursive: true })

      expect(results).toHaveLength(5) // 3 users + 2 children of user 1
    })

    it('should return empty array for non-existent prefix', async () => {
      const results = await pathStore.list('/nonexistent')

      expect(results).toHaveLength(0)
    })

    it('should respect limit option', async () => {
      const results = await pathStore.list('/users', { limit: 2 })

      expect(results).toHaveLength(2)
    })

    it('should respect offset option', async () => {
      const results = await pathStore.list('/users', { offset: 1, limit: 2 })

      expect(results).toHaveLength(2)
    })

    it('should sort results by path', async () => {
      const results = await pathStore.list('/users')

      const paths = results.map((r) => r.path)
      expect(paths).toEqual([...paths].sort())
    })

    it('should filter by metadata', async () => {
      await pathStore.set('/items/1', { name: 'A' }, { metadata: { type: 'active' } })
      await pathStore.set('/items/2', { name: 'B' }, { metadata: { type: 'archived' } })
      await pathStore.set('/items/3', { name: 'C' }, { metadata: { type: 'active' } })

      const results = await pathStore.list('/items', {
        filter: { metadata: { type: 'active' } },
      })

      expect(results).toHaveLength(2)
    })
  })

  describe('exists', () => {
    it('should return true for existing path', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      const result = await pathStore.exists('/users/123')

      expect(result).toBe(true)
    })

    it('should return false for non-existent path', async () => {
      const result = await pathStore.exists('/nonexistent')

      expect(result).toBe(false)
    })

    it('should normalize path before check', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      expect(await pathStore.exists('/users/123/')).toBe(true)
      expect(await pathStore.exists('/users//123')).toBe(true)
    })
  })

  describe('count', () => {
    beforeEach(async () => {
      await pathStore.set('/users/1', { name: 'Alice' })
      await pathStore.set('/users/2', { name: 'Bob' })
      await pathStore.set('/users/3', { name: 'Charlie' })
      await pathStore.set('/users/1/profile', { bio: 'Dev' })
    })

    it('should count direct children', async () => {
      const count = await pathStore.count('/users')

      expect(count).toBe(3)
    })

    it('should count recursively', async () => {
      const count = await pathStore.count('/users', { recursive: true })

      expect(count).toBe(4)
    })

    it('should return 0 for non-existent prefix', async () => {
      const count = await pathStore.count('/nonexistent')

      expect(count).toBe(0)
    })
  })

  describe('move', () => {
    it('should move a path to a new location', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      await pathStore.move('/users/123', '/archived/123')

      expect(await pathStore.exists('/users/123')).toBe(false)
      expect(await pathStore.exists('/archived/123')).toBe(true)

      const moved = await pathStore.get('/archived/123')
      expect(moved?.data).toEqual({ name: 'John' })
    })

    it('should move children when recursive', async () => {
      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/users/123/profile', { bio: 'Dev' })
      await pathStore.set('/users/123/settings', { theme: 'dark' })

      await pathStore.move('/users/123', '/archived/123', { recursive: true })

      expect(await pathStore.exists('/users/123')).toBe(false)
      expect(await pathStore.exists('/users/123/profile')).toBe(false)
      expect(await pathStore.exists('/archived/123')).toBe(true)
      expect(await pathStore.exists('/archived/123/profile')).toBe(true)
    })

    it('should throw PathNotFoundError if source does not exist', async () => {
      await expect(pathStore.move('/nonexistent', '/target')).rejects.toThrow(PathNotFoundError)
    })

    it('should overwrite destination if it exists', async () => {
      await pathStore.set('/source', { a: 1 })
      await pathStore.set('/target', { b: 2 })

      await pathStore.move('/source', '/target')

      const result = await pathStore.get('/target')
      expect(result?.data).toEqual({ a: 1 })
    })

    it('should preserve metadata during move', async () => {
      await pathStore.set('/source', { name: 'Test' }, { metadata: { version: 5 } })

      await pathStore.move('/source', '/target')

      const result = await pathStore.get('/target')
      expect(result?.metadata).toEqual({ version: 5 })
    })
  })

  describe('copy', () => {
    it('should copy a path to a new location', async () => {
      await pathStore.set('/users/123', { name: 'John' })

      await pathStore.copy('/users/123', '/backup/123')

      expect(await pathStore.exists('/users/123')).toBe(true)
      expect(await pathStore.exists('/backup/123')).toBe(true)

      const original = await pathStore.get('/users/123')
      const copied = await pathStore.get('/backup/123')
      expect(original?.data).toEqual(copied?.data)
    })

    it('should copy children when recursive', async () => {
      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/users/123/profile', { bio: 'Dev' })

      await pathStore.copy('/users/123', '/backup/123', { recursive: true })

      expect(await pathStore.exists('/users/123')).toBe(true)
      expect(await pathStore.exists('/users/123/profile')).toBe(true)
      expect(await pathStore.exists('/backup/123')).toBe(true)
      expect(await pathStore.exists('/backup/123/profile')).toBe(true)
    })

    it('should throw PathNotFoundError if source does not exist', async () => {
      await expect(pathStore.copy('/nonexistent', '/target')).rejects.toThrow(PathNotFoundError)
    })

    it('should generate new timestamps for copied documents', async () => {
      await pathStore.set('/source', { name: 'Test' })
      const original = await pathStore.get('/source')

      await new Promise((resolve) => setTimeout(resolve, 10))

      await pathStore.copy('/source', '/target')
      const copied = await pathStore.get('/target')

      expect(copied?.createdAt.getTime()).toBeGreaterThan(original!.createdAt.getTime())
    })
  })

  describe('query', () => {
    beforeEach(async () => {
      await pathStore.set('/users/1', { name: 'Alice', age: 30, active: true })
      await pathStore.set('/users/2', { name: 'Bob', age: 25, active: false })
      await pathStore.set('/users/3', { name: 'Charlie', age: 35, active: true })
      await pathStore.set('/users/4', { name: 'Diana', age: 28, active: true })
    })

    it('should query by data field equality', async () => {
      const query: PathQuery = {
        prefix: '/users',
        where: { active: true },
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(3)
      expect(results.every((r) => (r.data as { active: boolean }).active)).toBe(true)
    })

    it('should query with comparison operators', async () => {
      const query: PathQuery = {
        prefix: '/users',
        where: { age: { $gte: 30 } },
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(2)
      expect(results.every((r) => (r.data as { age: number }).age >= 30)).toBe(true)
    })

    it('should query with multiple conditions', async () => {
      const query: PathQuery = {
        prefix: '/users',
        where: {
          active: true,
          age: { $lt: 35 },
        },
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(2)
    })

    it('should query with $or operator', async () => {
      const query: PathQuery = {
        prefix: '/users',
        where: {
          $or: [{ name: 'Alice' }, { name: 'Bob' }],
        },
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(2)
    })

    it('should query with ordering', async () => {
      const query: PathQuery = {
        prefix: '/users',
        orderBy: { field: 'age', direction: 'desc' },
      }

      const results = await pathStore.query(query)

      const ages = results.map((r) => (r.data as { age: number }).age)
      expect(ages).toEqual([35, 30, 28, 25])
    })

    it('should query with pagination', async () => {
      const query: PathQuery = {
        prefix: '/users',
        orderBy: { field: 'name', direction: 'asc' },
        limit: 2,
        offset: 1,
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(2)
      expect((results[0].data as { name: string }).name).toBe('Bob')
    })

    it('should query with projection', async () => {
      const query: PathQuery = {
        prefix: '/users',
        select: ['name', 'age'],
      }

      const results = await pathStore.query(query)

      expect(results).toHaveLength(4)
      expect(results[0].data).toHaveProperty('name')
      expect(results[0].data).toHaveProperty('age')
      expect(results[0].data).not.toHaveProperty('active')
    })
  })

  describe('watch', () => {
    it('should emit events on data changes', async () => {
      const events: Array<{ type: string; path: string }> = []

      const unwatch = pathStore.watch('/users', (event) => {
        events.push({ type: event.type, path: event.path })
      })

      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/users/123', { name: 'Jane' })
      await pathStore.delete('/users/123')

      // Give time for events to propagate
      await new Promise((resolve) => setTimeout(resolve, 50))

      expect(events).toHaveLength(3)
      expect(events[0]).toEqual({ type: 'create', path: '/users/123' })
      expect(events[1]).toEqual({ type: 'update', path: '/users/123' })
      expect(events[2]).toEqual({ type: 'delete', path: '/users/123' })

      unwatch()
    })

    it('should only emit events for matching prefix', async () => {
      const events: string[] = []

      const unwatch = pathStore.watch('/users', (event) => {
        events.push(event.path)
      })

      await pathStore.set('/users/123', { name: 'John' })
      await pathStore.set('/posts/456', { title: 'Hello' })

      await new Promise((resolve) => setTimeout(resolve, 50))

      expect(events).toHaveLength(1)
      expect(events[0]).toBe('/users/123')

      unwatch()
    })

    it('should include data in events', async () => {
      let receivedData: unknown = null

      const unwatch = pathStore.watch('/users', (event) => {
        receivedData = event.data
      })

      await pathStore.set('/users/123', { name: 'John' })

      await new Promise((resolve) => setTimeout(resolve, 50))

      expect(receivedData).toEqual({ name: 'John' })

      unwatch()
    })

    it('should stop watching after unwatch is called', async () => {
      const events: string[] = []

      const unwatch = pathStore.watch('/users', (event) => {
        events.push(event.path)
      })

      await pathStore.set('/users/1', {})
      unwatch()
      await pathStore.set('/users/2', {})

      await new Promise((resolve) => setTimeout(resolve, 50))

      expect(events).toHaveLength(1)
    })
  })

  describe('transaction', () => {
    it('should execute multiple operations atomically', async () => {
      await pathStore.transaction(async (tx) => {
        await tx.set('/users/1', { name: 'Alice' })
        await tx.set('/users/2', { name: 'Bob' })
        await tx.set('/stats/userCount', 2)
      })

      expect(await pathStore.get('/users/1')).not.toBeNull()
      expect(await pathStore.get('/users/2')).not.toBeNull()
      expect(await pathStore.get('/stats/userCount')).not.toBeNull()
    })

    it('should rollback all operations on error', async () => {
      await pathStore.set('/users/1', { name: 'Original' })

      try {
        await pathStore.transaction(async (tx) => {
          await tx.set('/users/1', { name: 'Modified' })
          await tx.set('/users/2', { name: 'New' })
          throw new Error('Simulated error')
        })
      } catch {
        // Expected
      }

      const user1 = await pathStore.get('/users/1')
      expect(user1?.data).toEqual({ name: 'Original' })

      const user2 = await pathStore.get('/users/2')
      expect(user2).toBeNull()
    })

    it('should return value from transaction', async () => {
      const result = await pathStore.transaction(async (tx) => {
        await tx.set('/counter', 1)
        return 'success'
      })

      expect(result).toBe('success')
    })

    it('should allow reads within transaction', async () => {
      await pathStore.set('/users/1', { balance: 100 })

      await pathStore.transaction(async (tx) => {
        const user = await tx.get('/users/1')
        const newBalance = (user?.data as { balance: number }).balance + 50
        await tx.set('/users/1', { balance: newBalance })
      })

      const result = await pathStore.get('/users/1')
      expect((result?.data as { balance: number }).balance).toBe(150)
    })
  })
})

// ============================================================================
// DatabaseClient Tests
// ============================================================================

describe('DatabaseClient', () => {
  describe('createDatabaseClient', () => {
    it('should create a client with default options', () => {
      const client = createDatabaseClient({
        url: 'https://database.do',
      })

      expect(client).toBeDefined()
      expect(client.get).toBeDefined()
      expect(client.set).toBeDefined()
      expect(client.delete).toBeDefined()
      expect(client.list).toBeDefined()
    })

    it('should accept custom namespace', () => {
      const client = createDatabaseClient({
        url: 'https://database.do',
        namespace: 'my-app',
      })

      expect(client).toBeDefined()
    })

    it('should accept authentication token', () => {
      const client = createDatabaseClient({
        url: 'https://database.do',
        token: 'my-secret-token',
      })

      expect(client).toBeDefined()
    })
  })
})

// ============================================================================
// Error Handling Tests
// ============================================================================

describe('Error Handling', () => {
  let mockClient: DatabaseClient
  let pathStore: PathStore

  beforeEach(() => {
    mockClient = createMockDatabaseClient()
    pathStore = new PathStore(mockClient)
  })

  describe('PathNotFoundError', () => {
    it('should be thrown when trying to move non-existent path', async () => {
      try {
        await pathStore.move('/nonexistent', '/target')
        expect.fail('Should have thrown')
      } catch (error) {
        expect(error).toBeInstanceOf(PathNotFoundError)
        expect((error as PathNotFoundError).path).toBe('/nonexistent')
      }
    })

    it('should be thrown when trying to copy non-existent path', async () => {
      try {
        await pathStore.copy('/nonexistent', '/target')
        expect.fail('Should have thrown')
      } catch (error) {
        expect(error).toBeInstanceOf(PathNotFoundError)
      }
    })
  })

  describe('InvalidPathError', () => {
    it('should be thrown for paths with invalid characters', async () => {
      try {
        await pathStore.get('/invalid?path')
        expect.fail('Should have thrown')
      } catch (error) {
        expect(error).toBeInstanceOf(InvalidPathError)
        expect((error as InvalidPathError).message).toContain('invalid')
      }
    })

    it('should include the invalid path in the error', async () => {
      try {
        await pathStore.set('/bad<path>', {})
        expect.fail('Should have thrown')
      } catch (error) {
        expect(error).toBeInstanceOf(InvalidPathError)
        expect((error as InvalidPathError).path).toBe('/bad<path>')
      }
    })
  })
})

// ============================================================================
// Integration Patterns Tests
// ============================================================================

describe('Integration Patterns', () => {
  let mockClient: DatabaseClient
  let pathStore: PathStore

  beforeEach(() => {
    mockClient = createMockDatabaseClient()
    pathStore = new PathStore(mockClient)
  })

  describe('Hierarchical Data', () => {
    it('should support document hierarchy', async () => {
      // Create a project with nested structure
      await pathStore.set('/projects/web-app', {
        name: 'Web Application',
        status: 'active',
      })

      await pathStore.set('/projects/web-app/tasks/1', {
        title: 'Setup repository',
        completed: true,
      })

      await pathStore.set('/projects/web-app/tasks/2', {
        title: 'Design database',
        completed: false,
      })

      await pathStore.set('/projects/web-app/team/alice', {
        role: 'developer',
      })

      await pathStore.set('/projects/web-app/team/bob', {
        role: 'designer',
      })

      // List tasks
      const tasks = await pathStore.list('/projects/web-app/tasks')
      expect(tasks).toHaveLength(2)

      // List team members
      const team = await pathStore.list('/projects/web-app/team')
      expect(team).toHaveLength(2)
    })

    it('should support multi-tenant data isolation', async () => {
      await pathStore.set('/tenants/acme/users/1', { name: 'Alice' })
      await pathStore.set('/tenants/acme/users/2', { name: 'Bob' })
      await pathStore.set('/tenants/globex/users/1', { name: 'Charlie' })

      const acmeUsers = await pathStore.list('/tenants/acme/users')
      const globexUsers = await pathStore.list('/tenants/globex/users')

      expect(acmeUsers).toHaveLength(2)
      expect(globexUsers).toHaveLength(1)
    })
  })

  describe('Versioned Documents', () => {
    it('should support storing versions at paths', async () => {
      const docId = 'doc-123'

      await pathStore.set(`/docs/${docId}/versions/1`, {
        content: 'First version',
        author: 'alice',
      })

      await pathStore.set(`/docs/${docId}/versions/2`, {
        content: 'Second version',
        author: 'bob',
      })

      await pathStore.set(`/docs/${docId}/current`, {
        version: 2,
      })

      const versions = await pathStore.list(`/docs/${docId}/versions`)
      expect(versions).toHaveLength(2)

      const current = await pathStore.get(`/docs/${docId}/current`)
      expect(current?.data).toEqual({ version: 2 })
    })
  })

  describe('Configuration Trees', () => {
    it('should support hierarchical configuration', async () => {
      // Global defaults
      await pathStore.set('/config/defaults', {
        theme: 'light',
        language: 'en',
        notifications: true,
      })

      // User overrides
      await pathStore.set('/config/users/123', {
        theme: 'dark',
      })

      const defaults = await pathStore.get('/config/defaults')
      const userConfig = await pathStore.get('/config/users/123')

      // Merge configs (application would do this)
      const merged = {
        ...(defaults?.data as object),
        ...(userConfig?.data as object),
      }

      expect(merged).toEqual({
        theme: 'dark', // overridden
        language: 'en',
        notifications: true,
      })
    })
  })
})
