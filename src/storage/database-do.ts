/**
 * Database Durable Object Implementation
 *
 * Firebase-style realtime database with child event differentiation:
 * - onChildAdded: Emitted when a new document is added to a collection
 * - onChildChanged: Emitted when an existing document is modified
 * - onChildRemoved: Emitted when a document is deleted from a collection
 * - onChildMoved: Emitted when a document's order/position changes
 */

// ============================================================================
// Types
// ============================================================================

/** Event types for child event differentiation */
export type ChildEventType = 'child_added' | 'child_changed' | 'child_removed' | 'child_moved'

/** Base child event payload */
export interface ChildEvent<T = unknown> {
  type: ChildEventType
  collection: string
  key: string
  data: T | null
  previousData?: T | null
  previousKey?: string | null
  newPreviousKey?: string | null
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
export interface DatabaseDOInterface extends ChildEventEmitter {
  insert<T>(collection: string, key: string, data: T): Promise<T>
  update<T>(collection: string, key: string, data: Partial<T>): Promise<T>
  delete(collection: string, key: string): Promise<boolean>
  move<T>(collection: string, key: string, newPosition: number): Promise<T>
  get<T>(collection: string, key: string): Promise<T | null>
  list<T>(collection: string): Promise<T[]>
}

// ============================================================================
// Internal Types
// ============================================================================

interface ListenerEntry<T = unknown> {
  id: string
  callback: ChildEventCallback<T>
}

interface Document<T = unknown> {
  key: string
  data: T
  position: number
}

// ============================================================================
// DatabaseDO Implementation
// ============================================================================

export class DatabaseDO implements DatabaseDOInterface {
  private collections: Map<string, Map<string, Document>> = new Map()
  private listeners: Map<string, Map<ChildEventType, ListenerEntry[]>> = new Map()
  private listenerIdCounter = 0

  // ==========================================================================
  // Event Listener Registration
  // ==========================================================================

  onChildAdded<T>(collection: string, callback: ChildEventCallback<T>): () => void {
    return this.addListener(collection, 'child_added', callback as ChildEventCallback)
  }

  onChildChanged<T>(collection: string, callback: ChildEventCallback<T>): () => void {
    return this.addListener(collection, 'child_changed', callback as ChildEventCallback)
  }

  onChildRemoved<T>(collection: string, callback: ChildEventCallback<T>): () => void {
    return this.addListener(collection, 'child_removed', callback as ChildEventCallback)
  }

  onChildMoved<T>(collection: string, callback: ChildEventCallback<T>): () => void {
    return this.addListener(collection, 'child_moved', callback as ChildEventCallback)
  }

  off(collection: string, eventType?: ChildEventType): void {
    const collectionListeners = this.listeners.get(collection)
    if (!collectionListeners) return

    if (eventType) {
      collectionListeners.delete(eventType)
    } else {
      this.listeners.delete(collection)
    }
  }

  private addListener(collection: string, eventType: ChildEventType, callback: ChildEventCallback): () => void {
    if (!this.listeners.has(collection)) {
      this.listeners.set(collection, new Map())
    }

    const collectionListeners = this.listeners.get(collection)!
    if (!collectionListeners.has(eventType)) {
      collectionListeners.set(eventType, [])
    }

    const id = `listener_${++this.listenerIdCounter}`
    const entry: ListenerEntry = { id, callback }
    collectionListeners.get(eventType)!.push(entry)

    return () => {
      const listeners = collectionListeners.get(eventType)
      if (listeners) {
        const index = listeners.findIndex((l) => l.id === id)
        if (index >= 0) {
          listeners.splice(index, 1)
        }
      }
    }
  }

  private emitEvent(collection: string, eventType: ChildEventType, event: ChildEvent): void {
    const collectionListeners = this.listeners.get(collection)
    if (!collectionListeners) return

    const listeners = collectionListeners.get(eventType)
    if (!listeners) return

    for (const listener of listeners) {
      try {
        listener.callback(event)
      } catch {
        // Continue emitting to other listeners even if one throws
      }
    }
  }

  // ==========================================================================
  // CRUD Operations
  // ==========================================================================

  async insert<T>(collection: string, key: string, data: T): Promise<T> {
    if (!this.collections.has(collection)) {
      this.collections.set(collection, new Map())
    }

    const collectionData = this.collections.get(collection)!
    const position = collectionData.size

    const doc: Document<T> = { key, data, position }
    collectionData.set(key, doc as Document)

    this.emitEvent(collection, 'child_added', {
      type: 'child_added',
      collection,
      key,
      data,
      previousData: null,
      timestamp: Date.now(),
    })

    return data
  }

  async update<T>(collection: string, key: string, data: Partial<T>): Promise<T> {
    const collectionData = this.collections.get(collection)
    if (!collectionData) {
      throw new Error(`Collection not found: ${collection}`)
    }

    const doc = collectionData.get(key)
    if (!doc) {
      throw new Error(`Document not found: ${key}`)
    }

    const previousData = doc.data
    const newData = { ...(previousData as object), ...(data as object) } as T

    // Check if data actually changed
    const hasChanged = JSON.stringify(previousData) !== JSON.stringify(newData)

    doc.data = newData
    collectionData.set(key, doc)

    if (hasChanged) {
      this.emitEvent(collection, 'child_changed', {
        type: 'child_changed',
        collection,
        key,
        data: newData,
        previousData,
        timestamp: Date.now(),
      })
    }

    return newData
  }

  async delete(collection: string, key: string): Promise<boolean> {
    const collectionData = this.collections.get(collection)
    if (!collectionData) {
      return false
    }

    const doc = collectionData.get(key)
    if (!doc) {
      return false
    }

    const previousData = doc.data
    collectionData.delete(key)

    this.emitEvent(collection, 'child_removed', {
      type: 'child_removed',
      collection,
      key,
      data: null,
      previousData,
      timestamp: Date.now(),
    })

    return true
  }

  async move<T>(collection: string, key: string, newPosition: number): Promise<T> {
    const collectionData = this.collections.get(collection)
    if (!collectionData) {
      throw new Error(`Collection not found: ${collection}`)
    }

    const doc = collectionData.get(key)
    if (!doc) {
      throw new Error(`Document not found: ${key}`)
    }

    // Get ordered documents
    const orderedDocs = Array.from(collectionData.values()).sort((a, b) => a.position - b.position)
    const oldPosition = orderedDocs.findIndex((d) => d.key === key)

    // Clamp newPosition to valid range
    const maxPosition = orderedDocs.length - 1
    const clampedNewPosition = Math.max(0, Math.min(newPosition, maxPosition))

    // No change if moving to same position
    if (oldPosition === clampedNewPosition) {
      return doc.data as T
    }

    // Find previous key at old position
    const previousKey = oldPosition > 0 ? orderedDocs[oldPosition - 1]?.key ?? null : null

    // Remove doc from old position and insert at new position
    const docToMove = orderedDocs.splice(oldPosition, 1)[0]
    orderedDocs.splice(clampedNewPosition, 0, docToMove)

    // Calculate new previous key after the move
    const newPreviousKey = clampedNewPosition > 0 ? orderedDocs[clampedNewPosition - 1]?.key ?? null : null

    // Reassign all positions based on new order
    orderedDocs.forEach((d, index) => {
      d.position = index
      collectionData.set(d.key, d)
    })

    this.emitEvent(collection, 'child_moved', {
      type: 'child_moved',
      collection,
      key,
      data: doc.data,
      previousKey,
      newPreviousKey,
      timestamp: Date.now(),
    })

    return doc.data as T
  }

  async get<T>(collection: string, key: string): Promise<T | null> {
    const collectionData = this.collections.get(collection)
    if (!collectionData) {
      return null
    }

    const doc = collectionData.get(key)
    return doc ? (doc.data as T) : null
  }

  async list<T>(collection: string): Promise<T[]> {
    const collectionData = this.collections.get(collection)
    if (!collectionData) {
      return []
    }

    return Array.from(collectionData.values())
      .sort((a, b) => a.position - b.position)
      .map((doc) => doc.data as T)
  }
}
