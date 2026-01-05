/**
 * WebSocket Server for database.do (RTDB Protocol)
 *
 * Stub implementation for TDD RED phase.
 * These exports exist so tests can import them, but implementations will fail.
 */

import type {
  WebSocketLike,
  WebSocketUpgradeRequest,
  WebSocketUpgradeResult,
  RTDBHelloMessage,
  RTDBHelloResponse,
  RTDBMessage,
  RTDBServerOptions,
  HelloResponseOptions,
  CloseEventData,
  EventHandler,
} from './websocket-types'

// Re-export types
export type {
  WebSocketLike,
  WebSocketUpgradeRequest,
  WebSocketUpgradeResult,
  RTDBHelloMessage,
  RTDBHelloResponse,
  RTDBMessage,
  RTDBServerOptions,
  HelloResponseOptions,
}

// ============================================================================
// Error Classes
// ============================================================================

/** Error during WebSocket handshake */
export class RTDBHandshakeError extends Error {
  name = 'RTDBHandshakeError'

  constructor(
    message: string,
    public readonly code?: number
  ) {
    super(message)
  }
}

/** Error in RTDB protocol */
export class RTDBProtocolError extends Error {
  name = 'RTDBProtocolError'

  constructor(message: string) {
    super(message)
  }
}

/** Error with RTDB connection */
export class RTDBConnectionError extends Error {
  name = 'RTDBConnectionError'

  constructor(
    message: string,
    public readonly cause?: Error
  ) {
    super(message)
  }
}

// ============================================================================
// Enums
// ============================================================================

/** Connection state enum */
export enum ConnectionState {
  CONNECTING = 'connecting',
  HANDSHAKING = 'handshaking',
  CONNECTED = 'connected',
  CLOSING = 'closing',
  CLOSED = 'closed',
}

/** Message type enum */
export enum MessageType {
  HELLO = 'hello',
  GOODBYE = 'goodbye',
  PING = 'ping',
  PONG = 'pong',
  ERROR = 'error',
  SUBSCRIBE = 'subscribe',
  UNSUBSCRIBE = 'unsubscribe',
  UPDATE = 'update',
  ACK = 'ack',
}

// ============================================================================
// RTDBConnection Class (Stub)
// ============================================================================

/** Represents a single WebSocket connection */
export class RTDBConnection {
  private _state: ConnectionState = ConnectionState.CONNECTING
  private _id: string
  private _clientId?: string
  private _sessionId?: string
  private _createdAt: Date
  private _lastActivityAt: Date
  private _eventHandlers: Map<string, EventHandler[]> = new Map()

  constructor(private readonly ws: WebSocketLike) {
    this._id = `conn-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    this._createdAt = new Date()
    this._lastActivityAt = new Date()
  }

  get id(): string {
    return this._id
  }

  get state(): ConnectionState {
    return this._state
  }

  get clientId(): string | undefined {
    return this._clientId
  }

  get sessionId(): string | undefined {
    return this._sessionId
  }

  get createdAt(): Date {
    return this._createdAt
  }

  get lastActivityAt(): Date {
    return this._lastActivityAt
  }

  /** Register event handler */
  on(event: string, handler: EventHandler): void {
    if (!this._eventHandlers.has(event)) {
      this._eventHandlers.set(event, [])
    }
    this._eventHandlers.get(event)!.push(handler)
  }

  /** Remove event handler */
  off(event: string, handler: EventHandler): void {
    const handlers = this._eventHandlers.get(event)
    if (handlers) {
      const idx = handlers.indexOf(handler)
      if (idx >= 0) handlers.splice(idx, 1)
    }
  }

  /** Send a message */
  send(_message: RTDBMessage): void {
    // Stub - will fail tests
    throw new RTDBConnectionError('Not implemented')
  }

  /** Close the connection */
  async close(_reason?: string): Promise<void> {
    // Stub - will fail tests
    throw new RTDBConnectionError('Not implemented')
  }
}

// ============================================================================
// RTDBServer Class (Stub)
// ============================================================================

/** RTDB WebSocket server */
export class RTDBServer {
  private _connections: Map<string, RTDBConnection> = new Map()
  private _closed = false

  constructor(private readonly options: RTDBServerOptions = {}) {}

  get activeConnections(): number {
    return this._connections.size
  }

  /** Validate WebSocket upgrade request */
  async validateUpgrade(_request: WebSocketUpgradeRequest): Promise<WebSocketUpgradeResult> {
    // Stub - will fail tests
    throw new Error('Not implemented')
  }

  /** Get upgrade response headers */
  async getUpgradeResponseHeaders(_request: WebSocketUpgradeRequest): Promise<Record<string, string>> {
    // Stub - will fail tests
    throw new Error('Not implemented')
  }

  /** Handle a new WebSocket connection */
  async handleConnection(_ws: WebSocketLike): Promise<RTDBConnection> {
    // Stub - will fail tests
    throw new RTDBConnectionError('Not implemented')
  }

  /** Get all connections */
  getConnections(): RTDBConnection[] {
    return Array.from(this._connections.values())
  }

  /** Get connection by id */
  getConnection(id: string): RTDBConnection | undefined {
    return this._connections.get(id)
  }

  /** Get connection by client id */
  getConnectionByClientId(clientId: string): RTDBConnection | undefined {
    for (const conn of this._connections.values()) {
      if (conn.clientId === clientId) {
        return conn
      }
    }
    return undefined
  }

  /** Close the server */
  async close(): Promise<void> {
    // Stub - mark as closed but don't do full implementation
    this._closed = true
  }
}

// ============================================================================
// RTDBWebSocket Class (Stub - Client)
// ============================================================================

/** RTDB WebSocket client */
export class RTDBWebSocket {
  constructor(_url: string, _options?: RTDBServerOptions) {
    // Stub
  }

  async connect(): Promise<void> {
    throw new Error('Not implemented')
  }

  async close(): Promise<void> {
    throw new Error('Not implemented')
  }
}

// ============================================================================
// Utility Functions (Stubs)
// ============================================================================

/** Create an RTDB server */
export function createRTDBServer(options?: RTDBServerOptions): RTDBServer {
  return new RTDBServer(options)
}

/** Parse hello message from client */
export function parseHelloMessage(_data: string): RTDBHelloMessage {
  // Stub - will fail tests
  throw new RTDBProtocolError('Not implemented')
}

/** Create hello response */
export function createHelloResponse(_options: HelloResponseOptions): RTDBHelloResponse {
  // Stub - will fail tests
  throw new Error('Not implemented')
}

/** Validate handshake message */
export function validateHandshake(_hello: RTDBHelloMessage): boolean {
  // Stub - will fail tests
  throw new Error('Not implemented')
}
