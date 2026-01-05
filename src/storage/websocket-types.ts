/**
 * WebSocket Types for database.do (RTDB Protocol)
 *
 * Type definitions for real-time database WebSocket communication.
 */

// ============================================================================
// WebSocket Interface
// ============================================================================

/** WebSocket-like interface for testing and compatibility */
export interface WebSocketLike {
  readonly readyState: number
  send(data: string): void
  close(code?: number, reason?: string): void
  addEventListener(event: string, handler: (data?: unknown) => void): void
  removeEventListener(event: string, handler: (data?: unknown) => void): void
}

/** WebSocket upgrade request */
export interface WebSocketUpgradeRequest {
  url: string
  headers: Record<string, string>
  method: string
}

/** WebSocket upgrade validation result */
export interface WebSocketUpgradeResult {
  valid: boolean
  error?: string
  protocol?: string
  accept?: string
}

// ============================================================================
// RTDB Protocol Messages
// ============================================================================

/** Base message interface */
export interface RTDBMessage {
  type: string
  [key: string]: unknown
}

/** Hello message from client to server */
export interface RTDBHelloMessage extends RTDBMessage {
  type: 'hello'
  version: number
  clientId?: string
  auth?: {
    token?: string
    apiKey?: string
  }
  capabilities?: string[]
}

/** Hello response from server to client */
export interface RTDBHelloResponse extends RTDBMessage {
  type: 'hello'
  version: number
  sessionId: string
  clientId: string
  serverTime: number
  capabilities?: string[]
  heartbeatInterval?: number
}

/** Goodbye message */
export interface RTDBGoodbyeMessage extends RTDBMessage {
  type: 'goodbye'
  reason?: string
}

/** Ping message for heartbeat */
export interface RTDBPingMessage extends RTDBMessage {
  type: 'ping'
  timestamp: number
}

/** Pong response to ping */
export interface RTDBPongMessage extends RTDBMessage {
  type: 'pong'
  timestamp: number
  echoTimestamp: number
}

/** Error message */
export interface RTDBErrorMessage extends RTDBMessage {
  type: 'error'
  code: number
  message: string
  details?: unknown
}

// ============================================================================
// Connection Options
// ============================================================================

/** Options for RTDB connection */
export interface RTDBConnectionOptions {
  /** Connection timeout in milliseconds */
  timeout?: number
  /** Heartbeat interval in milliseconds */
  heartbeatInterval?: number
  /** Heartbeat timeout in milliseconds */
  heartbeatTimeout?: number
  /** Custom headers for upgrade request */
  headers?: Record<string, string>
  /** Authentication options */
  auth?: {
    token?: string
    apiKey?: string
  }
  /** Client capabilities */
  capabilities?: string[]
}

/** Options for RTDB server */
export interface RTDBServerOptions {
  /** Handshake timeout in milliseconds */
  handshakeTimeout?: number
  /** Heartbeat interval in milliseconds (0 to disable) */
  heartbeatInterval?: number
  /** Heartbeat timeout in milliseconds */
  heartbeatTimeout?: number
  /** Origin validation function */
  validateOrigin?: (origin: string) => boolean
  /** Authentication function */
  authenticate?: (auth?: RTDBHelloMessage['auth']) => Promise<boolean>
  /** Maximum connections (0 for unlimited) */
  maxConnections?: number
}

/** Options for hello response */
export interface HelloResponseOptions {
  sessionId: string
  clientId?: string
  capabilities?: string[]
  heartbeatInterval?: number
}

// ============================================================================
// Connection Events
// ============================================================================

/** Close event data */
export interface CloseEventData {
  code: number
  reason: string
}

/** Message event data */
export interface MessageEventData {
  data: string
}

/** Connection event types */
export type ConnectionEvent =
  | 'open'
  | 'close'
  | 'error'
  | 'message'
  | 'handshake'
  | 'connected'

/** Event handler type */
export type EventHandler<T = unknown> = (data: T) => void
