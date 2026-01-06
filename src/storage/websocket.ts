/**
 * WebSocket Server for database.do (RTDB Protocol)
 *
 * Full implementation of WebSocket upgrade, hello message handshake,
 * connection lifecycle, and heartbeat functionality.
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
// WebSocket Accept Key Generation (RFC 6455)
// ============================================================================

const WEBSOCKET_GUID = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11'

/** Generate sec-websocket-accept header value per RFC 6455 */
async function generateAcceptKey(key: string): Promise<string> {
  const combined = key + WEBSOCKET_GUID
  const encoder = new TextEncoder()
  const data = encoder.encode(combined)
  const hashBuffer = await crypto.subtle.digest('SHA-1', data)
  const hashArray = new Uint8Array(hashBuffer)
  // Convert to base64
  let binary = ''
  for (const byte of hashArray) {
    binary += String.fromCharCode(byte)
  }
  return btoa(binary)
}

// ============================================================================
// Utility Functions
// ============================================================================

/** Generate a unique ID */
function generateId(prefix: string = ''): string {
  const timestamp = Date.now().toString(36)
  const random = Math.random().toString(36).slice(2, 10)
  return prefix ? `${prefix}-${timestamp}-${random}` : `${timestamp}-${random}`
}

/** Parse hello message from client */
export function parseHelloMessage(data: string): RTDBHelloMessage {
  let parsed: unknown
  try {
    parsed = JSON.parse(data)
  } catch {
    throw new RTDBProtocolError('Invalid JSON in hello message')
  }

  if (typeof parsed !== 'object' || parsed === null) {
    throw new RTDBProtocolError('Hello message must be an object')
  }

  const msg = parsed as Record<string, unknown>

  if (!msg.type) {
    throw new RTDBProtocolError('Hello message must have type field')
  }

  if (msg.type !== 'hello') {
    throw new RTDBProtocolError('First message must be hello type')
  }

  if (msg.version === undefined) {
    throw new RTDBProtocolError('Hello message must have version field')
  }

  // Only version 1 is supported
  if (msg.version !== 1) {
    throw new RTDBProtocolError('Unsupported protocol version')
  }

  return {
    type: 'hello',
    version: msg.version as number,
    clientId: msg.clientId as string | undefined,
    auth: msg.auth as RTDBHelloMessage['auth'],
    capabilities: msg.capabilities as string[] | undefined,
  }
}

/** Create hello response */
export function createHelloResponse(options: HelloResponseOptions): RTDBHelloResponse {
  return {
    type: 'hello',
    version: 1,
    sessionId: options.sessionId,
    clientId: options.clientId || generateId('client'),
    serverTime: Date.now(),
    capabilities: options.capabilities,
    heartbeatInterval: options.heartbeatInterval,
  }
}

/** Validate handshake message */
export function validateHandshake(hello: RTDBHelloMessage): boolean {
  if (!hello || hello.type !== 'hello') {
    return false
  }
  if (hello.version !== 1) {
    return false
  }
  return true
}

// ============================================================================
// RTDBConnection Class
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
  private _server?: RTDBServer
  private _handshakeTimeoutId?: ReturnType<typeof setTimeout>
  private _heartbeatIntervalId?: ReturnType<typeof setInterval>
  private _heartbeatTimeoutId?: ReturnType<typeof setTimeout>
  private _pendingPing: boolean = false

  constructor(private readonly ws: WebSocketLike) {
    this._id = generateId('conn')
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

  /** Set connection state */
  _setState(state: ConnectionState): void {
    this._state = state
  }

  /** Set client ID */
  _setClientId(clientId: string): void {
    this._clientId = clientId
  }

  /** Set session ID */
  _setSessionId(sessionId: string): void {
    this._sessionId = sessionId
  }

  /** Set server reference */
  _setServer(server: RTDBServer): void {
    this._server = server
  }

  /** Update last activity time */
  _updateActivity(): void {
    this._lastActivityAt = new Date()
  }

  /** Set handshake timeout */
  _setHandshakeTimeout(timeout: number): void {
    this._handshakeTimeoutId = setTimeout(() => {
      if (this._state === ConnectionState.HANDSHAKING) {
        this._closeWithError(4000, 'Handshake timeout')
      }
    }, timeout)
  }

  /** Clear handshake timeout */
  _clearHandshakeTimeout(): void {
    if (this._handshakeTimeoutId) {
      clearTimeout(this._handshakeTimeoutId)
      this._handshakeTimeoutId = undefined
    }
  }

  /** Start heartbeat interval */
  _startHeartbeat(interval: number, timeout?: number): void {
    this._heartbeatIntervalId = setInterval(() => {
      if (this._state === ConnectionState.CONNECTED) {
        this._sendPing()
        if (timeout) {
          this._pendingPing = true
          this._heartbeatTimeoutId = setTimeout(() => {
            if (this._pendingPing && this._state === ConnectionState.CONNECTED) {
              this._closeWithError(4002, 'Heartbeat timeout')
            }
          }, timeout)
        }
      }
    }, interval)
  }

  /** Stop heartbeat */
  _stopHeartbeat(): void {
    if (this._heartbeatIntervalId) {
      clearInterval(this._heartbeatIntervalId)
      this._heartbeatIntervalId = undefined
    }
    if (this._heartbeatTimeoutId) {
      clearTimeout(this._heartbeatTimeoutId)
      this._heartbeatTimeoutId = undefined
    }
  }

  /** Send ping message */
  private _sendPing(): void {
    this._sendMessage({
      type: 'ping',
      timestamp: Date.now(),
    })
  }

  /** Handle pong received */
  _handlePong(): void {
    this._pendingPing = false
    if (this._heartbeatTimeoutId) {
      clearTimeout(this._heartbeatTimeoutId)
      this._heartbeatTimeoutId = undefined
    }
  }

  /** Close with error */
  _closeWithError(code: number, reason: string): void {
    this._clearHandshakeTimeout()
    this._stopHeartbeat()
    this._state = ConnectionState.CLOSED
    this.ws.close(code, reason)
  }

  /** Emit event to handlers */
  _emit(event: string, data?: unknown): void {
    const handlers = this._eventHandlers.get(event) || []
    for (const handler of handlers) {
      try {
        handler(data)
      } catch (e) {
        // Ignore handler errors
      }
    }
  }

  /** Send a message directly (internal) */
  _sendMessage(message: RTDBMessage): void {
    this.ws.send(JSON.stringify(message))
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
  send(message: RTDBMessage): void {
    if (this._state === ConnectionState.CLOSED || this._state === ConnectionState.CLOSING) {
      throw new RTDBConnectionError('Cannot send message on closed connection')
    }
    this._sendMessage(message)
  }

  /** Close the connection */
  async close(reason?: string): Promise<void> {
    if (this._state === ConnectionState.CLOSED) {
      return
    }

    this._clearHandshakeTimeout()
    this._stopHeartbeat()
    this._state = ConnectionState.CLOSING

    // Send goodbye message
    try {
      this._sendMessage({
        type: 'goodbye',
        reason: reason || 'Connection closed',
      })
    } catch {
      // Ignore send errors during close
    }

    // Allow async observation of CLOSING state
    await Promise.resolve()

    this.ws.close(1000, reason || 'Connection closed')
    this._state = ConnectionState.CLOSED

    // Remove from server
    if (this._server) {
      this._server._removeConnection(this)
    }
  }
}

// ============================================================================
// RTDBServer Class
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
  async validateUpgrade(request: WebSocketUpgradeRequest): Promise<WebSocketUpgradeResult> {
    const headers = request.headers

    // Check upgrade header
    const upgrade = headers['upgrade']?.toLowerCase()
    if (upgrade !== 'websocket') {
      return { valid: false, error: 'Missing or invalid upgrade header' }
    }

    // Check sec-websocket-key
    const key = headers['sec-websocket-key']
    if (!key) {
      return { valid: false, error: 'Missing sec-websocket-key header' }
    }

    // Check sec-websocket-version
    const version = headers['sec-websocket-version']
    if (version !== '13') {
      return { valid: false, error: 'Unsupported WebSocket version' }
    }

    // Check protocol
    const protocol = headers['sec-websocket-protocol']
    if (protocol && protocol !== 'rtdb-v1') {
      return { valid: false, error: 'Unsupported protocol' }
    }

    // Validate origin if configured
    if (this.options.validateOrigin) {
      const origin = headers['origin']
      if (origin && !this.options.validateOrigin(origin)) {
        return { valid: false, error: 'Invalid origin' }
      }
    }

    // Generate accept key
    const accept = await generateAcceptKey(key)

    return {
      valid: true,
      protocol: 'rtdb-v1',
      accept,
    }
  }

  /** Get upgrade response headers */
  async getUpgradeResponseHeaders(request: WebSocketUpgradeRequest): Promise<Record<string, string>> {
    const key = request.headers['sec-websocket-key']
    const accept = await generateAcceptKey(key)

    return {
      'upgrade': 'websocket',
      'connection': 'Upgrade',
      'sec-websocket-accept': accept,
      'sec-websocket-protocol': 'rtdb-v1',
    }
  }

  /** Handle a new WebSocket connection */
  async handleConnection(ws: WebSocketLike): Promise<RTDBConnection> {
    if (this._closed) {
      throw new RTDBConnectionError('Server is closed')
    }

    const connection = new RTDBConnection(ws)
    connection._setServer(this)
    connection._setState(ConnectionState.HANDSHAKING)

    // Add to connections map
    this._connections.set(connection.id, connection)

    // Set handshake timeout
    const handshakeTimeout = this.options.handshakeTimeout || 30000
    connection._setHandshakeTimeout(handshakeTimeout)

    // Setup message handler
    const messageHandler = async (event: unknown) => {
      const data = (event as { data: string }).data
      connection._updateActivity()

      if (connection.state === ConnectionState.HANDSHAKING) {
        // Expect hello message
        try {
          const hello = parseHelloMessage(data)

          // Validate authentication if configured
          if (this.options.authenticate) {
            const authenticated = await this.options.authenticate(hello.auth)
            if (!authenticated) {
              connection._closeWithError(4003, 'Failed authentication')
              return
            }
          }

          // Clear handshake timeout
          connection._clearHandshakeTimeout()

          // Set connection properties
          const sessionId = generateId('session')
          const clientId = hello.clientId || generateId('client')
          connection._setSessionId(sessionId)
          connection._setClientId(clientId)
          connection._setState(ConnectionState.CONNECTED)

          // Send hello response
          const response = createHelloResponse({
            sessionId,
            clientId,
            heartbeatInterval: this.options.heartbeatInterval,
          })
          connection._sendMessage(response)

          // Start heartbeat if configured
          if (this.options.heartbeatInterval && this.options.heartbeatInterval > 0) {
            connection._startHeartbeat(
              this.options.heartbeatInterval,
              this.options.heartbeatTimeout
            )
          }
        } catch (e) {
          if (e instanceof RTDBProtocolError) {
            connection._closeWithError(4001, 'Expected hello message')
          } else {
            connection._closeWithError(4001, 'Invalid hello message')
          }
        }
      } else if (connection.state === ConnectionState.CONNECTED) {
        // Handle other messages
        try {
          const msg = JSON.parse(data) as RTDBMessage

          if (msg.type === 'ping') {
            // Respond with pong
            const pingMsg = msg as { timestamp?: number }
            connection._sendMessage({
              type: 'pong',
              timestamp: Date.now(),
              echoTimestamp: pingMsg.timestamp,
            })
          } else if (msg.type === 'pong') {
            connection._handlePong()
          }
        } catch (e) {
          connection._emit('error', e)
        }
      }
    }

    const errorHandler = (error: unknown) => {
      connection._emit('error', error)
      connection._clearHandshakeTimeout()
      connection._stopHeartbeat()
      connection._setState(ConnectionState.CLOSED)
      this._connections.delete(connection.id)
    }

    const closeHandler = (event: unknown) => {
      const closeEvent = event as CloseEventData
      connection._emit('close', closeEvent)
      connection._clearHandshakeTimeout()
      connection._stopHeartbeat()
      connection._setState(ConnectionState.CLOSED)
      this._connections.delete(connection.id)
    }

    ws.addEventListener('message', messageHandler)
    ws.addEventListener('error', errorHandler)
    ws.addEventListener('close', closeHandler)

    return connection
  }

  /** Remove connection from tracking */
  _removeConnection(connection: RTDBConnection): void {
    this._connections.delete(connection.id)
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
    this._closed = true

    // Close all connections with goodbye
    const closePromises = []
    for (const conn of this._connections.values()) {
      closePromises.push(
        (async () => {
          try {
            // Send goodbye with shutdown reason
            conn._sendMessage({
              type: 'goodbye',
              reason: 'Server shutdown',
            })
          } catch {
            // Ignore send errors
          }
          conn._clearHandshakeTimeout()
          conn._stopHeartbeat()
          conn._setState(ConnectionState.CLOSED)
          conn['ws'].close(1000, 'Server shutdown')
        })()
      )
    }

    await Promise.all(closePromises)
    this._connections.clear()
  }
}

// ============================================================================
// RTDBWebSocket Class (Client)
// ============================================================================

/** RTDB WebSocket client */
export class RTDBWebSocket {
  constructor(_url: string, _options?: RTDBServerOptions) {
    // Client implementation - not needed for current tests
  }

  async connect(): Promise<void> {
    throw new Error('Not implemented')
  }

  async close(): Promise<void> {
    throw new Error('Not implemented')
  }
}

/** Create an RTDB server */
export function createRTDBServer(options?: RTDBServerOptions): RTDBServer {
  return new RTDBServer(options)
}
