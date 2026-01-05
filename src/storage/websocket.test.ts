/**
 * WebSocket Connection Tests for database.do (RTDB Protocol)
 *
 * TDD RED: Tests for WebSocket upgrade, hello message, and connection lifecycle.
 * These tests should FAIL until implementation is complete.
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest'
import {
  RTDBWebSocket,
  RTDBServer,
  RTDBConnection,
  RTDBHandshakeError,
  RTDBProtocolError,
  RTDBConnectionError,
  createRTDBServer,
  parseHelloMessage,
  createHelloResponse,
  validateHandshake,
  ConnectionState,
  MessageType,
} from './websocket'
import type {
  RTDBHelloMessage,
  RTDBHelloResponse,
  RTDBMessage,
  RTDBConnectionOptions,
  RTDBServerOptions,
  WebSocketUpgradeRequest,
  WebSocketLike,
} from './websocket-types'

// ============================================================================
// Mock WebSocket Implementation
// ============================================================================

function createMockWebSocket(): WebSocketLike & {
  mockTrigger: (event: string, data?: unknown) => void
  sentMessages: string[]
} {
  const listeners: Map<string, Array<(data?: unknown) => void>> = new Map()
  const sentMessages: string[] = []

  return {
    readyState: 1, // OPEN
    sentMessages,

    send: vi.fn((data: string) => {
      sentMessages.push(data)
    }),

    close: vi.fn((code?: number, reason?: string) => {}),

    addEventListener: vi.fn((event: string, handler: (data?: unknown) => void) => {
      if (!listeners.has(event)) {
        listeners.set(event, [])
      }
      listeners.get(event)!.push(handler)
    }),

    removeEventListener: vi.fn((event: string, handler: (data?: unknown) => void) => {
      const handlers = listeners.get(event)
      if (handlers) {
        const idx = handlers.indexOf(handler)
        if (idx >= 0) handlers.splice(idx, 1)
      }
    }),

    mockTrigger: (event: string, data?: unknown) => {
      const handlers = listeners.get(event) || []
      for (const handler of handlers) {
        handler(data)
      }
    },
  }
}

function createMockUpgradeRequest(overrides?: Partial<WebSocketUpgradeRequest>): WebSocketUpgradeRequest {
  return {
    url: '/rtdb/connect',
    headers: {
      'upgrade': 'websocket',
      'connection': 'upgrade',
      'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
      'sec-websocket-version': '13',
      'sec-websocket-protocol': 'rtdb-v1',
      ...overrides?.headers,
    },
    method: 'GET',
    ...overrides,
  }
}

// ============================================================================
// WebSocket Upgrade Tests
// ============================================================================

describe('WebSocket Upgrade', () => {
  let server: RTDBServer

  beforeEach(() => {
    server = createRTDBServer({
      validateOrigin: (origin) => true,
    })
  })

  afterEach(async () => {
    await server.close()
  })

  describe('HTTP Upgrade Request Validation', () => {
    it('should accept valid WebSocket upgrade request', async () => {
      const request = createMockUpgradeRequest()

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(true)
      expect(result.protocol).toBe('rtdb-v1')
    })

    it('should reject request without upgrade header', async () => {
      const request = createMockUpgradeRequest({
        headers: {
          'connection': 'keep-alive',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '13',
        },
      })

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(false)
      expect(result.error).toContain('upgrade')
    })

    it('should reject request without sec-websocket-key', async () => {
      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-version': '13',
        },
      })

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(false)
      expect(result.error).toContain('sec-websocket-key')
    })

    it('should reject unsupported WebSocket version', async () => {
      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '8',
        },
      })

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(false)
      expect(result.error).toContain('version')
    })

    it('should require rtdb-v1 protocol', async () => {
      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '13',
          'sec-websocket-protocol': 'graphql-ws',
        },
      })

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(false)
      expect(result.error).toContain('protocol')
    })

    it('should validate origin when configured', async () => {
      const serverWithOriginCheck = createRTDBServer({
        validateOrigin: (origin) => origin === 'https://example.com',
      })

      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '13',
          'sec-websocket-protocol': 'rtdb-v1',
          'origin': 'https://malicious.com',
        },
      })

      const result = await serverWithOriginCheck.validateUpgrade(request)

      expect(result.valid).toBe(false)
      expect(result.error).toContain('origin')

      await serverWithOriginCheck.close()
    })

    it('should accept valid origin', async () => {
      const serverWithOriginCheck = createRTDBServer({
        validateOrigin: (origin) => origin === 'https://example.com',
      })

      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '13',
          'sec-websocket-protocol': 'rtdb-v1',
          'origin': 'https://example.com',
        },
      })

      const result = await serverWithOriginCheck.validateUpgrade(request)

      expect(result.valid).toBe(true)

      await serverWithOriginCheck.close()
    })

    it('should generate correct sec-websocket-accept header', async () => {
      const request = createMockUpgradeRequest({
        headers: {
          'upgrade': 'websocket',
          'connection': 'upgrade',
          'sec-websocket-key': 'dGhlIHNhbXBsZSBub25jZQ==',
          'sec-websocket-version': '13',
          'sec-websocket-protocol': 'rtdb-v1',
        },
      })

      const result = await server.validateUpgrade(request)

      expect(result.valid).toBe(true)
      // RFC 6455 specifies the exact algorithm for sec-websocket-accept
      expect(result.accept).toBe('s3pPLMBiTxaQ9kYGzzhZRbK+xOo=')
    })
  })

  describe('Upgrade Response Headers', () => {
    it('should return correct upgrade response headers', async () => {
      const request = createMockUpgradeRequest()

      const headers = await server.getUpgradeResponseHeaders(request)

      expect(headers['upgrade']).toBe('websocket')
      expect(headers['connection']).toBe('Upgrade')
      expect(headers['sec-websocket-accept']).toBeDefined()
      expect(headers['sec-websocket-protocol']).toBe('rtdb-v1')
    })
  })
})

// ============================================================================
// Hello Message Handshake Tests
// ============================================================================

describe('Hello Message Handshake', () => {
  let server: RTDBServer
  let mockWs: ReturnType<typeof createMockWebSocket>
  let connection: RTDBConnection

  beforeEach(() => {
    server = createRTDBServer()
    mockWs = createMockWebSocket()
  })

  afterEach(async () => {
    if (connection) {
      await connection.close()
    }
    await server.close()
  })

  describe('Client Hello Message', () => {
    it('should parse valid hello message', () => {
      const helloJson = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
        auth: {
          token: 'bearer-token-xyz',
        },
      })

      const hello = parseHelloMessage(helloJson)

      expect(hello.type).toBe('hello')
      expect(hello.version).toBe(1)
      expect(hello.clientId).toBe('client-123')
      expect(hello.auth?.token).toBe('bearer-token-xyz')
    })

    it('should reject non-JSON hello message', () => {
      expect(() => {
        parseHelloMessage('not json')
      }).toThrow(RTDBProtocolError)
    })

    it('should reject hello message without type', () => {
      const helloJson = JSON.stringify({
        version: 1,
        clientId: 'client-123',
      })

      expect(() => {
        parseHelloMessage(helloJson)
      }).toThrow(RTDBProtocolError)
    })

    it('should reject hello message with wrong type', () => {
      const helloJson = JSON.stringify({
        type: 'subscribe',
        version: 1,
        clientId: 'client-123',
      })

      expect(() => {
        parseHelloMessage(helloJson)
      }).toThrow(RTDBProtocolError)
    })

    it('should reject hello message without version', () => {
      const helloJson = JSON.stringify({
        type: 'hello',
        clientId: 'client-123',
      })

      expect(() => {
        parseHelloMessage(helloJson)
      }).toThrow(RTDBProtocolError)
    })

    it('should reject unsupported protocol version', () => {
      const helloJson = JSON.stringify({
        type: 'hello',
        version: 99,
        clientId: 'client-123',
      })

      expect(() => {
        parseHelloMessage(helloJson)
      }).toThrow(RTDBProtocolError)
    })

    it('should accept hello message without clientId (server generates one)', () => {
      const helloJson = JSON.stringify({
        type: 'hello',
        version: 1,
      })

      const hello = parseHelloMessage(helloJson)

      expect(hello.type).toBe('hello')
      expect(hello.version).toBe(1)
      expect(hello.clientId).toBeUndefined()
    })

    it('should parse hello message with capabilities', () => {
      const helloJson = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
        capabilities: ['compression', 'binary'],
      })

      const hello = parseHelloMessage(helloJson)

      expect(hello.capabilities).toContain('compression')
      expect(hello.capabilities).toContain('binary')
    })
  })

  describe('Server Hello Response', () => {
    it('should create hello response with session id', () => {
      const response = createHelloResponse({
        sessionId: 'session-456',
        clientId: 'client-123',
      })

      expect(response.type).toBe('hello')
      expect(response.version).toBe(1)
      expect(response.sessionId).toBe('session-456')
      expect(response.clientId).toBe('client-123')
    })

    it('should generate clientId if not provided by client', () => {
      const response = createHelloResponse({
        sessionId: 'session-456',
      })

      expect(response.clientId).toBeDefined()
      expect(response.clientId.length).toBeGreaterThan(0)
    })

    it('should include server timestamp', () => {
      const before = Date.now()
      const response = createHelloResponse({
        sessionId: 'session-456',
        clientId: 'client-123',
      })
      const after = Date.now()

      expect(response.serverTime).toBeGreaterThanOrEqual(before)
      expect(response.serverTime).toBeLessThanOrEqual(after)
    })

    it('should include negotiated capabilities', () => {
      const response = createHelloResponse({
        sessionId: 'session-456',
        clientId: 'client-123',
        capabilities: ['compression'],
      })

      expect(response.capabilities).toContain('compression')
    })

    it('should include heartbeat interval', () => {
      const response = createHelloResponse({
        sessionId: 'session-456',
        clientId: 'client-123',
        heartbeatInterval: 30000,
      })

      expect(response.heartbeatInterval).toBe(30000)
    })
  })

  describe('Handshake Flow', () => {
    it('should complete handshake when client sends valid hello', async () => {
      connection = await server.handleConnection(mockWs)

      expect(connection.state).toBe(ConnectionState.HANDSHAKING)

      // Client sends hello
      const helloMessage = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
      })

      mockWs.mockTrigger('message', { data: helloMessage })

      // Wait for async processing
      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CONNECTED)
      })

      expect(connection.clientId).toBe('client-123')
      expect(connection.sessionId).toBeDefined()
    })

    it('should send hello response after receiving client hello', async () => {
      connection = await server.handleConnection(mockWs)

      const helloMessage = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
      })

      mockWs.mockTrigger('message', { data: helloMessage })

      await vi.waitFor(() => {
        expect(mockWs.sentMessages.length).toBeGreaterThan(0)
      })

      const response = JSON.parse(mockWs.sentMessages[0])

      expect(response.type).toBe('hello')
      expect(response.sessionId).toBeDefined()
      expect(response.clientId).toBe('client-123')
    })

    it('should timeout if client does not send hello within timeout', async () => {
      const serverWithShortTimeout = createRTDBServer({
        handshakeTimeout: 100, // 100ms
      })

      connection = await serverWithShortTimeout.handleConnection(mockWs)

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CLOSED)
      }, { timeout: 500 })

      expect(mockWs.close).toHaveBeenCalledWith(
        4000,
        expect.stringContaining('timeout')
      )

      await serverWithShortTimeout.close()
    })

    it('should reject connection if first message is not hello', async () => {
      connection = await server.handleConnection(mockWs)

      const subscribeMessage = JSON.stringify({
        type: 'subscribe',
        path: '/cells/A1',
      })

      mockWs.mockTrigger('message', { data: subscribeMessage })

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CLOSED)
      })

      expect(mockWs.close).toHaveBeenCalledWith(
        4001,
        expect.stringContaining('hello')
      )
    })

    it('should validate authentication token during handshake', async () => {
      const serverWithAuth = createRTDBServer({
        authenticate: async (auth) => {
          return auth?.token === 'valid-token'
        },
      })

      connection = await serverWithAuth.handleConnection(mockWs)

      const helloMessage = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
        auth: {
          token: 'invalid-token',
        },
      })

      mockWs.mockTrigger('message', { data: helloMessage })

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CLOSED)
      })

      expect(mockWs.close).toHaveBeenCalledWith(
        4003,
        expect.stringContaining('authentication')
      )

      await serverWithAuth.close()
    })

    it('should accept valid authentication token', async () => {
      const serverWithAuth = createRTDBServer({
        authenticate: async (auth) => {
          return auth?.token === 'valid-token'
        },
      })

      connection = await serverWithAuth.handleConnection(mockWs)

      const helloMessage = JSON.stringify({
        type: 'hello',
        version: 1,
        clientId: 'client-123',
        auth: {
          token: 'valid-token',
        },
      })

      mockWs.mockTrigger('message', { data: helloMessage })

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CONNECTED)
      })

      await serverWithAuth.close()
    })
  })

  describe('validateHandshake', () => {
    it('should return true for valid hello message', () => {
      const hello: RTDBHelloMessage = {
        type: 'hello',
        version: 1,
        clientId: 'client-123',
      }

      expect(validateHandshake(hello)).toBe(true)
    })

    it('should return false for invalid version', () => {
      const hello = {
        type: 'hello',
        version: 0,
        clientId: 'client-123',
      } as RTDBHelloMessage

      expect(validateHandshake(hello)).toBe(false)
    })

    it('should return false for missing type', () => {
      const hello = {
        version: 1,
        clientId: 'client-123',
      } as unknown as RTDBHelloMessage

      expect(validateHandshake(hello)).toBe(false)
    })
  })
})

// ============================================================================
// Connection Lifecycle Tests
// ============================================================================

describe('Connection Lifecycle', () => {
  let server: RTDBServer
  let mockWs: ReturnType<typeof createMockWebSocket>
  let connection: RTDBConnection

  beforeEach(async () => {
    server = createRTDBServer()
    mockWs = createMockWebSocket()
    connection = await server.handleConnection(mockWs)

    // Complete handshake
    const helloMessage = JSON.stringify({
      type: 'hello',
      version: 1,
      clientId: 'client-123',
    })
    mockWs.mockTrigger('message', { data: helloMessage })

    await vi.waitFor(() => {
      expect(connection.state).toBe(ConnectionState.CONNECTED)
    })
  })

  afterEach(async () => {
    if (connection) {
      await connection.close()
    }
    await server.close()
  })

  describe('Connection States', () => {
    it('should start in CONNECTING state before upgrade', () => {
      const freshWs = createMockWebSocket()
      const freshConnection = new RTDBConnection(freshWs)

      expect(freshConnection.state).toBe(ConnectionState.CONNECTING)
    })

    it('should transition to HANDSHAKING after WebSocket opens', async () => {
      const freshWs = createMockWebSocket()
      const freshConnection = await server.handleConnection(freshWs)

      expect(freshConnection.state).toBe(ConnectionState.HANDSHAKING)
    })

    it('should transition to CONNECTED after successful handshake', () => {
      expect(connection.state).toBe(ConnectionState.CONNECTED)
    })

    it('should transition to CLOSING when close is initiated', async () => {
      const closePromise = connection.close()

      expect(connection.state).toBe(ConnectionState.CLOSING)

      await closePromise
    })

    it('should transition to CLOSED after close completes', async () => {
      await connection.close()

      expect(connection.state).toBe(ConnectionState.CLOSED)
    })

    it('should transition to CLOSED on WebSocket error', async () => {
      mockWs.mockTrigger('error', new Error('Connection lost'))

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CLOSED)
      })
    })

    it('should transition to CLOSED on WebSocket close', async () => {
      mockWs.mockTrigger('close', { code: 1000, reason: 'Normal closure' })

      await vi.waitFor(() => {
        expect(connection.state).toBe(ConnectionState.CLOSED)
      })
    })
  })

  describe('Connection Properties', () => {
    it('should have unique connection id', () => {
      expect(connection.id).toBeDefined()
      expect(connection.id.length).toBeGreaterThan(0)
    })

    it('should track client id after handshake', () => {
      expect(connection.clientId).toBe('client-123')
    })

    it('should have session id after handshake', () => {
      expect(connection.sessionId).toBeDefined()
    })

    it('should track connection creation time', () => {
      expect(connection.createdAt).toBeInstanceOf(Date)
      expect(connection.createdAt.getTime()).toBeLessThanOrEqual(Date.now())
    })

    it('should track last activity time', () => {
      expect(connection.lastActivityAt).toBeInstanceOf(Date)
    })

    it('should update last activity on message received', async () => {
      const initialActivity = connection.lastActivityAt

      await new Promise((resolve) => setTimeout(resolve, 10))

      mockWs.mockTrigger('message', { data: JSON.stringify({ type: 'ping' }) })

      expect(connection.lastActivityAt.getTime()).toBeGreaterThan(
        initialActivity.getTime()
      )
    })
  })

  describe('Heartbeat / Keep-Alive', () => {
    it('should respond to ping with pong', async () => {
      const pingMessage = JSON.stringify({
        type: 'ping',
        timestamp: Date.now(),
      })

      mockWs.mockTrigger('message', { data: pingMessage })

      await vi.waitFor(() => {
        const lastMessage = mockWs.sentMessages[mockWs.sentMessages.length - 1]
        const parsed = JSON.parse(lastMessage)
        expect(parsed.type).toBe('pong')
      })
    })

    it('should include original timestamp in pong', async () => {
      const timestamp = Date.now()
      const pingMessage = JSON.stringify({
        type: 'ping',
        timestamp,
      })

      mockWs.mockTrigger('message', { data: pingMessage })

      await vi.waitFor(() => {
        const lastMessage = mockWs.sentMessages[mockWs.sentMessages.length - 1]
        const parsed = JSON.parse(lastMessage)
        expect(parsed.echoTimestamp).toBe(timestamp)
      })
    })

    it('should send periodic pings when configured', async () => {
      const serverWithHeartbeat = createRTDBServer({
        heartbeatInterval: 50, // 50ms for testing
      })

      const ws = createMockWebSocket()
      const conn = await serverWithHeartbeat.handleConnection(ws)

      // Complete handshake
      ws.mockTrigger('message', {
        data: JSON.stringify({
          type: 'hello',
          version: 1,
          clientId: 'test',
        }),
      })

      await vi.waitFor(() => {
        expect(conn.state).toBe(ConnectionState.CONNECTED)
      })

      // Wait for heartbeat
      await vi.waitFor(
        () => {
          const messages = ws.sentMessages.filter((m) => {
            const parsed = JSON.parse(m)
            return parsed.type === 'ping'
          })
          expect(messages.length).toBeGreaterThan(0)
        },
        { timeout: 200 }
      )

      await conn.close()
      await serverWithHeartbeat.close()
    })

    it('should close connection if no pong received within timeout', async () => {
      const serverWithStrictHeartbeat = createRTDBServer({
        heartbeatInterval: 50,
        heartbeatTimeout: 100,
      })

      const ws = createMockWebSocket()
      const conn = await serverWithStrictHeartbeat.handleConnection(ws)

      // Complete handshake
      ws.mockTrigger('message', {
        data: JSON.stringify({
          type: 'hello',
          version: 1,
          clientId: 'test',
        }),
      })

      await vi.waitFor(() => {
        expect(conn.state).toBe(ConnectionState.CONNECTED)
      })

      // Don't respond to pings - connection should close
      await vi.waitFor(
        () => {
          expect(conn.state).toBe(ConnectionState.CLOSED)
        },
        { timeout: 500 }
      )

      await serverWithStrictHeartbeat.close()
    })
  })

  describe('Graceful Shutdown', () => {
    it('should send goodbye message before closing', async () => {
      await connection.close()

      const messages = mockWs.sentMessages
      const lastMessage = JSON.parse(messages[messages.length - 1])

      expect(lastMessage.type).toBe('goodbye')
    })

    it('should include reason in goodbye message', async () => {
      await connection.close('Session expired')

      const messages = mockWs.sentMessages
      const lastMessage = JSON.parse(messages[messages.length - 1])

      expect(lastMessage.type).toBe('goodbye')
      expect(lastMessage.reason).toBe('Session expired')
    })

    it('should close WebSocket with normal close code', async () => {
      await connection.close()

      expect(mockWs.close).toHaveBeenCalledWith(1000, expect.any(String))
    })

    it('should handle close during handshake', async () => {
      const freshWs = createMockWebSocket()
      const freshConnection = await server.handleConnection(freshWs)

      expect(freshConnection.state).toBe(ConnectionState.HANDSHAKING)

      await freshConnection.close()

      expect(freshConnection.state).toBe(ConnectionState.CLOSED)
    })
  })

  describe('Error Handling', () => {
    it('should emit error event on connection error', async () => {
      const errorHandler = vi.fn()
      connection.on('error', errorHandler)

      mockWs.mockTrigger('error', new Error('Network error'))

      await vi.waitFor(() => {
        expect(errorHandler).toHaveBeenCalled()
      })
    })

    it('should emit close event when connection closes', async () => {
      const closeHandler = vi.fn()
      connection.on('close', closeHandler)

      mockWs.mockTrigger('close', { code: 1000, reason: 'Normal' })

      await vi.waitFor(() => {
        expect(closeHandler).toHaveBeenCalled()
      })
    })

    it('should include close code in close event', async () => {
      const closeHandler = vi.fn()
      connection.on('close', closeHandler)

      mockWs.mockTrigger('close', { code: 1001, reason: 'Going away' })

      await vi.waitFor(() => {
        expect(closeHandler).toHaveBeenCalledWith(
          expect.objectContaining({
            code: 1001,
            reason: 'Going away',
          })
        )
      })
    })

    it('should not accept messages after close', async () => {
      await connection.close()

      expect(() => {
        connection.send({ type: 'test' })
      }).toThrow(RTDBConnectionError)
    })

    it('should handle malformed messages gracefully', async () => {
      const errorHandler = vi.fn()
      connection.on('error', errorHandler)

      mockWs.mockTrigger('message', { data: 'not json {{{' })

      await vi.waitFor(() => {
        expect(errorHandler).toHaveBeenCalled()
      })

      // Connection should remain open
      expect(connection.state).toBe(ConnectionState.CONNECTED)
    })
  })

  describe('Connection Tracking', () => {
    it('should track active connections on server', async () => {
      expect(server.activeConnections).toBe(1)
    })

    it('should decrement active connections on close', async () => {
      await connection.close()

      expect(server.activeConnections).toBe(0)
    })

    it('should track multiple connections', async () => {
      const ws2 = createMockWebSocket()
      const conn2 = await server.handleConnection(ws2)

      // Complete handshake for second connection
      ws2.mockTrigger('message', {
        data: JSON.stringify({
          type: 'hello',
          version: 1,
          clientId: 'client-456',
        }),
      })

      await vi.waitFor(() => {
        expect(conn2.state).toBe(ConnectionState.CONNECTED)
      })

      expect(server.activeConnections).toBe(2)

      await conn2.close()

      expect(server.activeConnections).toBe(1)
    })

    it('should provide list of all connections', () => {
      const connections = server.getConnections()

      expect(connections).toHaveLength(1)
      expect(connections[0].clientId).toBe('client-123')
    })

    it('should find connection by id', () => {
      const found = server.getConnection(connection.id)

      expect(found).toBe(connection)
    })

    it('should find connection by client id', () => {
      const found = server.getConnectionByClientId('client-123')

      expect(found).toBe(connection)
    })

    it('should return undefined for unknown connection id', () => {
      const found = server.getConnection('unknown-id')

      expect(found).toBeUndefined()
    })
  })

  describe('Server Shutdown', () => {
    it('should close all connections on server shutdown', async () => {
      const ws2 = createMockWebSocket()
      const conn2 = await server.handleConnection(ws2)

      ws2.mockTrigger('message', {
        data: JSON.stringify({
          type: 'hello',
          version: 1,
          clientId: 'client-456',
        }),
      })

      await vi.waitFor(() => {
        expect(conn2.state).toBe(ConnectionState.CONNECTED)
      })

      await server.close()

      expect(connection.state).toBe(ConnectionState.CLOSED)
      expect(conn2.state).toBe(ConnectionState.CLOSED)
    })

    it('should send goodbye to all connections on shutdown', async () => {
      await server.close()

      const lastMessage = JSON.parse(
        mockWs.sentMessages[mockWs.sentMessages.length - 1]
      )

      expect(lastMessage.type).toBe('goodbye')
      expect(lastMessage.reason).toContain('shutdown')
    })

    it('should not accept new connections after shutdown', async () => {
      await server.close()

      const ws2 = createMockWebSocket()

      await expect(server.handleConnection(ws2)).rejects.toThrow(
        RTDBConnectionError
      )
    })
  })
})

// ============================================================================
// Error Classes Tests
// ============================================================================

describe('Error Classes', () => {
  describe('RTDBHandshakeError', () => {
    it('should have correct name and message', () => {
      const error = new RTDBHandshakeError('Invalid hello message')

      expect(error.name).toBe('RTDBHandshakeError')
      expect(error.message).toBe('Invalid hello message')
    })

    it('should have close code', () => {
      const error = new RTDBHandshakeError('Timeout', 4000)

      expect(error.code).toBe(4000)
    })
  })

  describe('RTDBProtocolError', () => {
    it('should have correct name and message', () => {
      const error = new RTDBProtocolError('Invalid message format')

      expect(error.name).toBe('RTDBProtocolError')
      expect(error.message).toBe('Invalid message format')
    })
  })

  describe('RTDBConnectionError', () => {
    it('should have correct name and message', () => {
      const error = new RTDBConnectionError('Connection closed')

      expect(error.name).toBe('RTDBConnectionError')
      expect(error.message).toBe('Connection closed')
    })

    it('should include cause', () => {
      const cause = new Error('Network failure')
      const error = new RTDBConnectionError('Connection failed', cause)

      expect(error.cause).toBe(cause)
    })
  })
})

// ============================================================================
// Message Types Tests
// ============================================================================

describe('Message Types', () => {
  describe('MessageType enum', () => {
    it('should define hello message type', () => {
      expect(MessageType.HELLO).toBe('hello')
    })

    it('should define goodbye message type', () => {
      expect(MessageType.GOODBYE).toBe('goodbye')
    })

    it('should define ping message type', () => {
      expect(MessageType.PING).toBe('ping')
    })

    it('should define pong message type', () => {
      expect(MessageType.PONG).toBe('pong')
    })

    it('should define error message type', () => {
      expect(MessageType.ERROR).toBe('error')
    })
  })
})
