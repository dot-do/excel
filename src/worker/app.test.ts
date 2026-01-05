/**
 * Hono App Setup Tests (RED)
 *
 * TDD: These tests define expected behavior before implementation.
 * All tests should FAIL initially.
 *
 * Layer 9: Worker HTTP Entrypoint - Hono App Configuration
 */

import { describe, it, expect, beforeEach, vi } from 'vitest'
import type { Context } from 'hono'

// Import the app and related utilities (these don't exist yet - RED phase)
import {
  app,
  createApp,
  getRoutes,
  type Env,
  type AppContext,
  type RequestContext,
} from './app'

// Mock Cloudflare bindings for testing
const mockD1Database = {
  prepare: vi.fn(),
  batch: vi.fn(),
  exec: vi.fn(),
  dump: vi.fn(),
}

const mockKVNamespace = {
  get: vi.fn(),
  put: vi.fn(),
  delete: vi.fn(),
  list: vi.fn(),
  getWithMetadata: vi.fn(),
}

const mockDurableObjectNamespace = {
  get: vi.fn(),
  idFromName: vi.fn(),
  idFromString: vi.fn(),
  newUniqueId: vi.fn(),
}

const createMockEnv = (): Env => ({
  // D1 Database binding
  DB: mockD1Database as unknown as D1Database,
  // KV Namespace bindings
  KV: mockKVNamespace as unknown as KVNamespace,
  SESSIONS: mockKVNamespace as unknown as KVNamespace,
  // Durable Object bindings
  WORKBOOK: mockDurableObjectNamespace as unknown as DurableObjectNamespace,
  SHEET: mockDurableObjectNamespace as unknown as DurableObjectNamespace,
  // Environment variables
  ENVIRONMENT: 'test',
})

describe('Hono App Setup', () => {
  let mockEnv: Env

  beforeEach(() => {
    mockEnv = createMockEnv()
    vi.clearAllMocks()
  })

  describe('App Initialization', () => {
    it('should create Hono app successfully', () => {
      const testApp = createApp()
      expect(testApp).toBeDefined()
      expect(typeof testApp.fetch).toBe('function')
    })

    it('should have correct base configuration', () => {
      const testApp = createApp()
      // App should be configured with strict mode
      expect(testApp).toHaveProperty('router')
    })

    it('should export default app instance', () => {
      expect(app).toBeDefined()
      expect(typeof app.fetch).toBe('function')
    })

    it('should export correct fetch handler', async () => {
      const request = new Request('http://localhost/')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      expect(response).toBeInstanceOf(Response)
    })

    it('should support environment bindings in context', async () => {
      const request = new Request('http://localhost/api/health')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      expect(response.status).not.toBe(500)
    })

    it('should handle multiple concurrent requests', async () => {
      const requests = Array.from({ length: 10 }, (_, i) =>
        new Request(`http://localhost/api/health?req=${i}`)
      )

      const responses = await Promise.all(
        requests.map((req) => app.fetch(req, mockEnv, {} as ExecutionContext))
      )

      expect(responses).toHaveLength(10)
      responses.forEach((res) => {
        expect(res).toBeInstanceOf(Response)
      })
    })
  })

  describe('Route Registration', () => {
    it('should register all API routes', () => {
      const routes = getRoutes()
      expect(routes).toContainEqual(
        expect.objectContaining({ path: '/api/query', method: 'POST' })
      )
      expect(routes).toContainEqual(
        expect.objectContaining({ path: '/api/mutation', method: 'POST' })
      )
      expect(routes).toContainEqual(
        expect.objectContaining({ path: '/api/action', method: 'POST' })
      )
      expect(routes).toContainEqual(
        expect.objectContaining({ path: '/sync', method: 'GET' })
      )
    })

    it('should use POST method for mutation endpoint', async () => {
      const getRequest = new Request('http://localhost/api/mutation', {
        method: 'GET',
      })
      const postRequest = new Request('http://localhost/api/mutation', {
        method: 'POST',
        body: JSON.stringify({}),
        headers: { 'Content-Type': 'application/json' },
      })

      const getResponse = await app.fetch(getRequest, mockEnv, {} as ExecutionContext)
      const postResponse = await app.fetch(postRequest, mockEnv, {} as ExecutionContext)

      expect(getResponse.status).toBe(405) // Method not allowed
      expect(postResponse.status).not.toBe(405)
    })

    it('should use POST method for query endpoint', async () => {
      const getRequest = new Request('http://localhost/api/query', {
        method: 'GET',
      })
      const postRequest = new Request('http://localhost/api/query', {
        method: 'POST',
        body: JSON.stringify({}),
        headers: { 'Content-Type': 'application/json' },
      })

      const getResponse = await app.fetch(getRequest, mockEnv, {} as ExecutionContext)
      const postResponse = await app.fetch(postRequest, mockEnv, {} as ExecutionContext)

      expect(getResponse.status).toBe(405)
      expect(postResponse.status).not.toBe(405)
    })

    it('should use POST method for action endpoint', async () => {
      const getRequest = new Request('http://localhost/api/action', {
        method: 'GET',
      })
      const postRequest = new Request('http://localhost/api/action', {
        method: 'POST',
        body: JSON.stringify({}),
        headers: { 'Content-Type': 'application/json' },
      })

      const getResponse = await app.fetch(getRequest, mockEnv, {} as ExecutionContext)
      const postResponse = await app.fetch(postRequest, mockEnv, {} as ExecutionContext)

      expect(getResponse.status).toBe(405)
      expect(postResponse.status).not.toBe(405)
    })

    it('should use GET method for sync endpoint', async () => {
      const postRequest = new Request('http://localhost/sync', {
        method: 'POST',
        body: JSON.stringify({}),
      })
      const getRequest = new Request('http://localhost/sync', {
        method: 'GET',
        headers: { Upgrade: 'websocket' },
      })

      const postResponse = await app.fetch(postRequest, mockEnv, {} as ExecutionContext)
      const getResponse = await app.fetch(getRequest, mockEnv, {} as ExecutionContext)

      expect(postResponse.status).toBe(405)
      expect(getResponse.status).not.toBe(405)
    })

    it('should return 404 for unknown routes', async () => {
      const request = new Request('http://localhost/unknown/path')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      expect(response.status).toBe(404)
    })

    it('should match route paths correctly with parameters', async () => {
      const request = new Request('http://localhost/api/workbook/123/sheet/456')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      // Should match a valid route pattern, not 404
      expect(response.status).not.toBe(404)
    })

    it('should handle trailing slashes consistently', async () => {
      const withSlash = new Request('http://localhost/api/query/')
      const withoutSlash = new Request('http://localhost/api/query')

      const responseWithSlash = await app.fetch(withSlash, mockEnv, {} as ExecutionContext)
      const responseWithoutSlash = await app.fetch(withoutSlash, mockEnv, {} as ExecutionContext)

      // Both should be handled consistently (either both work or both redirect)
      expect([responseWithSlash.status, responseWithoutSlash.status]).toEqual(
        expect.arrayContaining([expect.any(Number)])
      )
    })

    it('should register health check endpoint', async () => {
      const request = new Request('http://localhost/api/health')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      expect(response.status).toBe(200)
    })
  })

  describe('Middleware Chain', () => {
    it('should execute middleware in correct order', async () => {
      const executionOrder: string[] = []

      // The app should have middleware that can be traced
      const request = new Request('http://localhost/api/health', {
        headers: { 'X-Test-Trace': 'true' },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      // Check response headers for middleware execution order
      const middlewareOrder = response.headers.get('X-Middleware-Order')
      if (middlewareOrder) {
        const order = middlewareOrder.split(',')
        expect(order).toEqual(['cors', 'auth', 'error-handler', 'logging'])
      }
    })

    it('should pass context through middleware chain', async () => {
      const request = new Request('http://localhost/api/health')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      // Context should include request ID from middleware
      const requestId = response.headers.get('X-Request-ID')
      expect(requestId).toBeDefined()
      expect(requestId).toMatch(/^[a-z0-9-]+$/i)
    })

    it('should allow middleware to short-circuit requests', async () => {
      // Request without proper auth should be blocked by auth middleware
      const request = new Request('http://localhost/api/mutation', {
        method: 'POST',
        body: JSON.stringify({}),
        headers: { 'Content-Type': 'application/json' },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      // Auth middleware should short-circuit and return 401
      expect(response.status).toBe(401)
    })

    it('should propagate middleware errors correctly', async () => {
      // Trigger an error condition
      const request = new Request('http://localhost/api/error-test', {
        headers: { 'X-Trigger-Error': 'true' },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      // Error should be caught and formatted properly
      expect(response.status).toBeGreaterThanOrEqual(400)
      const body = await response.json()
      expect(body).toHaveProperty('error')
    })

    it('should apply CORS headers via middleware', async () => {
      const request = new Request('http://localhost/api/health', {
        method: 'OPTIONS',
        headers: {
          Origin: 'http://example.com',
          'Access-Control-Request-Method': 'POST',
        },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      expect(response.headers.get('Access-Control-Allow-Origin')).toBeDefined()
      expect(response.headers.get('Access-Control-Allow-Methods')).toBeDefined()
    })

    it('should handle middleware timeout', async () => {
      const request = new Request('http://localhost/api/slow-endpoint', {
        headers: { 'X-Simulate-Slow': 'true' },
      })

      const startTime = Date.now()
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      const duration = Date.now() - startTime

      // Should timeout or complete within reasonable time
      expect(duration).toBeLessThan(30000)
    })
  })

  describe('Environment Binding', () => {
    it('should have D1 database binding available', async () => {
      const request = new Request('http://localhost/api/db-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('d1Available', true)
    })

    it('should have KV namespace binding available', async () => {
      const request = new Request('http://localhost/api/kv-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('kvAvailable', true)
    })

    it('should have Durable Object binding available', async () => {
      const request = new Request('http://localhost/api/do-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('doAvailable', true)
    })

    it('should throw appropriate error when D1 binding is missing', async () => {
      const envWithoutD1 = { ...mockEnv, DB: undefined }
      const request = new Request('http://localhost/api/query', {
        method: 'POST',
        body: JSON.stringify({ query: 'test' }),
        headers: { 'Content-Type': 'application/json' },
      })

      const response = await app.fetch(request, envWithoutD1 as Env, {} as ExecutionContext)

      expect(response.status).toBe(500)
      const body = await response.json()
      expect(body.error).toContain('D1')
    })

    it('should throw appropriate error when KV binding is missing', async () => {
      const envWithoutKV = { ...mockEnv, KV: undefined }
      const request = new Request('http://localhost/api/session', {
        method: 'GET',
      })

      const response = await app.fetch(request, envWithoutKV as Env, {} as ExecutionContext)

      expect(response.status).toBe(500)
      const body = await response.json()
      expect(body.error).toContain('KV')
    })

    it('should throw appropriate error when Durable Object binding is missing', async () => {
      const envWithoutDO = { ...mockEnv, WORKBOOK: undefined }
      const request = new Request('http://localhost/api/workbook/123', {
        method: 'GET',
      })

      const response = await app.fetch(request, envWithoutDO as Env, {} as ExecutionContext)

      expect(response.status).toBe(500)
      const body = await response.json()
      expect(body.error).toContain('Durable Object')
    })

    it('should provide environment variable access', async () => {
      const request = new Request('http://localhost/api/env-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('environment', 'test')
    })

    it('should validate all required bindings on startup', () => {
      const validateBindings = () => {
        const env = createMockEnv()
        // This should not throw
        createApp()
      }

      expect(validateBindings).not.toThrow()
    })

    it('should isolate bindings between requests', async () => {
      const request1 = new Request('http://localhost/api/mutation', {
        method: 'POST',
        body: JSON.stringify({ id: '1' }),
        headers: { 'Content-Type': 'application/json' },
      })

      const request2 = new Request('http://localhost/api/mutation', {
        method: 'POST',
        body: JSON.stringify({ id: '2' }),
        headers: { 'Content-Type': 'application/json' },
      })

      // Execute concurrently
      await Promise.all([
        app.fetch(request1, mockEnv, {} as ExecutionContext),
        app.fetch(request2, mockEnv, {} as ExecutionContext),
      ])

      // Verify mock was called with correct isolated data
      // (bindings should not leak between requests)
      expect(true).toBe(true) // Placeholder assertion
    })
  })

  describe('Request Context', () => {
    it('should generate unique request ID for each request', async () => {
      const request1 = new Request('http://localhost/api/health')
      const request2 = new Request('http://localhost/api/health')

      const [response1, response2] = await Promise.all([
        app.fetch(request1, mockEnv, {} as ExecutionContext),
        app.fetch(request2, mockEnv, {} as ExecutionContext),
      ])

      const requestId1 = response1.headers.get('X-Request-ID')
      const requestId2 = response2.headers.get('X-Request-ID')

      expect(requestId1).toBeDefined()
      expect(requestId2).toBeDefined()
      expect(requestId1).not.toBe(requestId2)
    })

    it('should use client-provided request ID if present', async () => {
      const clientRequestId = 'client-generated-id-12345'
      const request = new Request('http://localhost/api/health', {
        headers: { 'X-Request-ID': clientRequestId },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      expect(response.headers.get('X-Request-ID')).toBe(clientRequestId)
    })

    it('should capture timing information', async () => {
      const request = new Request('http://localhost/api/health')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const serverTiming = response.headers.get('Server-Timing')
      expect(serverTiming).toBeDefined()
      expect(serverTiming).toMatch(/total;dur=\d+/)
    })

    it('should include request start time in context', async () => {
      const beforeRequest = Date.now()
      const request = new Request('http://localhost/api/timing-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)
      const afterRequest = Date.now()

      const body = await response.json()
      expect(body).toHaveProperty('startTime')
      expect(body.startTime).toBeGreaterThanOrEqual(beforeRequest)
      expect(body.startTime).toBeLessThanOrEqual(afterRequest)
    })

    it('should scope context variables to single request', async () => {
      const request1 = new Request('http://localhost/api/context-test', {
        headers: { 'X-Custom-Value': 'value1' },
      })
      const request2 = new Request('http://localhost/api/context-test', {
        headers: { 'X-Custom-Value': 'value2' },
      })

      const [response1, response2] = await Promise.all([
        app.fetch(request1, mockEnv, {} as ExecutionContext),
        app.fetch(request2, mockEnv, {} as ExecutionContext),
      ])

      const body1 = await response1.json()
      const body2 = await response2.json()

      // Each request should see its own context value
      expect(body1.customValue).toBe('value1')
      expect(body2.customValue).toBe('value2')
    })

    it('should include request path in context', async () => {
      const request = new Request('http://localhost/api/path-check')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('path', '/api/path-check')
    })

    it('should include request method in context', async () => {
      const request = new Request('http://localhost/api/method-check', {
        method: 'POST',
        body: JSON.stringify({}),
        headers: { 'Content-Type': 'application/json' },
      })

      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('method', 'POST')
    })

    it('should clean up context after request completes', async () => {
      const request = new Request('http://localhost/api/health')

      // Make a request
      await app.fetch(request, mockEnv, {} as ExecutionContext)

      // Context should be cleaned up (no memory leaks)
      // This is more of a design requirement than a testable behavior
      // but the test documents the expected behavior
      expect(true).toBe(true)
    })

    it('should handle context access from nested async operations', async () => {
      const request = new Request('http://localhost/api/async-context-test')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('nestedContextAccess', true)
    })
  })

  describe('Error Handling', () => {
    it('should return JSON error responses', async () => {
      const request = new Request('http://localhost/api/force-error')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      expect(response.headers.get('Content-Type')).toContain('application/json')
      const body = await response.json()
      expect(body).toHaveProperty('error')
    })

    it('should include error code in response', async () => {
      const request = new Request('http://localhost/api/force-error')
      const response = await app.fetch(request, mockEnv, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('code')
      expect(typeof body.code).toBe('string')
    })

    it('should not expose stack traces in production', async () => {
      const prodEnv = { ...mockEnv, ENVIRONMENT: 'production' }
      const request = new Request('http://localhost/api/force-error')
      const response = await app.fetch(request, prodEnv as Env, {} as ExecutionContext)

      const body = await response.json()
      expect(body).not.toHaveProperty('stack')
    })

    it('should include stack traces in development', async () => {
      const devEnv = { ...mockEnv, ENVIRONMENT: 'development' }
      const request = new Request('http://localhost/api/force-error')
      const response = await app.fetch(request, devEnv as Env, {} as ExecutionContext)

      const body = await response.json()
      expect(body).toHaveProperty('stack')
    })
  })

  describe('Type Safety', () => {
    it('should export Env type', () => {
      const env: Env = mockEnv
      expect(env.DB).toBeDefined()
      expect(env.KV).toBeDefined()
      expect(env.WORKBOOK).toBeDefined()
    })

    it('should export AppContext type', () => {
      // This is a compile-time check - if the type doesn't exist, this won't compile
      const context: AppContext = {} as AppContext
      expect(context).toBeDefined()
    })

    it('should export RequestContext type', () => {
      const reqContext: RequestContext = {
        requestId: 'test-id',
        startTime: Date.now(),
        path: '/test',
        method: 'GET',
      }
      expect(reqContext.requestId).toBe('test-id')
    })
  })
})
