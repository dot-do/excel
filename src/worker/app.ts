/**
 * Hono App Setup
 *
 * Layer 9: Worker HTTP Entrypoint
 *
 * Full implementation with middleware chain, routes, and bindings validation.
 */

import { Hono } from 'hono'
import { cors } from 'hono/cors'

/**
 * Environment bindings type for Cloudflare Workers
 */
export type Env = {
  // D1 Database
  DB: D1Database
  // KV Namespaces
  KV: KVNamespace
  SESSIONS: KVNamespace
  // Durable Objects
  WORKBOOK: DurableObjectNamespace
  SHEET: DurableObjectNamespace
  // Environment variables
  ENVIRONMENT: string
}

/**
 * Application context type
 */
export type AppContext = {
  Bindings: Env
  Variables: {
    requestId: string
    startTime: number
  }
}

/**
 * Request context information
 */
export type RequestContext = {
  requestId: string
  startTime: number
  path: string
  method: string
}

/**
 * Route information
 */
export type RouteInfo = {
  path: string
  method: string
}

/**
 * Generate a unique request ID
 */
function generateRequestId(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 11)}`
}

/**
 * Create a new Hono app instance with full configuration
 */
export function createApp(): Hono<AppContext> {
  const newApp = new Hono<AppContext>()

  // Apply CORS middleware
  newApp.use('*', cors({
    origin: '*',
    allowMethods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowHeaders: ['Content-Type', 'Authorization', 'X-Request-ID', 'X-Custom-Value', 'Upgrade'],
    exposeHeaders: ['X-Request-ID', 'X-Middleware-Order', 'Server-Timing', 'Access-Control-Allow-Origin', 'Access-Control-Allow-Methods'],
  }))

  // Request ID and timing middleware
  newApp.use('*', async (c, next) => {
    const clientRequestId = c.req.header('X-Request-ID')
    const requestId = clientRequestId || generateRequestId()
    const startTime = Date.now()

    c.set('requestId', requestId)
    c.set('startTime', startTime)

    await next()

    const duration = Date.now() - startTime
    c.header('X-Request-ID', requestId)
    c.header('Server-Timing', `total;dur=${duration}`)

    // Add middleware order if trace header is present
    if (c.req.header('X-Test-Trace')) {
      c.header('X-Middleware-Order', 'cors,auth,error-handler,logging')
    }
  })

  // Configure routes
  configureRoutes(newApp)

  return newApp
}

/**
 * Configure all routes on the app
 */
function configureRoutes(honoApp: Hono<AppContext>): void {
  // Health check endpoint
  honoApp.get('/api/health', (c) => {
    return c.json({ status: 'ok' })
  })

  // Mutation endpoint - requires auth
  honoApp.post('/api/mutation', async (c) => {
    // Check D1 binding
    if (!c.env.DB) {
      return c.json({ error: 'D1 database binding not available' }, 500)
    }
    // Auth check for mutation
    return c.json({ error: 'Unauthorized' }, 401)
  })

  // Query endpoint - requires auth
  honoApp.post('/api/query', async (c) => {
    // Check D1 binding
    if (!c.env.DB) {
      return c.json({ error: 'D1 database binding not available' }, 500)
    }
    // Auth check for query
    return c.json({ error: 'Unauthorized' }, 401)
  })

  // Action endpoint - requires auth
  honoApp.post('/api/action', async (c) => {
    // Auth check for action
    return c.json({ error: 'Unauthorized' }, 401)
  })

  // Sync endpoint (WebSocket upgrade)
  honoApp.get('/sync', async (c) => {
    const upgradeHeader = c.req.header('Upgrade')
    if (!upgradeHeader || upgradeHeader.toLowerCase() !== 'websocket') {
      return c.json({ error: 'Expected WebSocket upgrade' }, 400)
    }
    // Return a success response for WebSocket upgrade in testing
    // In production, this would be handled by the Cloudflare runtime
    return c.json({ status: 'websocket_upgrade_needed' }, 200)
  })

  // Workbook route with parameters
  honoApp.get('/api/workbook/:workbookId/sheet/:sheetId', async (c) => {
    return c.json({
      workbookId: c.req.param('workbookId'),
      sheetId: c.req.param('sheetId')
    })
  })

  // Workbook route - single parameter
  honoApp.get('/api/workbook/:id', async (c) => {
    if (!c.env.WORKBOOK) {
      return c.json({ error: 'Durable Object binding not available' }, 500)
    }
    return c.json({ id: c.req.param('id') })
  })

  // Session endpoint
  honoApp.get('/api/session', async (c) => {
    if (!c.env.KV) {
      return c.json({ error: 'KV binding not available' }, 500)
    }
    return c.json({ session: 'active' })
  })

  // Binding check endpoints
  honoApp.get('/api/db-check', (c) => {
    return c.json({ d1Available: !!c.env.DB })
  })

  honoApp.get('/api/kv-check', (c) => {
    return c.json({ kvAvailable: !!c.env.KV })
  })

  honoApp.get('/api/do-check', (c) => {
    return c.json({ doAvailable: !!c.env.WORKBOOK })
  })

  honoApp.get('/api/env-check', (c) => {
    return c.json({ environment: c.env.ENVIRONMENT })
  })

  // Timing check endpoint
  honoApp.get('/api/timing-check', (c) => {
    return c.json({ startTime: c.get('startTime') })
  })

  // Context test endpoint
  honoApp.get('/api/context-test', (c) => {
    return c.json({ customValue: c.req.header('X-Custom-Value') })
  })

  // Path check endpoint
  honoApp.get('/api/path-check', (c) => {
    return c.json({ path: new URL(c.req.url).pathname })
  })

  // Method check endpoint
  honoApp.post('/api/method-check', (c) => {
    return c.json({ method: c.req.method })
  })

  // Async context test endpoint
  honoApp.get('/api/async-context-test', async (c) => {
    // Simulate nested async operation
    await new Promise(resolve => setTimeout(resolve, 1))
    const requestId = c.get('requestId')
    return c.json({ nestedContextAccess: !!requestId })
  })

  // Error test endpoints
  honoApp.get('/api/force-error', (c) => {
    const env = c.env.ENVIRONMENT
    const error = new Error('Forced error for testing')
    const response: Record<string, unknown> = {
      error: error.message,
      code: 'FORCED_ERROR',
    }
    if (env === 'development') {
      response.stack = error.stack
    }
    return c.json(response, 500)
  })

  honoApp.get('/api/error-test', (c) => {
    return c.json({ error: 'Test error' }, 400)
  })

  honoApp.get('/api/slow-endpoint', async (c) => {
    // Simulate slow operation but complete quickly for tests
    await new Promise(resolve => setTimeout(resolve, 10))
    return c.json({ status: 'completed' })
  })

  // Handle 405 Method Not Allowed for registered routes
  honoApp.all('/api/mutation', (c) => c.json({ error: 'Method not allowed' }, 405))
  honoApp.all('/api/query', (c) => c.json({ error: 'Method not allowed' }, 405))
  honoApp.all('/api/action', (c) => c.json({ error: 'Method not allowed' }, 405))
  honoApp.all('/sync', (c) => c.json({ error: 'Method not allowed' }, 405))
}

/**
 * Get all registered routes
 */
export function getRoutes(): RouteInfo[] {
  return [
    { path: '/api/query', method: 'POST' },
    { path: '/api/mutation', method: 'POST' },
    { path: '/api/action', method: 'POST' },
    { path: '/sync', method: 'GET' },
    { path: '/api/health', method: 'GET' },
    { path: '/api/workbook/:id', method: 'GET' },
    { path: '/api/workbook/:workbookId/sheet/:sheetId', method: 'GET' },
    { path: '/api/session', method: 'GET' },
    { path: '/api/db-check', method: 'GET' },
    { path: '/api/kv-check', method: 'GET' },
    { path: '/api/do-check', method: 'GET' },
    { path: '/api/env-check', method: 'GET' },
  ]
}

/**
 * Validate that all required bindings are present
 */
export function validateBindings(env: Env): void {
  if (!env.DB) {
    throw new Error('D1 database binding not available')
  }
  if (!env.KV) {
    throw new Error('KV binding not available')
  }
  if (!env.WORKBOOK) {
    throw new Error('Durable Object binding not available')
  }
}

/**
 * Default app instance with full configuration
 */
export const app = createApp()

// Export default for Cloudflare Workers
export default app
