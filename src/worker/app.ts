/**
 * Hono App Setup (STUB - Implementation Required)
 *
 * Layer 9: Worker HTTP Entrypoint
 *
 * This is a stub file for RED phase TDD. All exports are minimal
 * implementations that will cause tests to fail appropriately.
 */

import { Hono } from 'hono'

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
 * Create a new Hono app instance
 * TODO: Implement proper app configuration
 */
export function createApp(): Hono<AppContext> {
  // Stub implementation - will fail tests
  throw new Error('createApp not implemented')
}

/**
 * Get all registered routes
 * TODO: Implement route introspection
 */
export function getRoutes(): RouteInfo[] {
  // Stub implementation - will fail tests
  throw new Error('getRoutes not implemented')
}

/**
 * Default app instance
 * TODO: Implement with proper configuration
 */
export const app = new Hono<AppContext>()

// Export default for Cloudflare Workers
export default app
