/**
 * Excel Formula Parser
 *
 * Stub implementation for TDD - all functions throw NotImplementedError
 * to ensure tests are in RED state.
 */

import type {
  Token,
  ProgramNode,
  ParseResult,
  ParserOptions,
} from './types'

/**
 * Error thrown when a function is not yet implemented
 */
export class NotImplementedError extends Error {
  constructor(functionName: string) {
    super(`${functionName} is not yet implemented`)
    this.name = 'NotImplementedError'
  }
}

/**
 * Tokenize a formula string into an array of tokens.
 *
 * @param formula - The formula string to tokenize (with or without leading '=')
 * @returns Array of tokens
 * @throws {NotImplementedError} - Not yet implemented
 *
 * @example
 * tokenize('=A1+B1') // Returns tokens for A1, +, B1
 * tokenize('=SUM(A1:A10)') // Returns tokens for SUM(, A1:A10, )
 */
export function tokenize(_formula: string): Token[] {
  throw new NotImplementedError('tokenize')
}

/**
 * Parse a formula string into an Abstract Syntax Tree.
 *
 * @param formula - The formula string to parse (with or without leading '=')
 * @param options - Optional parser configuration
 * @returns ParseResult containing the AST and any errors
 * @throws {NotImplementedError} - Not yet implemented
 *
 * @example
 * parse('=A1+B1') // Returns AST with BinaryExpression
 * parse('=SUM(A1:A10)') // Returns AST with FunctionCall
 */
export function parse(_formula: string, _options?: ParserOptions): ParseResult {
  throw new NotImplementedError('parse')
}

/**
 * Parse tokens into an Abstract Syntax Tree.
 *
 * @param tokens - Array of tokens from tokenize()
 * @param options - Optional parser configuration
 * @returns The root AST node
 * @throws {NotImplementedError} - Not yet implemented
 */
export function parseTokens(_tokens: Token[], _options?: ParserOptions): ProgramNode {
  throw new NotImplementedError('parseTokens')
}

/**
 * Validate a formula string without fully parsing it.
 * Performs quick syntax validation.
 *
 * @param formula - The formula string to validate
 * @returns true if formula appears valid, false otherwise
 * @throws {NotImplementedError} - Not yet implemented
 */
export function validateFormula(_formula: string): boolean {
  throw new NotImplementedError('validateFormula')
}

/**
 * Extract all cell references from a formula.
 *
 * @param formula - The formula string to analyze
 * @returns Array of cell reference strings (e.g., ['A1', 'B2', 'Sheet1!C3'])
 * @throws {NotImplementedError} - Not yet implemented
 */
export function extractCellReferences(_formula: string): string[] {
  throw new NotImplementedError('extractCellReferences')
}

/**
 * Extract all range references from a formula.
 *
 * @param formula - The formula string to analyze
 * @returns Array of range reference strings (e.g., ['A1:B10', 'Sheet1!C1:C100'])
 * @throws {NotImplementedError} - Not yet implemented
 */
export function extractRangeReferences(_formula: string): string[] {
  throw new NotImplementedError('extractRangeReferences')
}

/**
 * Get all dependencies (cells and ranges) that a formula references.
 *
 * @param formula - The formula string to analyze
 * @returns Object with cells and ranges arrays
 * @throws {NotImplementedError} - Not yet implemented
 */
export function getDependencies(_formula: string): {
  cells: string[]
  ranges: string[]
} {
  throw new NotImplementedError('getDependencies')
}
