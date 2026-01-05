/**
 * Types for Excel Formula Parser
 *
 * Defines Token types and AST node structures for parsing Excel formulas.
 */

// ============================================================================
// Token Types
// ============================================================================

/**
 * Token type enumeration for formula lexer
 */
export type TokenType =
  // Literals
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'ERROR'
  // References
  | 'CELL_REF'
  | 'RANGE'
  | 'NAMED_RANGE'
  // Operators
  | 'OPERATOR_ADD'
  | 'OPERATOR_SUB'
  | 'OPERATOR_MUL'
  | 'OPERATOR_DIV'
  | 'OPERATOR_POW'
  | 'OPERATOR_CONCAT'
  | 'OPERATOR_EQ'
  | 'OPERATOR_NE'
  | 'OPERATOR_LT'
  | 'OPERATOR_GT'
  | 'OPERATOR_LTE'
  | 'OPERATOR_GTE'
  // Punctuation
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'COLON'
  | 'SEMICOLON'
  // Functions
  | 'FUNCTION'
  // Special
  | 'WHITESPACE'
  | 'EOF'

/**
 * Represents a single token produced by the lexer
 */
export interface Token {
  /** The type of token */
  type: TokenType
  /** The raw string value of the token */
  value: string
  /** Starting position in the formula string */
  start: number
  /** Ending position in the formula string */
  end: number
}

// ============================================================================
// AST Node Types
// ============================================================================

/**
 * AST node type enumeration
 */
export type ASTNodeType =
  | 'Program'
  | 'BinaryExpression'
  | 'UnaryExpression'
  | 'FunctionCall'
  | 'CellReference'
  | 'RangeReference'
  | 'Literal'
  | 'ArrayLiteral'
  | 'ErrorValue'

/**
 * Base interface for all AST nodes
 */
export interface ASTNodeBase {
  type: ASTNodeType
  /** Starting position in source */
  start: number
  /** Ending position in source */
  end: number
}

/**
 * Root node of the AST
 */
export interface ProgramNode extends ASTNodeBase {
  type: 'Program'
  body: ASTNode
}

/**
 * Binary expression: A1 + B1, 5 * 3, etc.
 */
export interface BinaryExpressionNode extends ASTNodeBase {
  type: 'BinaryExpression'
  operator: string
  left: ASTNode
  right: ASTNode
}

/**
 * Unary expression: -A1, +5, etc.
 */
export interface UnaryExpressionNode extends ASTNodeBase {
  type: 'UnaryExpression'
  operator: string
  argument: ASTNode
  prefix: boolean
}

/**
 * Function call: SUM(A1:A10), IF(A1>0, "Yes", "No")
 */
export interface FunctionCallNode extends ASTNodeBase {
  type: 'FunctionCall'
  name: string
  arguments: ASTNode[]
}

/**
 * Cell reference: A1, $A$1, Sheet1!A1, etc.
 */
export interface CellReferenceNode extends ASTNodeBase {
  type: 'CellReference'
  /** Column letter(s) */
  column: string
  /** Row number */
  row: number
  /** Is column absolute? ($A vs A) */
  columnAbsolute: boolean
  /** Is row absolute? ($1 vs 1) */
  rowAbsolute: boolean
  /** Sheet name if cross-sheet reference */
  sheet?: string
}

/**
 * Range reference: A1:B10, $A$1:$B$10
 */
export interface RangeReferenceNode extends Omit<ASTNodeBase, 'start' | 'end'> {
  type: 'RangeReference'
  /** Starting position in source */
  start: number
  /** Ending position in source */
  end: number
  /** Start cell of range */
  startCell: CellReferenceNode
  /** End cell of range */
  endCell: CellReferenceNode
  /** Sheet name if cross-sheet reference */
  sheet?: string
}

/**
 * Literal value: number, string, boolean
 */
export interface LiteralNode extends ASTNodeBase {
  type: 'Literal'
  value: number | string | boolean | null
  /** Original type hint */
  valueType: 'number' | 'string' | 'boolean' | 'null'
}

/**
 * Array literal: {1,2,3} or {1,2;3,4}
 */
export interface ArrayLiteralNode extends ASTNodeBase {
  type: 'ArrayLiteral'
  elements: ASTNode[][]
}

/**
 * Error value: #REF!, #N/A, etc.
 */
export interface ErrorValueNode extends ASTNodeBase {
  type: 'ErrorValue'
  error: string
}

/**
 * Union type of all AST nodes
 */
export type ASTNode =
  | ProgramNode
  | BinaryExpressionNode
  | UnaryExpressionNode
  | FunctionCallNode
  | CellReferenceNode
  | RangeReferenceNode
  | LiteralNode
  | ArrayLiteralNode
  | ErrorValueNode

// ============================================================================
// Parser Options and Results
// ============================================================================

/**
 * Options for the parser
 */
export interface ParserOptions {
  /** Whether to include position information in AST nodes */
  includePositions?: boolean
  /** Whether to allow array formulas */
  allowArrayFormulas?: boolean
  /** Locale settings for decimal/argument separators */
  locale?: {
    decimalSeparator: string
    argumentSeparator: string
  }
}

/**
 * Result of parsing a formula
 */
export interface ParseResult {
  /** The parsed AST */
  ast: ProgramNode
  /** Any errors encountered during parsing */
  errors: ParseError[]
  /** The original formula string */
  formula: string
}

/**
 * Parsing error information
 */
export interface ParseError {
  /** Error message */
  message: string
  /** Position in the formula where error occurred */
  position: number
  /** Length of the problematic token/section */
  length: number
  /** Error code for programmatic handling */
  code: ParseErrorCode
}

/**
 * Error codes for parse errors
 */
export type ParseErrorCode =
  | 'UNEXPECTED_TOKEN'
  | 'UNEXPECTED_EOF'
  | 'INVALID_CELL_REF'
  | 'INVALID_RANGE'
  | 'UNCLOSED_PAREN'
  | 'UNCLOSED_STRING'
  | 'INVALID_FUNCTION'
  | 'MISSING_OPERAND'
  | 'INVALID_OPERATOR'
  | 'SYNTAX_ERROR'
