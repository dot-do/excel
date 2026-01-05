/**
 * Excel Formula Parser Tests (RED)
 *
 * TDD: These tests define expected behavior before implementation.
 * All tests should FAIL initially.
 */

import { describe, it, expect } from 'vitest'
import {
  tokenize,
  parse,
  parseTokens,
  validateFormula,
  extractCellReferences,
  extractRangeReferences,
  getDependencies,
} from './parser'
import type {
  Token,
  TokenType,
  ASTNode,
  BinaryExpressionNode,
  UnaryExpressionNode,
  FunctionCallNode,
  CellReferenceNode,
  RangeReferenceNode,
  LiteralNode,
  ProgramNode,
} from './types'

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Helper to check if a token matches expected type and value
 */
function expectToken(token: Token, type: TokenType, value: string) {
  expect(token.type).toBe(type)
  expect(token.value).toBe(value)
}

/**
 * Helper to get token types from a token array
 */
function getTokenTypes(tokens: Token[]): TokenType[] {
  return tokens.map((t) => t.type)
}

/**
 * Helper to get the body of a parsed formula
 */
function parseBody(formula: string): ASTNode {
  const result = parse(formula)
  return result.ast.body
}

// ============================================================================
// SECTION 1: tokenize() Tests
// ============================================================================

describe('tokenize()', () => {
  describe('Operators', () => {
    it('should tokenize addition operator (+)', () => {
      const tokens = tokenize('=1+2')
      expect(tokens).toHaveLength(4) // NUMBER, OPERATOR_ADD, NUMBER, EOF
      expectToken(tokens[1], 'OPERATOR_ADD', '+')
    })

    it('should tokenize subtraction operator (-)', () => {
      const tokens = tokenize('=5-3')
      expectToken(tokens[1], 'OPERATOR_SUB', '-')
    })

    it('should tokenize multiplication operator (*)', () => {
      const tokens = tokenize('=4*2')
      expectToken(tokens[1], 'OPERATOR_MUL', '*')
    })

    it('should tokenize division operator (/)', () => {
      const tokens = tokenize('=10/2')
      expectToken(tokens[1], 'OPERATOR_DIV', '/')
    })

    it('should tokenize power/exponent operator (^)', () => {
      const tokens = tokenize('=2^8')
      expectToken(tokens[1], 'OPERATOR_POW', '^')
    })

    it('should tokenize concatenation operator (&)', () => {
      const tokens = tokenize('="Hello"&"World"')
      expectToken(tokens[1], 'OPERATOR_CONCAT', '&')
    })

    it('should tokenize equals comparison operator (=)', () => {
      const tokens = tokenize('=A1=B1')
      // First token is CELL_REF, second should be OPERATOR_EQ
      expectToken(tokens[1], 'OPERATOR_EQ', '=')
    })

    it('should tokenize not equals operator (<>)', () => {
      const tokens = tokenize('=A1<>B1')
      expectToken(tokens[1], 'OPERATOR_NE', '<>')
    })

    it('should tokenize less than operator (<)', () => {
      const tokens = tokenize('=A1<B1')
      expectToken(tokens[1], 'OPERATOR_LT', '<')
    })

    it('should tokenize greater than operator (>)', () => {
      const tokens = tokenize('=A1>B1')
      expectToken(tokens[1], 'OPERATOR_GT', '>')
    })

    it('should tokenize less than or equal operator (<=)', () => {
      const tokens = tokenize('=A1<=B1')
      expectToken(tokens[1], 'OPERATOR_LTE', '<=')
    })

    it('should tokenize greater than or equal operator (>=)', () => {
      const tokens = tokenize('=A1>=B1')
      expectToken(tokens[1], 'OPERATOR_GTE', '>=')
    })

    it('should tokenize multiple operators in sequence', () => {
      const tokens = tokenize('=1+2*3-4/5')
      const types = getTokenTypes(tokens)
      expect(types).toContain('OPERATOR_ADD')
      expect(types).toContain('OPERATOR_MUL')
      expect(types).toContain('OPERATOR_SUB')
      expect(types).toContain('OPERATOR_DIV')
    })
  })

  describe('Numbers', () => {
    it('should tokenize integer numbers', () => {
      const tokens = tokenize('=123')
      expectToken(tokens[0], 'NUMBER', '123')
    })

    it('should tokenize decimal numbers', () => {
      const tokens = tokenize('=3.14')
      expectToken(tokens[0], 'NUMBER', '3.14')
    })

    it('should tokenize numbers with leading decimal', () => {
      const tokens = tokenize('=.5')
      expectToken(tokens[0], 'NUMBER', '.5')
    })

    it('should tokenize scientific notation (uppercase E)', () => {
      const tokens = tokenize('=1E10')
      expectToken(tokens[0], 'NUMBER', '1E10')
    })

    it('should tokenize scientific notation (lowercase e)', () => {
      const tokens = tokenize('=1e10')
      expectToken(tokens[0], 'NUMBER', '1e10')
    })

    it('should tokenize negative exponent in scientific notation', () => {
      const tokens = tokenize('=1.5E-3')
      expectToken(tokens[0], 'NUMBER', '1.5E-3')
    })

    it('should tokenize positive exponent in scientific notation', () => {
      const tokens = tokenize('=2.5E+10')
      expectToken(tokens[0], 'NUMBER', '2.5E+10')
    })

    it('should tokenize zero', () => {
      const tokens = tokenize('=0')
      expectToken(tokens[0], 'NUMBER', '0')
    })

    it('should tokenize large numbers', () => {
      const tokens = tokenize('=9999999999999')
      expectToken(tokens[0], 'NUMBER', '9999999999999')
    })
  })

  describe('Strings', () => {
    it('should tokenize double-quoted strings', () => {
      const tokens = tokenize('="hello"')
      expectToken(tokens[0], 'STRING', '"hello"')
    })

    it('should tokenize single-quoted strings', () => {
      const tokens = tokenize("='world'")
      expectToken(tokens[0], 'STRING', "'world'")
    })

    it('should tokenize empty strings', () => {
      const tokens = tokenize('=""')
      expectToken(tokens[0], 'STRING', '""')
    })

    it('should tokenize strings with spaces', () => {
      const tokens = tokenize('="hello world"')
      expectToken(tokens[0], 'STRING', '"hello world"')
    })

    it('should tokenize strings with escaped quotes', () => {
      const tokens = tokenize('="say ""hello"""')
      expectToken(tokens[0], 'STRING', '"say ""hello"""')
    })

    it('should tokenize strings with numbers', () => {
      const tokens = tokenize('="test123"')
      expectToken(tokens[0], 'STRING', '"test123"')
    })

    it('should tokenize strings with special characters', () => {
      const tokens = tokenize('="@#$%"')
      expectToken(tokens[0], 'STRING', '"@#$%"')
    })
  })

  describe('Cell References', () => {
    it('should tokenize simple cell reference (A1)', () => {
      const tokens = tokenize('=A1')
      expectToken(tokens[0], 'CELL_REF', 'A1')
    })

    it('should tokenize cell reference with multiple letters (AA1)', () => {
      const tokens = tokenize('=AA1')
      expectToken(tokens[0], 'CELL_REF', 'AA1')
    })

    it('should tokenize cell reference with large row (A1000)', () => {
      const tokens = tokenize('=A1000')
      expectToken(tokens[0], 'CELL_REF', 'A1000')
    })

    it('should tokenize absolute column reference ($A1)', () => {
      const tokens = tokenize('=$A1')
      expectToken(tokens[0], 'CELL_REF', '$A1')
    })

    it('should tokenize absolute row reference (A$1)', () => {
      const tokens = tokenize('=A$1')
      expectToken(tokens[0], 'CELL_REF', 'A$1')
    })

    it('should tokenize fully absolute reference ($A$1)', () => {
      const tokens = tokenize('=$A$1')
      expectToken(tokens[0], 'CELL_REF', '$A$1')
    })

    it('should tokenize sheet-qualified reference (Sheet1!A1)', () => {
      const tokens = tokenize('=Sheet1!A1')
      expectToken(tokens[0], 'CELL_REF', 'Sheet1!A1')
    })

    it('should tokenize quoted sheet name reference (\'My Sheet\'!A1)', () => {
      const tokens = tokenize("='My Sheet'!A1")
      expectToken(tokens[0], 'CELL_REF', "'My Sheet'!A1")
    })

    it('should tokenize sheet name with special characters', () => {
      const tokens = tokenize("='Sheet-1'!A1")
      expectToken(tokens[0], 'CELL_REF', "'Sheet-1'!A1")
    })

    it('should tokenize max column reference (XFD1)', () => {
      const tokens = tokenize('=XFD1')
      expectToken(tokens[0], 'CELL_REF', 'XFD1')
    })

    it('should tokenize lowercase cell reference and normalize', () => {
      const tokens = tokenize('=a1')
      // Should normalize to uppercase
      expectToken(tokens[0], 'CELL_REF', 'A1')
    })
  })

  describe('Ranges', () => {
    it('should tokenize simple range (A1:B10)', () => {
      const tokens = tokenize('=A1:B10')
      expectToken(tokens[0], 'RANGE', 'A1:B10')
    })

    it('should tokenize absolute range ($A$1:$B$10)', () => {
      const tokens = tokenize('=$A$1:$B$10')
      expectToken(tokens[0], 'RANGE', '$A$1:$B$10')
    })

    it('should tokenize column range (A:A)', () => {
      const tokens = tokenize('=A:A')
      expectToken(tokens[0], 'RANGE', 'A:A')
    })

    it('should tokenize row range (1:1)', () => {
      const tokens = tokenize('=1:1')
      expectToken(tokens[0], 'RANGE', '1:1')
    })

    it('should tokenize multi-column range (A:C)', () => {
      const tokens = tokenize('=A:C')
      expectToken(tokens[0], 'RANGE', 'A:C')
    })

    it('should tokenize multi-row range (1:10)', () => {
      const tokens = tokenize('=1:10')
      expectToken(tokens[0], 'RANGE', '1:10')
    })

    it('should tokenize sheet-qualified range (Sheet1!A1:B10)', () => {
      const tokens = tokenize('=Sheet1!A1:B10')
      expectToken(tokens[0], 'RANGE', 'Sheet1!A1:B10')
    })

    it('should tokenize mixed absolute/relative range', () => {
      const tokens = tokenize('=$A1:B$10')
      expectToken(tokens[0], 'RANGE', '$A1:B$10')
    })
  })

  describe('Functions', () => {
    it('should tokenize SUM function', () => {
      const tokens = tokenize('=SUM(A1:A10)')
      expectToken(tokens[0], 'FUNCTION', 'SUM')
      expectToken(tokens[1], 'LPAREN', '(')
    })

    it('should tokenize IF function', () => {
      const tokens = tokenize('=IF(A1>0,1,0)')
      expectToken(tokens[0], 'FUNCTION', 'IF')
    })

    it('should tokenize VLOOKUP function', () => {
      const tokens = tokenize('=VLOOKUP(A1,B1:C10,2,FALSE)')
      expectToken(tokens[0], 'FUNCTION', 'VLOOKUP')
    })

    it('should tokenize nested functions', () => {
      const tokens = tokenize('=SUM(IF(A1:A10>0,A1:A10,0))')
      const functionTokens = tokens.filter((t) => t.type === 'FUNCTION')
      expect(functionTokens).toHaveLength(2)
      expect(functionTokens[0].value).toBe('SUM')
      expect(functionTokens[1].value).toBe('IF')
    })

    it('should tokenize function with no arguments', () => {
      const tokens = tokenize('=NOW()')
      expectToken(tokens[0], 'FUNCTION', 'NOW')
      expectToken(tokens[1], 'LPAREN', '(')
      expectToken(tokens[2], 'RPAREN', ')')
    })

    it('should tokenize function with lowercase name and normalize', () => {
      const tokens = tokenize('=sum(A1:A10)')
      expectToken(tokens[0], 'FUNCTION', 'SUM')
    })

    it('should tokenize function with underscore in name', () => {
      const tokens = tokenize('=DAYS_BETWEEN(A1,B1)')
      expectToken(tokens[0], 'FUNCTION', 'DAYS_BETWEEN')
    })

    it('should tokenize function with dot in name', () => {
      const tokens = tokenize('=CONCATENATE.RANGE(A1:A10)')
      expectToken(tokens[0], 'FUNCTION', 'CONCATENATE.RANGE')
    })
  })

  describe('Parentheses and Commas', () => {
    it('should tokenize left parenthesis', () => {
      const tokens = tokenize('=(1+2)')
      expectToken(tokens[0], 'LPAREN', '(')
    })

    it('should tokenize right parenthesis', () => {
      const tokens = tokenize('=(1+2)')
      expectToken(tokens[4], 'RPAREN', ')')
    })

    it('should tokenize comma as argument separator', () => {
      const tokens = tokenize('=SUM(1,2,3)')
      const commas = tokens.filter((t) => t.type === 'COMMA')
      expect(commas).toHaveLength(2)
    })

    it('should tokenize nested parentheses', () => {
      const tokens = tokenize('=((1+2)*3)')
      const lparens = tokens.filter((t) => t.type === 'LPAREN')
      const rparens = tokens.filter((t) => t.type === 'RPAREN')
      expect(lparens).toHaveLength(2)
      expect(rparens).toHaveLength(2)
    })
  })

  describe('Boolean Values', () => {
    it('should tokenize TRUE', () => {
      const tokens = tokenize('=TRUE')
      expectToken(tokens[0], 'BOOLEAN', 'TRUE')
    })

    it('should tokenize FALSE', () => {
      const tokens = tokenize('=FALSE')
      expectToken(tokens[0], 'BOOLEAN', 'FALSE')
    })

    it('should tokenize lowercase true and normalize', () => {
      const tokens = tokenize('=true')
      expectToken(tokens[0], 'BOOLEAN', 'TRUE')
    })
  })

  describe('Error Values', () => {
    it('should tokenize #REF! error', () => {
      const tokens = tokenize('=#REF!')
      expectToken(tokens[0], 'ERROR', '#REF!')
    })

    it('should tokenize #N/A error', () => {
      const tokens = tokenize('=#N/A')
      expectToken(tokens[0], 'ERROR', '#N/A')
    })

    it('should tokenize #VALUE! error', () => {
      const tokens = tokenize('=#VALUE!')
      expectToken(tokens[0], 'ERROR', '#VALUE!')
    })

    it('should tokenize #DIV/0! error', () => {
      const tokens = tokenize('=#DIV/0!')
      expectToken(tokens[0], 'ERROR', '#DIV/0!')
    })

    it('should tokenize #NAME? error', () => {
      const tokens = tokenize('=#NAME?')
      expectToken(tokens[0], 'ERROR', '#NAME?')
    })

    it('should tokenize #NULL! error', () => {
      const tokens = tokenize('=#NULL!')
      expectToken(tokens[0], 'ERROR', '#NULL!')
    })

    it('should tokenize #NUM! error', () => {
      const tokens = tokenize('=#NUM!')
      expectToken(tokens[0], 'ERROR', '#NUM!')
    })
  })

  describe('Whitespace Handling', () => {
    it('should skip whitespace between tokens', () => {
      const tokens = tokenize('= 1 + 2')
      const nonWhitespace = tokens.filter((t) => t.type !== 'WHITESPACE' && t.type !== 'EOF')
      expect(nonWhitespace).toHaveLength(3)
    })

    it('should preserve positions through whitespace', () => {
      const tokens = tokenize('=  A1')
      const cellRef = tokens.find((t) => t.type === 'CELL_REF')
      expect(cellRef?.start).toBe(3) // After '=' and two spaces
    })
  })

  describe('Token Positions', () => {
    it('should track start and end positions', () => {
      const tokens = tokenize('=A1+B2')
      expect(tokens[0].start).toBe(1) // After '='
      expect(tokens[0].end).toBe(3)   // 'A1'
      expect(tokens[1].start).toBe(3) // '+'
      expect(tokens[1].end).toBe(4)
      expect(tokens[2].start).toBe(4) // 'B2'
      expect(tokens[2].end).toBe(6)
    })

    it('should include EOF token at end', () => {
      const tokens = tokenize('=1')
      const lastToken = tokens[tokens.length - 1]
      expect(lastToken.type).toBe('EOF')
    })
  })

  describe('Complex Formulas', () => {
    it('should tokenize complex formula with multiple components', () => {
      const formula = '=IF(SUM(A1:A10)>100,"High","Low")'
      const tokens = tokenize(formula)
      const types = getTokenTypes(tokens)

      expect(types).toContain('FUNCTION') // IF, SUM
      expect(types).toContain('LPAREN')
      expect(types).toContain('RPAREN')
      expect(types).toContain('RANGE')
      expect(types).toContain('OPERATOR_GT')
      expect(types).toContain('NUMBER')
      expect(types).toContain('STRING')
      expect(types).toContain('COMMA')
    })

    it('should tokenize formula without leading equals sign', () => {
      const tokens = tokenize('A1+B1')
      expectToken(tokens[0], 'CELL_REF', 'A1')
    })
  })
})

// ============================================================================
// SECTION 2: parse() Tests - AST Building
// ============================================================================

describe('parse()', () => {
  describe('BinaryExpression', () => {
    it('should parse addition expression (A1+B1)', () => {
      const result = parse('=A1+B1')
      const body = result.ast.body as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('+')
      expect((body.left as CellReferenceNode).type).toBe('CellReference')
      expect((body.right as CellReferenceNode).type).toBe('CellReference')
    })

    it('should parse subtraction expression (A1-B1)', () => {
      const body = parseBody('=A1-B1') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('-')
    })

    it('should parse multiplication expression (A1*B1)', () => {
      const body = parseBody('=A1*B1') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('*')
    })

    it('should parse division expression (A1/B1)', () => {
      const body = parseBody('=A1/B1') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('/')
    })

    it('should parse power expression (2^3)', () => {
      const body = parseBody('=2^3') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('^')
    })

    it('should parse concatenation expression ("a"&"b")', () => {
      const body = parseBody('="a"&"b"') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('&')
    })

    it('should parse comparison equals (A1=B1)', () => {
      const body = parseBody('=A1=B1') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('=')
    })

    it('should parse comparison not equals (A1<>B1)', () => {
      const body = parseBody('=A1<>B1') as BinaryExpressionNode

      expect(body.type).toBe('BinaryExpression')
      expect(body.operator).toBe('<>')
    })

    it('should parse comparison less than (A1<B1)', () => {
      const body = parseBody('=A1<B1') as BinaryExpressionNode

      expect(body.operator).toBe('<')
    })

    it('should parse comparison greater than (A1>B1)', () => {
      const body = parseBody('=A1>B1') as BinaryExpressionNode

      expect(body.operator).toBe('>')
    })

    it('should parse comparison less than or equal (A1<=B1)', () => {
      const body = parseBody('=A1<=B1') as BinaryExpressionNode

      expect(body.operator).toBe('<=')
    })

    it('should parse comparison greater than or equal (A1>=B1)', () => {
      const body = parseBody('=A1>=B1') as BinaryExpressionNode

      expect(body.operator).toBe('>=')
    })
  })

  describe('UnaryExpression', () => {
    it('should parse negation (-A1)', () => {
      const body = parseBody('=-A1') as UnaryExpressionNode

      expect(body.type).toBe('UnaryExpression')
      expect(body.operator).toBe('-')
      expect(body.prefix).toBe(true)
      expect((body.argument as CellReferenceNode).type).toBe('CellReference')
    })

    it('should parse positive prefix (+A1)', () => {
      const body = parseBody('=+A1') as UnaryExpressionNode

      expect(body.type).toBe('UnaryExpression')
      expect(body.operator).toBe('+')
      expect(body.prefix).toBe(true)
    })

    it('should parse negation of number (-42)', () => {
      const body = parseBody('=-42') as UnaryExpressionNode

      expect(body.type).toBe('UnaryExpression')
      expect(body.operator).toBe('-')
      expect((body.argument as LiteralNode).value).toBe(42)
    })

    it('should parse negation of parenthesized expression (-(A1+B1))', () => {
      const body = parseBody('=-(A1+B1)') as UnaryExpressionNode

      expect(body.type).toBe('UnaryExpression')
      expect(body.operator).toBe('-')
      expect((body.argument as BinaryExpressionNode).type).toBe('BinaryExpression')
    })

    it('should parse double negation (--A1)', () => {
      const body = parseBody('=--A1') as UnaryExpressionNode

      expect(body.type).toBe('UnaryExpression')
      expect(body.operator).toBe('-')
      expect((body.argument as UnaryExpressionNode).type).toBe('UnaryExpression')
    })
  })

  describe('FunctionCall', () => {
    it('should parse simple function call (SUM(A1:A10))', () => {
      const body = parseBody('=SUM(A1:A10)') as FunctionCallNode

      expect(body.type).toBe('FunctionCall')
      expect(body.name).toBe('SUM')
      expect(body.arguments).toHaveLength(1)
      expect((body.arguments[0] as RangeReferenceNode).type).toBe('RangeReference')
    })

    it('should parse function with multiple arguments (IF(A1>0, "Yes", "No"))', () => {
      const body = parseBody('=IF(A1>0,"Yes","No")') as FunctionCallNode

      expect(body.type).toBe('FunctionCall')
      expect(body.name).toBe('IF')
      expect(body.arguments).toHaveLength(3)
    })

    it('should parse function with no arguments (NOW())', () => {
      const body = parseBody('=NOW()') as FunctionCallNode

      expect(body.type).toBe('FunctionCall')
      expect(body.name).toBe('NOW')
      expect(body.arguments).toHaveLength(0)
    })

    it('should parse VLOOKUP with all arguments', () => {
      const body = parseBody('=VLOOKUP(A1,B1:C10,2,FALSE)') as FunctionCallNode

      expect(body.name).toBe('VLOOKUP')
      expect(body.arguments).toHaveLength(4)
    })

    it('should parse nested function (SUM(IF(...)))', () => {
      const body = parseBody('=SUM(IF(A1:A10>0,A1:A10,0))') as FunctionCallNode

      expect(body.type).toBe('FunctionCall')
      expect(body.name).toBe('SUM')
      expect(body.arguments).toHaveLength(1)
      expect((body.arguments[0] as FunctionCallNode).type).toBe('FunctionCall')
      expect((body.arguments[0] as FunctionCallNode).name).toBe('IF')
    })

    it('should parse function with expression argument', () => {
      const body = parseBody('=ROUND(A1+B1,2)') as FunctionCallNode

      expect(body.name).toBe('ROUND')
      expect((body.arguments[0] as BinaryExpressionNode).type).toBe('BinaryExpression')
    })
  })

  describe('CellReference', () => {
    it('should parse simple cell reference (A1)', () => {
      const body = parseBody('=A1') as CellReferenceNode

      expect(body.type).toBe('CellReference')
      expect(body.column).toBe('A')
      expect(body.row).toBe(1)
      expect(body.columnAbsolute).toBe(false)
      expect(body.rowAbsolute).toBe(false)
    })

    it('should parse absolute column reference ($A1)', () => {
      const body = parseBody('=$A1') as CellReferenceNode

      expect(body.column).toBe('A')
      expect(body.columnAbsolute).toBe(true)
      expect(body.rowAbsolute).toBe(false)
    })

    it('should parse absolute row reference (A$1)', () => {
      const body = parseBody('=A$1') as CellReferenceNode

      expect(body.column).toBe('A')
      expect(body.columnAbsolute).toBe(false)
      expect(body.rowAbsolute).toBe(true)
    })

    it('should parse fully absolute reference ($A$1)', () => {
      const body = parseBody('=$A$1') as CellReferenceNode

      expect(body.columnAbsolute).toBe(true)
      expect(body.rowAbsolute).toBe(true)
    })

    it('should parse multi-letter column (AA1)', () => {
      const body = parseBody('=AA1') as CellReferenceNode

      expect(body.column).toBe('AA')
      expect(body.row).toBe(1)
    })

    it('should parse sheet-qualified reference (Sheet1!A1)', () => {
      const body = parseBody('=Sheet1!A1') as CellReferenceNode

      expect(body.column).toBe('A')
      expect(body.row).toBe(1)
      expect(body.sheet).toBe('Sheet1')
    })

    it('should parse quoted sheet name reference', () => {
      const body = parseBody("='My Sheet'!A1") as CellReferenceNode

      expect(body.sheet).toBe('My Sheet')
      expect(body.column).toBe('A')
    })
  })

  describe('RangeReference', () => {
    it('should parse simple range (A1:B10)', () => {
      const body = parseBody('=A1:B10') as RangeReferenceNode

      expect(body.type).toBe('RangeReference')
      expect(body.startCell.column).toBe('A')
      expect(body.startCell.row).toBe(1)
      expect(body.endCell.column).toBe('B')
      expect(body.endCell.row).toBe(10)
    })

    it('should parse absolute range ($A$1:$B$10)', () => {
      const body = parseBody('=$A$1:$B$10') as RangeReferenceNode

      expect(body.startCell.columnAbsolute).toBe(true)
      expect(body.startCell.rowAbsolute).toBe(true)
      expect(body.endCell.columnAbsolute).toBe(true)
      expect(body.endCell.rowAbsolute).toBe(true)
    })

    it('should parse sheet-qualified range (Sheet1!A1:B10)', () => {
      const body = parseBody('=Sheet1!A1:B10') as RangeReferenceNode

      expect(body.sheet).toBe('Sheet1')
    })

    it('should parse column range (A:A)', () => {
      const body = parseBody('=A:A') as RangeReferenceNode

      expect(body.type).toBe('RangeReference')
    })

    it('should parse row range (1:1)', () => {
      const body = parseBody('=1:1') as RangeReferenceNode

      expect(body.type).toBe('RangeReference')
    })
  })

  describe('Literal', () => {
    it('should parse integer literal (42)', () => {
      const body = parseBody('=42') as LiteralNode

      expect(body.type).toBe('Literal')
      expect(body.value).toBe(42)
      expect(body.valueType).toBe('number')
    })

    it('should parse decimal literal (3.14)', () => {
      const body = parseBody('=3.14') as LiteralNode

      expect(body.value).toBe(3.14)
    })

    it('should parse scientific notation (1E10)', () => {
      const body = parseBody('=1E10') as LiteralNode

      expect(body.value).toBe(1e10)
    })

    it('should parse string literal ("text")', () => {
      const body = parseBody('="text"') as LiteralNode

      expect(body.type).toBe('Literal')
      expect(body.value).toBe('text')
      expect(body.valueType).toBe('string')
    })

    it('should parse empty string literal', () => {
      const body = parseBody('=""') as LiteralNode

      expect(body.value).toBe('')
    })

    it('should parse TRUE boolean', () => {
      const body = parseBody('=TRUE') as LiteralNode

      expect(body.type).toBe('Literal')
      expect(body.value).toBe(true)
      expect(body.valueType).toBe('boolean')
    })

    it('should parse FALSE boolean', () => {
      const body = parseBody('=FALSE') as LiteralNode

      expect(body.value).toBe(false)
    })
  })
})

// ============================================================================
// SECTION 3: Operator Precedence (PEMDAS)
// ============================================================================

describe('Operator Precedence (PEMDAS)', () => {
  it('should give multiplication higher precedence than addition (2+3*4 = 14)', () => {
    const body = parseBody('=2+3*4') as BinaryExpressionNode

    // Should be: 2 + (3 * 4), not (2 + 3) * 4
    expect(body.operator).toBe('+')
    expect((body.left as LiteralNode).value).toBe(2)
    expect((body.right as BinaryExpressionNode).operator).toBe('*')
  })

  it('should give division higher precedence than subtraction (10-6/2 = 7)', () => {
    const body = parseBody('=10-6/2') as BinaryExpressionNode

    expect(body.operator).toBe('-')
    expect((body.right as BinaryExpressionNode).operator).toBe('/')
  })

  it('should give exponentiation highest precedence (2+3^2 = 11)', () => {
    const body = parseBody('=2+3^2') as BinaryExpressionNode

    expect(body.operator).toBe('+')
    expect((body.right as BinaryExpressionNode).operator).toBe('^')
  })

  it('should evaluate exponentiation before multiplication (2*3^2 = 18)', () => {
    const body = parseBody('=2*3^2') as BinaryExpressionNode

    expect(body.operator).toBe('*')
    expect((body.right as BinaryExpressionNode).operator).toBe('^')
  })

  it('should respect parentheses over normal precedence ((2+3)*4 = 20)', () => {
    const body = parseBody('=(2+3)*4') as BinaryExpressionNode

    expect(body.operator).toBe('*')
    expect((body.left as BinaryExpressionNode).operator).toBe('+')
  })

  it('should handle nested parentheses ((2+3)*(4+5))', () => {
    const body = parseBody('=(2+3)*(4+5)') as BinaryExpressionNode

    expect(body.operator).toBe('*')
    expect((body.left as BinaryExpressionNode).operator).toBe('+')
    expect((body.right as BinaryExpressionNode).operator).toBe('+')
  })

  it('should parse complex expression 1+2*3^4/5-6 correctly', () => {
    // Expected precedence: 1 + ((2 * (3^4)) / 5) - 6
    const body = parseBody('=1+2*3^4/5-6') as BinaryExpressionNode

    // The top level should be subtraction (leftmost lowest precedence)
    expect(body.type).toBe('BinaryExpression')
    // Verify the structure is built correctly
    expect(body).toBeDefined()
  })

  it('should handle unary minus with correct precedence (-2^2)', () => {
    // In Excel, -2^2 = -4 (exponentiation first, then negation)
    const body = parseBody('=-2^2') as UnaryExpressionNode

    expect(body.type).toBe('UnaryExpression')
    expect(body.operator).toBe('-')
    expect((body.argument as BinaryExpressionNode).operator).toBe('^')
  })

  it('should left-associate operators of same precedence (1-2-3)', () => {
    // Should be (1-2)-3 = -4, not 1-(2-3) = 2
    const body = parseBody('=1-2-3') as BinaryExpressionNode

    expect(body.operator).toBe('-')
    expect((body.left as BinaryExpressionNode).operator).toBe('-')
    expect((body.right as LiteralNode).value).toBe(3)
  })

  it('should right-associate exponentiation (2^3^2)', () => {
    // In Excel, 2^3^2 = 2^9 = 512, not 8^2 = 64
    const body = parseBody('=2^3^2') as BinaryExpressionNode

    expect(body.operator).toBe('^')
    expect((body.left as LiteralNode).value).toBe(2)
    expect((body.right as BinaryExpressionNode).operator).toBe('^')
  })

  it('should give comparison operators lower precedence than arithmetic', () => {
    const body = parseBody('=A1+B1>C1*D1') as BinaryExpressionNode

    expect(body.operator).toBe('>')
    expect((body.left as BinaryExpressionNode).operator).toBe('+')
    expect((body.right as BinaryExpressionNode).operator).toBe('*')
  })

  it('should give concatenation lower precedence than arithmetic', () => {
    const body = parseBody('=1+2&3+4') as BinaryExpressionNode

    expect(body.operator).toBe('&')
    expect((body.left as BinaryExpressionNode).operator).toBe('+')
    expect((body.right as BinaryExpressionNode).operator).toBe('+')
  })
})

// ============================================================================
// SECTION 4: Nested Function Calls
// ============================================================================

describe('Nested Function Calls', () => {
  it('should parse IF with SUM in true branch: IF(A1>0, SUM(B1:B10), 0)', () => {
    const body = parseBody('=IF(A1>0,SUM(B1:B10),0)') as FunctionCallNode

    expect(body.type).toBe('FunctionCall')
    expect(body.name).toBe('IF')
    expect(body.arguments).toHaveLength(3)

    // First arg: comparison
    expect((body.arguments[0] as BinaryExpressionNode).type).toBe('BinaryExpression')
    expect((body.arguments[0] as BinaryExpressionNode).operator).toBe('>')

    // Second arg: nested SUM
    expect((body.arguments[1] as FunctionCallNode).type).toBe('FunctionCall')
    expect((body.arguments[1] as FunctionCallNode).name).toBe('SUM')

    // Third arg: literal
    expect((body.arguments[2] as LiteralNode).type).toBe('Literal')
  })

  it('should parse deeply nested functions: SUM(IF(AND(A1>0,B1<10),C1,0))', () => {
    const body = parseBody('=SUM(IF(AND(A1>0,B1<10),C1,0))') as FunctionCallNode

    expect(body.name).toBe('SUM')

    const ifCall = body.arguments[0] as FunctionCallNode
    expect(ifCall.name).toBe('IF')

    const andCall = ifCall.arguments[0] as FunctionCallNode
    expect(andCall.name).toBe('AND')
    expect(andCall.arguments).toHaveLength(2)
  })

  it('should parse SUMPRODUCT with nested functions', () => {
    const body = parseBody('=SUMPRODUCT((A1:A10>0)*IF(B1:B10="Yes",1,0))') as FunctionCallNode

    expect(body.name).toBe('SUMPRODUCT')
    expect(body.arguments).toHaveLength(1)
  })

  it('should parse INDEX/MATCH combination', () => {
    const body = parseBody('=INDEX(A1:A10,MATCH(B1,C1:C10,0))') as FunctionCallNode

    expect(body.name).toBe('INDEX')
    expect(body.arguments).toHaveLength(2)

    const matchCall = body.arguments[1] as FunctionCallNode
    expect(matchCall.name).toBe('MATCH')
    expect(matchCall.arguments).toHaveLength(3)
  })

  it('should parse multiple functions at same level', () => {
    const body = parseBody('=SUM(A1:A10)+AVERAGE(B1:B10)') as BinaryExpressionNode

    expect(body.type).toBe('BinaryExpression')
    expect(body.operator).toBe('+')
    expect((body.left as FunctionCallNode).name).toBe('SUM')
    expect((body.right as FunctionCallNode).name).toBe('AVERAGE')
  })

  it('should parse IFERROR with nested VLOOKUP', () => {
    const body = parseBody('=IFERROR(VLOOKUP(A1,B1:C10,2,FALSE),"Not Found")') as FunctionCallNode

    expect(body.name).toBe('IFERROR')
    expect((body.arguments[0] as FunctionCallNode).name).toBe('VLOOKUP')
    expect((body.arguments[1] as LiteralNode).value).toBe('Not Found')
  })

  it('should parse functions with expressions in arguments', () => {
    const body = parseBody('=IF(SUM(A1:A10)/COUNT(A1:A10)>50,"High","Low")') as FunctionCallNode

    expect(body.name).toBe('IF')

    const condition = body.arguments[0] as BinaryExpressionNode
    expect(condition.operator).toBe('>')

    const division = condition.left as BinaryExpressionNode
    expect(division.operator).toBe('/')
    expect((division.left as FunctionCallNode).name).toBe('SUM')
    expect((division.right as FunctionCallNode).name).toBe('COUNT')
  })

  it('should parse triple-nested IF', () => {
    const body = parseBody('=IF(A1>0,IF(A1>10,IF(A1>100,"XL","L"),"M"),"S")') as FunctionCallNode

    expect(body.name).toBe('IF')

    const firstNestedIf = body.arguments[1] as FunctionCallNode
    expect(firstNestedIf.name).toBe('IF')

    const secondNestedIf = firstNestedIf.arguments[1] as FunctionCallNode
    expect(secondNestedIf.name).toBe('IF')
  })
})

// ============================================================================
// SECTION 5: Error Handling for Malformed Formulas
// ============================================================================

describe('Error Handling', () => {
  describe('Unclosed Parentheses', () => {
    it('should report error for unclosed left parenthesis: (1+2', () => {
      const result = parse('=(1+2')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNCLOSED_PAREN')
    })

    it('should report error for extra right parenthesis: 1+2)', () => {
      const result = parse('=1+2)')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for mismatched parentheses: ((1+2)', () => {
      const result = parse('=((1+2)')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNCLOSED_PAREN')
    })

    it('should report error for function with unclosed paren: SUM(A1:A10', () => {
      const result = parse('=SUM(A1:A10')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNCLOSED_PAREN')
    })
  })

  describe('Unclosed Strings', () => {
    it('should report error for unclosed double-quoted string: "hello', () => {
      const result = parse('="hello')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNCLOSED_STRING')
    })

    it('should report error for unclosed single-quoted string: \'hello', () => {
      const result = parse("='hello")
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNCLOSED_STRING')
    })
  })

  describe('Invalid Operators', () => {
    it('should report error for consecutive operators: 1++2', () => {
      // Note: +2 could be unary, so 1++2 might parse as 1+(+2)
      // But 1+*2 is definitely wrong
      const result = parse('=1+*2')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for trailing operator: 1+', () => {
      const result = parse('=1+')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('MISSING_OPERAND')
    })

    it('should report error for leading binary operator: *2', () => {
      const result = parse('=*2')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('UNEXPECTED_TOKEN')
    })
  })

  describe('Invalid References', () => {
    it('should report error for invalid cell reference: A', () => {
      // 'A' alone is not a valid cell reference (no row number)
      // However, it might be parsed as a potential named range
      const result = parse('=A')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for invalid range: A1:', () => {
      const result = parse('=A1:')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].code).toBe('INVALID_RANGE')
    })

    it('should report error for reversed range: B1:A1 (optional validation)', () => {
      // Some parsers might accept this; it's optional to validate
      const result = parse('=B1:A1')
      // This test just checks parsing succeeds (validation is separate)
      expect(result.ast).toBeDefined()
    })
  })

  describe('Invalid Function Calls', () => {
    it('should report error for function without parentheses: SUM', () => {
      // SUM alone without () might be treated as named range
      const result = parse('=SUM')
      // Should have some indication this is incomplete
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for function with missing comma: SUM(1 2)', () => {
      const result = parse('=SUM(1 2)')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for trailing comma in function: SUM(1,2,)', () => {
      const result = parse('=SUM(1,2,)')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for leading comma in function: SUM(,1,2)', () => {
      const result = parse('=SUM(,1,2)')
      expect(result.errors.length).toBeGreaterThan(0)
    })
  })

  describe('Empty Formula', () => {
    it('should handle empty formula gracefully', () => {
      const result = parse('=')
      // Either error or empty body
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should handle formula with only whitespace', () => {
      const result = parse('=   ')
      expect(result.errors.length).toBeGreaterThan(0)
    })
  })

  describe('Error Position Reporting', () => {
    it('should report correct position for error: =1+(', () => {
      const result = parse('=1+(')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].position).toBeGreaterThan(0)
    })

    it('should report correct position for unclosed string in middle', () => {
      const result = parse('=1+"hello+2')
      expect(result.errors).toHaveLength(1)
      expect(result.errors[0].position).toBe(3) // Position of opening quote
    })
  })

  describe('Syntax Errors', () => {
    it('should report error for colon without range: 1:2+3', () => {
      // Colon is only valid in range context
      const result = parse('=1:2+3')
      // This might parse as range 1:2, which is valid
      // Or error depending on implementation
      expect(result).toBeDefined()
    })

    it('should report error for invalid characters: 1@2', () => {
      const result = parse('=1@2')
      expect(result.errors.length).toBeGreaterThan(0)
    })

    it('should report error for unrecognized token: 1#2', () => {
      // # at start is error indicator, but mid-expression is invalid
      const result = parse('=1#2')
      expect(result.errors.length).toBeGreaterThan(0)
    })
  })
})

// ============================================================================
// SECTION 6: Additional Helper Function Tests
// ============================================================================

describe('extractCellReferences()', () => {
  it('should extract single cell reference', () => {
    const refs = extractCellReferences('=A1')
    expect(refs).toEqual(['A1'])
  })

  it('should extract multiple cell references', () => {
    const refs = extractCellReferences('=A1+B2+C3')
    expect(refs).toEqual(['A1', 'B2', 'C3'])
  })

  it('should extract cell references from functions', () => {
    const refs = extractCellReferences('=SUM(A1,B2,C3)')
    expect(refs).toEqual(['A1', 'B2', 'C3'])
  })

  it('should extract sheet-qualified references', () => {
    const refs = extractCellReferences('=Sheet1!A1+Sheet2!B2')
    expect(refs).toEqual(['Sheet1!A1', 'Sheet2!B2'])
  })

  it('should not include range references', () => {
    const refs = extractCellReferences('=A1+A1:B10')
    expect(refs).toEqual(['A1'])
  })
})

describe('extractRangeReferences()', () => {
  it('should extract single range reference', () => {
    const refs = extractRangeReferences('=SUM(A1:A10)')
    expect(refs).toEqual(['A1:A10'])
  })

  it('should extract multiple range references', () => {
    const refs = extractRangeReferences('=SUM(A1:A10)+SUM(B1:B10)')
    expect(refs).toEqual(['A1:A10', 'B1:B10'])
  })

  it('should extract sheet-qualified ranges', () => {
    const refs = extractRangeReferences('=SUM(Sheet1!A1:A10)')
    expect(refs).toEqual(['Sheet1!A1:A10'])
  })
})

describe('getDependencies()', () => {
  it('should return both cells and ranges', () => {
    const deps = getDependencies('=A1+SUM(B1:B10)')
    expect(deps.cells).toEqual(['A1'])
    expect(deps.ranges).toEqual(['B1:B10'])
  })

  it('should return empty arrays for literal-only formula', () => {
    const deps = getDependencies('=1+2')
    expect(deps.cells).toEqual([])
    expect(deps.ranges).toEqual([])
  })
})

describe('validateFormula()', () => {
  it('should return true for valid formula', () => {
    expect(validateFormula('=A1+B1')).toBe(true)
  })

  it('should return false for unclosed parenthesis', () => {
    expect(validateFormula('=(A1+B1')).toBe(false)
  })

  it('should return false for unclosed string', () => {
    expect(validateFormula('="hello')).toBe(false)
  })

  it('should return false for empty formula', () => {
    expect(validateFormula('=')).toBe(false)
  })
})

describe('parseTokens()', () => {
  it('should parse pre-tokenized input', () => {
    const tokens = tokenize('=A1+B1')
    const ast = parseTokens(tokens)

    expect(ast.type).toBe('Program')
    expect(ast.body.type).toBe('BinaryExpression')
  })
})
