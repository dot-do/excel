/**
 * Excel Formula Parser
 *
 * Stub implementation for TDD - all functions throw NotImplementedError
 * to ensure tests are in RED state.
 */

import type {
  Token,
  TokenType,
  ASTNode,
  ProgramNode,
  ParseResult,
  ParserOptions,
  ParseError,
  BinaryExpressionNode,
  UnaryExpressionNode,
  FunctionCallNode,
  CellReferenceNode,
  RangeReferenceNode,
  LiteralNode,
  ErrorValueNode,
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
 *
 * @example
 * tokenize('=A1+B1') // Returns tokens for A1, +, B1
 * tokenize('=SUM(A1:A10)') // Returns tokens for SUM(, A1:A10, )
 */
export function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let pos = 0

  // Strip optional leading '='
  if (formula.startsWith('=')) {
    pos = 1
  }

  while (pos < formula.length) {
    const start = pos
    const char = formula[pos]

    // Skip whitespace
    if (/\s/.test(char)) {
      pos++
      continue
    }

    // Error values: #REF!, #N/A, #VALUE!, #DIV/0!, #NAME?, #NULL!, #NUM!
    if (char === '#') {
      const errorPatterns = [
        '#DIV/0!',
        '#VALUE!',
        '#REF!',
        '#NAME?',
        '#NULL!',
        '#NUM!',
        '#N/A',
      ]
      let matched = false
      for (const pattern of errorPatterns) {
        if (formula.slice(pos).toUpperCase().startsWith(pattern)) {
          tokens.push({
            type: 'ERROR',
            value: pattern,
            start,
            end: pos + pattern.length,
          })
          pos += pattern.length
          matched = true
          break
        }
      }
      if (!matched) {
        // Unknown error - just skip the character
        pos++
      }
      continue
    }

    // Strings (double-quoted)
    if (char === '"') {
      let value = '"'
      pos++
      while (pos < formula.length) {
        if (formula[pos] === '"') {
          value += '"'
          pos++
          // Check for escaped quote ("")
          if (pos < formula.length && formula[pos] === '"') {
            value += '"'
            pos++
          } else {
            break
          }
        } else {
          value += formula[pos]
          pos++
        }
      }
      tokens.push({
        type: 'STRING',
        value,
        start,
        end: pos,
      })
      continue
    }

    // Strings (single-quoted) - could be sheet name or string literal
    if (char === "'") {
      // Check if it's a quoted sheet reference: 'Sheet Name'!A1
      const sheetMatch = formula.slice(pos).match(/^'([^']+)'!/)
      if (sheetMatch) {
        // This is a quoted sheet name, could be followed by a cell ref or range
        const sheetPrefix = sheetMatch[0]
        const afterSheet = formula.slice(pos + sheetPrefix.length)

        // Try to match a cell reference or range after the sheet prefix
        const cellOrRangeMatch = afterSheet.match(/^(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?|\$?[A-Za-z]+:\$?[A-Za-z]+|\d+:\d+)/)
        if (cellOrRangeMatch) {
          const refPart = cellOrRangeMatch[0].toUpperCase()
          const isRange = refPart.includes(':')
          tokens.push({
            type: isRange ? 'RANGE' : 'CELL_REF',
            value: sheetMatch[1].includes(' ') || sheetMatch[1].includes('-')
              ? "'" + sheetMatch[1] + "'!" + refPart
              : sheetMatch[1] + "!" + refPart,
            start,
            end: pos + sheetPrefix.length + cellOrRangeMatch[0].length,
          })
          pos += sheetPrefix.length + cellOrRangeMatch[0].length
          continue
        }
      }

      // Otherwise it's a string literal
      let value = "'"
      pos++
      while (pos < formula.length) {
        if (formula[pos] === "'") {
          value += "'"
          pos++
          // Check for escaped quote ('')
          if (pos < formula.length && formula[pos] === "'") {
            value += "'"
            pos++
          } else {
            break
          }
        } else {
          value += formula[pos]
          pos++
        }
      }
      tokens.push({
        type: 'STRING',
        value,
        start,
        end: pos,
      })
      continue
    }

    // Multi-character operators: <=, >=, <>
    if (char === '<') {
      if (pos + 1 < formula.length) {
        if (formula[pos + 1] === '=') {
          tokens.push({ type: 'OPERATOR_LTE', value: '<=', start, end: pos + 2 })
          pos += 2
          continue
        }
        if (formula[pos + 1] === '>') {
          tokens.push({ type: 'OPERATOR_NE', value: '<>', start, end: pos + 2 })
          pos += 2
          continue
        }
      }
      tokens.push({ type: 'OPERATOR_LT', value: '<', start, end: pos + 1 })
      pos++
      continue
    }

    if (char === '>') {
      if (pos + 1 < formula.length && formula[pos + 1] === '=') {
        tokens.push({ type: 'OPERATOR_GTE', value: '>=', start, end: pos + 2 })
        pos += 2
        continue
      }
      tokens.push({ type: 'OPERATOR_GT', value: '>', start, end: pos + 1 })
      pos++
      continue
    }

    // Single-character operators
    if (char === '+') {
      tokens.push({ type: 'OPERATOR_ADD', value: '+', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '-') {
      tokens.push({ type: 'OPERATOR_SUB', value: '-', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '*') {
      tokens.push({ type: 'OPERATOR_MUL', value: '*', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '/') {
      tokens.push({ type: 'OPERATOR_DIV', value: '/', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '^') {
      tokens.push({ type: 'OPERATOR_POW', value: '^', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '&') {
      tokens.push({ type: 'OPERATOR_CONCAT', value: '&', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === '=') {
      tokens.push({ type: 'OPERATOR_EQ', value: '=', start, end: pos + 1 })
      pos++
      continue
    }

    // Punctuation
    if (char === '(') {
      tokens.push({ type: 'LPAREN', value: '(', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === ')') {
      tokens.push({ type: 'RPAREN', value: ')', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === ',') {
      tokens.push({ type: 'COMMA', value: ',', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === ';') {
      tokens.push({ type: 'SEMICOLON', value: ';', start, end: pos + 1 })
      pos++
      continue
    }
    if (char === ':') {
      tokens.push({ type: 'COLON', value: ':', start, end: pos + 1 })
      pos++
      continue
    }

    // Numbers (including leading decimal and scientific notation)
    if (/\d/.test(char) || (char === '.' && pos + 1 < formula.length && /\d/.test(formula[pos + 1]))) {
      // Check if this could be a row range like 1:1 or 1:10
      const rowRangeMatch = formula.slice(pos).match(/^(\d+:\d+)(?![A-Za-z\d])/)
      if (rowRangeMatch) {
        tokens.push({
          type: 'RANGE',
          value: rowRangeMatch[1],
          start,
          end: pos + rowRangeMatch[1].length,
        })
        pos += rowRangeMatch[1].length
        continue
      }

      // Parse number: digits, optional decimal, optional exponent
      let numStr = ''
      // Digits before decimal point
      while (pos < formula.length && /\d/.test(formula[pos])) {
        numStr += formula[pos]
        pos++
      }
      // Decimal part
      if (pos < formula.length && formula[pos] === '.') {
        numStr += '.'
        pos++
        while (pos < formula.length && /\d/.test(formula[pos])) {
          numStr += formula[pos]
          pos++
        }
      }
      // Exponent part
      if (pos < formula.length && (formula[pos] === 'e' || formula[pos] === 'E')) {
        numStr += formula[pos]
        pos++
        if (pos < formula.length && (formula[pos] === '+' || formula[pos] === '-')) {
          numStr += formula[pos]
          pos++
        }
        while (pos < formula.length && /\d/.test(formula[pos])) {
          numStr += formula[pos]
          pos++
        }
      }
      tokens.push({
        type: 'NUMBER',
        value: numStr,
        start,
        end: pos,
      })
      continue
    }

    // Identifiers: functions, cell references, booleans, or sheet-qualified refs
    if (/[A-Za-z$_]/.test(char)) {
      // Try to match a sheet-qualified reference: Sheet1!A1 or Sheet1!A1:B10
      const sheetQualifiedMatch = formula.slice(pos).match(/^([A-Za-z_][A-Za-z0-9_]*)!(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?|\$?[A-Za-z]+:\$?[A-Za-z]+|\d+:\d+)/)
      if (sheetQualifiedMatch) {
        const fullMatch = sheetQualifiedMatch[0]
        const refPart = sheetQualifiedMatch[2].toUpperCase()
        const sheetName = sheetQualifiedMatch[1]
        const isRange = refPart.includes(':')
        tokens.push({
          type: isRange ? 'RANGE' : 'CELL_REF',
          value: sheetName + '!' + refPart,
          start,
          end: pos + fullMatch.length,
        })
        pos += fullMatch.length
        continue
      }

      // Try to match a cell reference or range with possible $ markers
      // Cell reference: $?[A-Z]+$?\d+
      // Range: $?[A-Z]+$?\d+:$?[A-Z]+$?\d+ OR $?[A-Z]+:$?[A-Z]+ (column range)
      const rangeOrCellMatch = formula.slice(pos).match(/^(\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+|\$?[A-Za-z]+:\$?[A-Za-z]+|\$?[A-Za-z]+\$?\d+)/)
      if (rangeOrCellMatch) {
        const matchVal = rangeOrCellMatch[1]
        // Normalize to uppercase
        const normalized = matchVal.toUpperCase()

        if (normalized.includes(':')) {
          // It's a range
          tokens.push({
            type: 'RANGE',
            value: normalized,
            start,
            end: pos + matchVal.length,
          })
          pos += matchVal.length
          continue
        }

        // It's a potential cell reference, but need to check if it's a function, boolean, or just identifier
        // Check for function (followed by '(')
        const afterMatch = formula.slice(pos + matchVal.length)

        // First check if it's a valid cell reference pattern
        const cellRefPattern = /^\$?([A-Za-z]+)\$?(\d+)$/
        const cellMatch = matchVal.match(cellRefPattern)

        if (cellMatch) {
          // Could be a cell reference or a function/boolean
          const upperVal = matchVal.toUpperCase()

          // Check for boolean
          if (upperVal === 'TRUE' || upperVal === 'FALSE') {
            // But also check if followed by '(' - then it's a function
            if (afterMatch.trimStart().startsWith('(')) {
              tokens.push({
                type: 'FUNCTION',
                value: upperVal,
                start,
                end: pos + matchVal.length,
              })
              pos += matchVal.length
              continue
            }
            tokens.push({
              type: 'BOOLEAN',
              value: upperVal,
              start,
              end: pos + matchVal.length,
            })
            pos += matchVal.length
            continue
          }

          // Check if followed by '(' - then it's a function
          if (afterMatch.trimStart().startsWith('(')) {
            tokens.push({
              type: 'FUNCTION',
              value: normalized,
              start,
              end: pos + matchVal.length,
            })
            pos += matchVal.length
            continue
          }

          // It's a cell reference
          tokens.push({
            type: 'CELL_REF',
            value: normalized,
            start,
            end: pos + matchVal.length,
          })
          pos += matchVal.length
          continue
        }
      }

      // Try to match just column range like A:C
      const columnRangeMatch = formula.slice(pos).match(/^(\$?[A-Za-z]+:\$?[A-Za-z]+)(?![A-Za-z0-9])/)
      if (columnRangeMatch) {
        tokens.push({
          type: 'RANGE',
          value: columnRangeMatch[1].toUpperCase(),
          start,
          end: pos + columnRangeMatch[1].length,
        })
        pos += columnRangeMatch[1].length
        continue
      }

      // Match identifier (function name, boolean, or named range)
      let ident = ''
      while (pos < formula.length && /[A-Za-z0-9_.]/.test(formula[pos])) {
        ident += formula[pos]
        pos++
      }

      const upperIdent = ident.toUpperCase()

      // Check for boolean
      if (upperIdent === 'TRUE' || upperIdent === 'FALSE') {
        // Check if followed by '(' - function named TRUE/FALSE
        const afterIdent = formula.slice(pos).trimStart()
        if (afterIdent.startsWith('(')) {
          tokens.push({
            type: 'FUNCTION',
            value: upperIdent,
            start,
            end: start + ident.length,
          })
          continue
        }
        tokens.push({
          type: 'BOOLEAN',
          value: upperIdent,
          start,
          end: start + ident.length,
        })
        continue
      }

      // Check if it's a function (followed by '(')
      const afterIdent = formula.slice(pos).trimStart()
      if (afterIdent.startsWith('(')) {
        tokens.push({
          type: 'FUNCTION',
          value: upperIdent,
          start,
          end: start + ident.length,
        })
        continue
      }

      // Otherwise treat as named range (or error - but tokenizer just produces tokens)
      tokens.push({
        type: 'NAMED_RANGE',
        value: ident,
        start,
        end: start + ident.length,
      })
      continue
    }

    // If we get here, it's an unrecognized character - skip it
    pos++
  }

  // Add EOF token
  tokens.push({
    type: 'EOF',
    value: '',
    start: pos,
    end: pos,
  })

  return tokens
}

/**
 * Parse a formula string into an Abstract Syntax Tree.
 *
 * @param formula - The formula string to parse (with or without leading '=')
 * @param options - Optional parser configuration
 * @returns ParseResult containing the AST and any errors
 *
 * @example
 * parse('=A1+B1') // Returns AST with BinaryExpression
 * parse('=SUM(A1:A10)') // Returns AST with FunctionCall
 */
export function parse(formula: string, _options?: ParserOptions): ParseResult {
  const tokens = tokenize(formula)
  const errors: ParseError[] = []

  // Check for unclosed strings in tokens (tokenizer might have caught partial string)
  for (const token of tokens) {
    if (token.type === 'STRING') {
      const val = token.value
      if ((val.startsWith('"') && !val.endsWith('"')) ||
          (val.startsWith("'") && !val.endsWith("'"))) {
        errors.push({
          code: 'UNCLOSED_STRING',
          message: 'Unclosed string literal',
          position: token.start,
          length: token.value.length,
        })
      }
    }
  }

  const parser = new Parser(tokens)
  const ast = parser.parse()

  // Only add parser errors if we didn't already find an unclosed string
  // (to avoid duplicate errors)
  const hasUnclosedStringError = errors.some(e => e.code === 'UNCLOSED_STRING')
  for (const err of parser.getErrors()) {
    // Skip duplicate unclosed string errors
    if (hasUnclosedStringError && err.code === 'UNCLOSED_STRING') {
      continue
    }
    errors.push(err)
  }

  return {
    ast,
    errors,
    formula,
  }
}

/**
 * Parser class for building AST from tokens
 */
class Parser {
  private tokens: Token[]
  private pos: number = 0
  private errors: ParseError[] = []

  constructor(tokens: Token[]) {
    this.tokens = tokens
  }

  private current(): Token {
    return this.tokens[this.pos] || { type: 'EOF', value: '', start: 0, end: 0 }
  }

  private peek(offset: number = 0): Token {
    return this.tokens[this.pos + offset] || { type: 'EOF', value: '', start: 0, end: 0 }
  }

  private advance(): Token {
    const token = this.current()
    if (token.type !== 'EOF') {
      this.pos++
    }
    return token
  }

  private match(...types: TokenType[]): boolean {
    return types.includes(this.current().type)
  }

  private expect(type: TokenType, errorCode: ParseError['code'], message: string): Token {
    if (this.current().type === type) {
      return this.advance()
    }
    this.addError(errorCode, message, this.current().start, 1)
    return this.current()
  }

  private addError(code: ParseError['code'], message: string, position: number, length: number) {
    this.errors.push({ code, message, position, length })
  }

  getErrors(): ParseError[] {
    return this.errors
  }

  /**
   * Parse the full program
   * program -> expression EOF
   */
  parse(): ProgramNode {
    if (this.current().type === 'EOF') {
      this.addError('UNEXPECTED_EOF', 'Empty formula', 0, 1)
      return {
        type: 'Program',
        body: { type: 'Literal', value: null, valueType: 'null', start: 0, end: 0 },
        start: 0,
        end: 0,
      }
    }

    const startPos = this.current().start
    const body = this.parseExpression()
    const endPos = this.peek(-1)?.end || this.current().end

    // Check for unexpected tokens after expression
    // Only report if we don't already have errors (to avoid cascading error messages)
    if (this.current().type !== 'EOF' && this.errors.length === 0) {
      const token = this.current()
      if (token.type === 'RPAREN') {
        this.addError('UNEXPECTED_TOKEN', 'Unexpected closing parenthesis', token.start, 1)
      } else if (token.type === 'COLON') {
        // A colon after a cell reference indicates an incomplete range (e.g., A1:)
        this.addError('INVALID_RANGE', 'Incomplete range reference', token.start, 1)
      } else {
        this.addError('UNEXPECTED_TOKEN', `Unexpected token: ${token.value}`, token.start, token.value.length)
      }
    }

    return {
      type: 'Program',
      body,
      start: startPos,
      end: endPos,
    }
  }

  /**
   * Parse expression (lowest precedence)
   * expression -> comparison
   */
  private parseExpression(): ASTNode {
    return this.parseComparison()
  }

  /**
   * Parse comparison operators (=, <>, <, >, <=, >=) - lowest precedence
   */
  private parseComparison(): ASTNode {
    let left = this.parseConcatenation()

    while (this.match('OPERATOR_EQ', 'OPERATOR_NE', 'OPERATOR_LT', 'OPERATOR_GT', 'OPERATOR_LTE', 'OPERATOR_GTE')) {
      const operator = this.advance()
      const right = this.parseConcatenation()
      left = {
        type: 'BinaryExpression',
        operator: operator.value,
        left,
        right,
        start: left.start,
        end: right.end,
      } as BinaryExpressionNode
    }

    return left
  }

  /**
   * Parse concatenation (&)
   */
  private parseConcatenation(): ASTNode {
    let left = this.parseAdditive()

    while (this.match('OPERATOR_CONCAT')) {
      const operator = this.advance()
      const right = this.parseAdditive()
      left = {
        type: 'BinaryExpression',
        operator: operator.value,
        left,
        right,
        start: left.start,
        end: right.end,
      } as BinaryExpressionNode
    }

    return left
  }

  /**
   * Parse addition and subtraction (+, -)
   */
  private parseAdditive(): ASTNode {
    let left = this.parseMultiplicative()

    while (this.match('OPERATOR_ADD', 'OPERATOR_SUB')) {
      const operator = this.advance()
      const right = this.parseMultiplicative()
      left = {
        type: 'BinaryExpression',
        operator: operator.value,
        left,
        right,
        start: left.start,
        end: right.end,
      } as BinaryExpressionNode
    }

    return left
  }

  /**
   * Parse multiplication and division (*, /)
   */
  private parseMultiplicative(): ASTNode {
    let left = this.parseUnary()

    while (this.match('OPERATOR_MUL', 'OPERATOR_DIV')) {
      const operator = this.advance()
      const right = this.parseUnary()
      left = {
        type: 'BinaryExpression',
        operator: operator.value,
        left,
        right,
        start: left.start,
        end: right.end,
      } as BinaryExpressionNode
    }

    return left
  }

  /**
   * Parse unary operators (-, +)
   * Unary has lower precedence than power, so -2^2 = -(2^2)
   */
  private parseUnary(): ASTNode {
    if (this.match('OPERATOR_SUB', 'OPERATOR_ADD')) {
      const operator = this.advance()
      // For -2^2 = -(2^2), we need unary to capture the whole power expression
      const argument = this.parseUnary() // Allow chained unary: --A1
      return {
        type: 'UnaryExpression',
        operator: operator.value,
        argument,
        prefix: true,
        start: operator.start,
        end: argument.end,
      } as UnaryExpressionNode
    }

    return this.parsePower()
  }

  /**
   * Parse power/exponentiation (^) - right associative
   * In Excel, -2^2 = -4 (exponentiation first, then negation)
   * Power has higher precedence than unary
   */
  private parsePower(): ASTNode {
    const left = this.parsePrimary()

    if (this.match('OPERATOR_POW')) {
      const operator = this.advance()
      const right = this.parseUnary() // Right side can have unary, e.g., 2^-3
      return {
        type: 'BinaryExpression',
        operator: operator.value,
        left,
        right,
        start: left.start,
        end: right.end,
      } as BinaryExpressionNode
    }

    return left
  }

  /**
   * Parse primary expressions (atoms)
   */
  private parsePrimary(): ASTNode {
    const token = this.current()

    switch (token.type) {
      case 'NUMBER':
        return this.parseNumber()

      case 'STRING':
        return this.parseString()

      case 'BOOLEAN':
        return this.parseBoolean()

      case 'ERROR':
        return this.parseError()

      case 'CELL_REF':
        return this.parseCellReference()

      case 'RANGE':
        return this.parseRangeReference()

      case 'FUNCTION':
        return this.parseFunctionCall()

      case 'LPAREN':
        return this.parseParenthesized()

      case 'NAMED_RANGE':
        // Named ranges without () are potentially invalid in our context
        // Check if this is a function without parens
        this.advance()
        this.addError('UNEXPECTED_TOKEN', `Unexpected identifier: ${token.value}`, token.start, token.value.length)
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode

      case 'OPERATOR_MUL':
      case 'OPERATOR_DIV':
      case 'OPERATOR_POW':
      case 'OPERATOR_CONCAT':
      case 'OPERATOR_EQ':
      case 'OPERATOR_NE':
      case 'OPERATOR_LT':
      case 'OPERATOR_GT':
      case 'OPERATOR_LTE':
      case 'OPERATOR_GTE':
        this.addError('UNEXPECTED_TOKEN', `Unexpected operator: ${token.value}`, token.start, token.value.length)
        this.advance()
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode

      case 'COLON':
        this.addError('INVALID_RANGE', 'Incomplete range reference', token.start, 1)
        this.advance()
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode

      case 'COMMA':
        this.addError('UNEXPECTED_TOKEN', 'Unexpected comma', token.start, 1)
        this.advance()
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode

      case 'EOF':
        this.addError('MISSING_OPERAND', 'Unexpected end of formula', token.start, 1)
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode

      default:
        this.addError('UNEXPECTED_TOKEN', `Unexpected token: ${token.value}`, token.start, token.value.length || 1)
        this.advance()
        return {
          type: 'Literal',
          value: null,
          valueType: 'null',
          start: token.start,
          end: token.end,
        } as LiteralNode
    }
  }

  private parseNumber(): LiteralNode {
    const token = this.advance()
    return {
      type: 'Literal',
      value: parseFloat(token.value),
      valueType: 'number',
      start: token.start,
      end: token.end,
    }
  }

  private parseString(): LiteralNode {
    const token = this.advance()
    // Remove surrounding quotes and handle escaped quotes
    let value = token.value
    const isUnclosed = (value.startsWith('"') && !value.endsWith('"')) ||
                       (value.startsWith("'") && !value.endsWith("'"))
    if (value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1).replace(/""/g, '"')
    } else if (value.startsWith("'") && value.endsWith("'")) {
      value = value.slice(1, -1).replace(/''/g, "'")
    } else if (isUnclosed) {
      // Handle unclosed string - just remove opening quote
      value = value.slice(1)
    }
    // Note: unclosed string error is detected in parse() function
    return {
      type: 'Literal',
      value,
      valueType: 'string',
      start: token.start,
      end: token.end,
    }
  }

  private parseBoolean(): LiteralNode {
    const token = this.advance()
    return {
      type: 'Literal',
      value: token.value === 'TRUE',
      valueType: 'boolean',
      start: token.start,
      end: token.end,
    }
  }

  private parseError(): ErrorValueNode {
    const token = this.advance()
    return {
      type: 'ErrorValue',
      error: token.value,
      start: token.start,
      end: token.end,
    }
  }

  private parseCellReference(): CellReferenceNode {
    const token = this.advance()
    return parseCellRef(token.value, token.start, token.end)
  }

  private parseRangeReference(): RangeReferenceNode {
    const token = this.advance()
    return parseRangeRef(token.value, token.start, token.end)
  }

  private parseFunctionCall(): FunctionCallNode {
    const nameToken = this.advance()
    const startPos = nameToken.start

    this.expect('LPAREN', 'SYNTAX_ERROR', 'Expected ( after function name')

    const args: ASTNode[] = []

    // Handle empty argument list
    if (this.match('RPAREN')) {
      const endToken = this.advance()
      return {
        type: 'FunctionCall',
        name: nameToken.value,
        arguments: args,
        start: startPos,
        end: endToken.end,
      }
    }

    // Check for leading comma
    if (this.match('COMMA')) {
      this.addError('UNEXPECTED_TOKEN', 'Unexpected leading comma in function arguments', this.current().start, 1)
    }

    // Parse first argument
    args.push(this.parseExpression())

    // Parse remaining arguments
    while (this.match('COMMA')) {
      this.advance()
      // Check for trailing comma
      if (this.match('RPAREN')) {
        this.addError('UNEXPECTED_TOKEN', 'Unexpected trailing comma in function arguments', this.peek(-1).start, 1)
        break
      }
      // Check for consecutive commas
      if (this.match('COMMA')) {
        this.addError('UNEXPECTED_TOKEN', 'Missing argument between commas', this.current().start, 1)
      }
      args.push(this.parseExpression())
    }

    // Check for missing comma between arguments
    if (!this.match('RPAREN') && !this.match('EOF')) {
      this.addError('SYNTAX_ERROR', 'Expected comma or ) in function arguments', this.current().start, 1)
    }

    if (this.match('RPAREN')) {
      const endToken = this.advance()
      return {
        type: 'FunctionCall',
        name: nameToken.value,
        arguments: args,
        start: startPos,
        end: endToken.end,
      }
    }

    this.addError('UNCLOSED_PAREN', 'Missing closing parenthesis for function', startPos, nameToken.value.length)
    return {
      type: 'FunctionCall',
      name: nameToken.value,
      arguments: args,
      start: startPos,
      end: this.current().end,
    }
  }

  private parseParenthesized(): ASTNode {
    const startToken = this.advance() // consume '('
    const errorCountBefore = this.errors.length
    const expr = this.parseExpression()

    if (this.match('RPAREN')) {
      this.advance()
    } else {
      // Only report unclosed paren if we didn't already encounter errors inside
      // (to avoid cascading error messages like "missing operand" + "unclosed paren")
      if (this.errors.length === errorCountBefore) {
        this.addError('UNCLOSED_PAREN', 'Missing closing parenthesis', startToken.start, 1)
      }
    }

    return expr
  }
}

/**
 * Helper to parse a cell reference string into a CellReferenceNode
 */
function parseCellRef(ref: string, start: number, end: number): CellReferenceNode {
  let sheet: string | undefined
  let cellPart = ref

  // Check for sheet prefix
  if (ref.includes('!')) {
    const [sheetPart, rest] = ref.split('!')
    sheet = sheetPart.replace(/^'|'$/g, '') // Remove surrounding quotes
    cellPart = rest
  }

  // Parse cell reference: $?[A-Z]+$?[0-9]+
  const match = cellPart.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i)
  if (!match) {
    return {
      type: 'CellReference',
      column: 'A',
      row: 1,
      columnAbsolute: false,
      rowAbsolute: false,
      sheet,
      start,
      end,
    }
  }

  const [, colAbs, col, rowAbs, row] = match
  return {
    type: 'CellReference',
    column: col.toUpperCase(),
    row: parseInt(row, 10),
    columnAbsolute: colAbs === '$',
    rowAbsolute: rowAbs === '$',
    sheet,
    start,
    end,
  }
}

/**
 * Helper to parse a range reference string into a RangeReferenceNode
 */
function parseRangeRef(ref: string, startPos: number, endPos: number): RangeReferenceNode {
  let sheet: string | undefined
  let rangePart = ref

  // Check for sheet prefix
  if (ref.includes('!')) {
    const bangIndex = ref.indexOf('!')
    const sheetPart = ref.slice(0, bangIndex)
    sheet = sheetPart.replace(/^'|'$/g, '') // Remove surrounding quotes
    rangePart = ref.slice(bangIndex + 1)
  }

  // Split by colon
  const [startRef, endRef] = rangePart.split(':')

  // Handle column ranges (A:C) and row ranges (1:10)
  const colRangeMatch = rangePart.match(/^(\$?[A-Z]+):(\$?[A-Z]+)$/i)
  const rowRangeMatch = rangePart.match(/^(\d+):(\d+)$/)

  if (colRangeMatch) {
    // Column range like A:C
    const startColAbs = colRangeMatch[1].startsWith('$')
    const startCol = colRangeMatch[1].replace('$', '').toUpperCase()
    const endColAbs = colRangeMatch[2].startsWith('$')
    const endCol = colRangeMatch[2].replace('$', '').toUpperCase()

    const rangeNode: RangeReferenceNode = {
      type: 'RangeReference',
      startCell: {
        type: 'CellReference',
        column: startCol,
        row: 1,
        columnAbsolute: startColAbs,
        rowAbsolute: false,
        start: startPos,
        end: endPos,
      },
      endCell: {
        type: 'CellReference',
        column: endCol,
        row: 1048576, // Max Excel row
        columnAbsolute: endColAbs,
        rowAbsolute: false,
        start: startPos,
        end: endPos,
      },
      sheet,
      start: startPos,
      end: endPos,
    }
    return rangeNode
  }

  if (rowRangeMatch) {
    // Row range like 1:10
    const rangeNode: RangeReferenceNode = {
      type: 'RangeReference',
      startCell: {
        type: 'CellReference',
        column: 'A',
        row: parseInt(rowRangeMatch[1], 10),
        columnAbsolute: false,
        rowAbsolute: false,
        start: startPos,
        end: endPos,
      },
      endCell: {
        type: 'CellReference',
        column: 'XFD', // Max Excel column
        row: parseInt(rowRangeMatch[2], 10),
        columnAbsolute: false,
        rowAbsolute: false,
        start: startPos,
        end: endPos,
      },
      sheet,
      start: startPos,
      end: endPos,
    }
    return rangeNode
  }

  // Standard cell range like A1:B10
  const startNode = parseCellRef(startRef, startPos, endPos)
  const endNode = parseCellRef(endRef, startPos, endPos)

  const rangeNode: RangeReferenceNode = {
    type: 'RangeReference',
    startCell: startNode,
    endCell: endNode,
    sheet,
    start: startPos,
    end: endPos,
  }
  return rangeNode
}

/**
 * Parse tokens into an Abstract Syntax Tree.
 *
 * @param tokens - Array of tokens from tokenize()
 * @param options - Optional parser configuration
 * @returns The root AST node
 */
export function parseTokens(tokens: Token[], _options?: ParserOptions): ProgramNode {
  const parser = new Parser(tokens)
  return parser.parse()
}

/**
 * Validate a formula string without fully parsing it.
 * Performs quick syntax validation.
 *
 * @param formula - The formula string to validate
 * @returns true if formula appears valid, false otherwise
 */
export function validateFormula(formula: string): boolean {
  const result = parse(formula)
  return result.errors.length === 0
}

/**
 * Extract all cell references from a formula.
 *
 * @param formula - The formula string to analyze
 * @returns Array of cell reference strings (e.g., ['A1', 'B2', 'Sheet1!C3'])
 */
export function extractCellReferences(formula: string): string[] {
  const tokens = tokenize(formula)
  return tokens
    .filter((t) => t.type === 'CELL_REF')
    .map((t) => t.value)
}

/**
 * Extract all range references from a formula.
 *
 * @param formula - The formula string to analyze
 * @returns Array of range reference strings (e.g., ['A1:B10', 'Sheet1!C1:C100'])
 */
export function extractRangeReferences(formula: string): string[] {
  const tokens = tokenize(formula)
  return tokens
    .filter((t) => t.type === 'RANGE')
    .map((t) => t.value)
}

/**
 * Get all dependencies (cells and ranges) that a formula references.
 *
 * @param formula - The formula string to analyze
 * @returns Object with cells and ranges arrays
 */
export function getDependencies(formula: string): {
  cells: string[]
  ranges: string[]
} {
  return {
    cells: extractCellReferences(formula),
    ranges: extractRangeReferences(formula),
  }
}
